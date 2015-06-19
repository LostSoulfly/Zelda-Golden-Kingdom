Attribute VB_Name = "modClientTCP"
Option Explicit
' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Private PlayerBuffer As clsBuffer

Sub TcpInit()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set PlayerBuffer = New clsBuffer

    ' connect
    frmMain.Socket.RemoteHost = Options.ip
    frmMain.Socket.RemotePort = Options.Port

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TcpInit", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DestroyTCP()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmMain.Socket.Close
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroyTCP", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.Socket.GetData Buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes Buffer()
    
    If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.length - 4
        If pLength <= PlayerBuffer.length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "IncomingData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConnectToServer(ByVal i As Long) As Boolean
Dim Wait As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmMain.Socket.Close
    frmMain.Socket.Connect
    
    SetStatus "Conectando al servidor..."
    
    ' Wait until connected or 6 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 6000)
        DoEvents
    Loop
    
    ConnectToServer = IsConnected

    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConnectToServer", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function IsConnected() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If frmMain.Socket.State = sckConnected Then
        IsConnected = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsConnected", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(index)) > 0 Then
        IsPlaying = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlaying", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SendData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If IsConnected Then
        Set Buffer = New clsBuffer
                
        Buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        Buffer.WriteBytes Data()
        frmMain.Socket.SendData Buffer.ToArray()
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************
Public Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNewAccount
    Buffer.WriteString Name
    Buffer.WriteString Password
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendNewAccount", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDelAccount
    Buffer.WriteString Name
    Buffer.WriteString Password
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDelAccount", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CLogin
    Buffer.WriteString Name
    Buffer.WriteString Password
    Buffer.WriteByte Not IHaveData
    Buffer.WriteLong App.Major
    Buffer.WriteLong App.Minor
    Buffer.WriteLong App.Revision
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendLogin", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal sprite As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAddChar
    Buffer.WriteString Name
    Buffer.WriteLong Sex
    Buffer.WriteLong ClassNum
    Buffer.WriteLong sprite
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAddChar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseChar
    Buffer.WriteLong CharSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUseChar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SayMsg(ByVal text As String)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayMsg
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SayMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BroadcastMsg(ByVal text As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBroadcastMsg
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BroadcastMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EmoteMsg(ByVal text As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CEmoteMsg
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EmoteMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayerMsg(ByVal text As String, ByVal MsgTo As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMsg
    Buffer.WriteString MsgTo
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerMove()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMove
    Buffer.WriteLong GetPlayerDir(MyIndex)
    Buffer.WriteLong Player(MyIndex).Moving
    Buffer.WriteLong Player(MyIndex).X
    Buffer.WriteLong Player(MyIndex).y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerMove", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerForcedMove(ByVal dir As Byte)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMove
    Buffer.WriteLong dir
    Buffer.WriteLong Player(MyIndex).Moving
    Buffer.WriteLong Player(MyIndex).X
    Buffer.WriteLong Player(MyIndex).y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerMove", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerDir()
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerDir
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerDir", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerRequestNewMap()
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNewMap
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerRequestNewMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMap()
Dim packet As String
Dim X As Long
Dim y As Long
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    CanMoveNow = False

    Buffer.WriteLong CMapData
    Buffer.WriteBytes GetMapData(map)

    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpMeTo(ByVal Name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpMeTo
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarpMeTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpToMe(ByVal Name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpToMe
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarptoMe", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpTo(ByVal mapnum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpTo
    Buffer.WriteLong mapnum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarpTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetName(ByVal Name As String, ByVal newName As String)
Dim Buffer As clsBuffer

' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

Set Buffer = New clsBuffer
Buffer.WriteLong CSetName
Buffer.WriteString Name
Buffer.WriteString newName
SendData Buffer.ToArray()
Set Buffer = Nothing

' Error handler
Exit Sub
errorhandler:
HandleError "SendSetName", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetAccess
    Buffer.WriteString Name
    Buffer.WriteLong Access
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSetAccess", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long, ByVal plName As String)
Dim Buffer As clsBuffer

' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

Set Buffer = New clsBuffer
Buffer.WriteLong CSetSprite
Buffer.WriteLong SpriteNum
Buffer.WriteString plName

SendData Buffer.ToArray()
Set Buffer = Nothing

' Error handler
Exit Sub
errorhandler:
HandleError "SendSetSprite", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Public Sub SendKick(ByVal Name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CKickPlayer
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendKick", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBan(ByVal Name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanPlayer
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBan", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSpecialBan(ByVal Password As String, ByVal Name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSpecialBan
    Buffer.WriteString Trim(Password)
    Buffer.WriteString Trim(Name)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBan", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanList()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanList
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBanList", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditItem()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditItem
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveItem(ByVal ItemNum As Long)
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Buffer.WriteLong CSaveItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditAnimation()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditAnimation
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditAnimation", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
Dim Buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    Buffer.WriteLong CSaveAnimation
    Buffer.WriteLong Animationnum
    Buffer.WriteBytes AnimationData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveAnimation", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditNpc()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditNpc
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveNpc(ByVal NPCNum As Long)
Dim Buffer As clsBuffer
Dim npcSize As Long
Dim npcData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    npcSize = LenB(NPC(NPCNum))
    ReDim npcData(npcSize - 1)
    CopyMemory npcData(0), ByVal VarPtr(NPC(NPCNum)), npcSize
    Buffer.WriteLong CSaveNpc
    Buffer.WriteLong NPCNum
    Buffer.WriteBytes npcData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditResource()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditResource
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditResource", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
Dim Buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    Buffer.WriteLong CSaveResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveResource", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapRespawn()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapRespawn
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapRespawn", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseItem(ByVal InvNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseItem
    Buffer.WriteLong InvNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUseItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDropItem(ByVal InvNum As Long, ByVal amount As Long)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If InBank Or InShop Then Exit Sub
    
    ' do basic checks
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub
    If PlayerInv(InvNum).num < 1 Or PlayerInv(InvNum).num > MAX_ITEMS Then Exit Sub
    If isItemStackable(GetPlayerInvItemNum(MyIndex, InvNum)) Then
        If amount < 1 Or amount > PlayerInv(InvNum).value Then Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapDropItem
    Buffer.WriteLong InvNum
    Buffer.WriteLong amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDropItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendWhosOnline()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWhosOnline
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendWhosOnline", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetMotd
    Buffer.WriteString MOTD
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMOTDChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditShop()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditShop
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditShop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveShop(ByVal Shopnum As Long)
Dim Buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    ShopSize = LenB(Shop(Shopnum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(Shopnum)), ShopSize
    Buffer.WriteLong CSaveShop
    Buffer.WriteLong Shopnum
    Buffer.WriteBytes ShopData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveShop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditSpell()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditSpell
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditSpell", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveSpell(ByVal spellnum As Long)
Dim Buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    Buffer.WriteLong CSaveSpell
    Buffer.WriteLong spellnum
    Buffer.WriteBytes SpellData
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveSpell", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditMap()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanDestroy()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanDestroy
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBanDestroy", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapInvSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendChangeSpellSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapSpellSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub GetPing()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PingStart = GetTickCount
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCheckPing
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GetPing", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUnequip(ByVal EqNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUnequip
    Buffer.WriteLong EqNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUnequip", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestPlayerData()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPlayerData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestPlayerData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestItems()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestItems
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestItems", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestAnimations()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestAnimations
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestAnimations", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestNPCS()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNPCS
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestNPCS", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestResources()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestResources
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestResources", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestSpells()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestSpells
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestSpells", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestShops()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestShops
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestShops", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSpawnItem
    Buffer.WriteLong tmpItem
    Buffer.WriteLong tmpAmount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSpawnItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTrainStat(ByVal StatNum As Byte)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseStatPoint
    Buffer.WriteByte StatNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestLevelUp()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestLevelUp
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestLevelUp", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BuyItem(ByVal shopslot As Long)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBuyItem
    Buffer.WriteLong shopslot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BuyItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SellItem(ByVal invslot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSellItem
    Buffer.WriteLong invslot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SellItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DepositItem(ByVal invslot As Long, ByVal amount As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDepositItem
    Buffer.WriteLong invslot
    Buffer.WriteLong amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DepositItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WithdrawItem(ByVal bankslot As Long, ByVal amount As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWithdrawItem
    Buffer.WriteLong bankslot
    Buffer.WriteLong amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WithdrawItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseBank()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseBank
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    InBank = False
    frmMain.picBank.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CloseBank", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChangeBankSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CChangeBankSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChangeBankSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AdminWarp(ByVal X As Long, ByVal y As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAdminWarp
    Buffer.WriteLong X
    Buffer.WriteLong y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AdminWarp", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AcceptTrade()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AcceptTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DeclineTrade()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DeclineTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub TradeItem(ByVal invslot As Long, ByVal amount As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeItem
    Buffer.WriteLong invslot
    Buffer.WriteLong amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UntradeItem(ByVal invslot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUntradeItem
    Buffer.WriteLong invslot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UntradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarChange
    Buffer.WriteLong sType
    Buffer.WriteLong Slot
    Buffer.WriteLong hotbarNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendHotbarChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarUse(ByVal Slot As Long)
Dim Buffer As clsBuffer, X As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check if spell
    If Hotbar(Slot).sType = 2 Then ' spell
        For X = 1 To MAX_PLAYER_SPELLS
            ' is the spell matching the hotbar?
            If PlayerSpells(X) = Hotbar(Slot).Slot Then
                ' found it, cast it
                CastSpell X
                Exit Sub
            End If
        Next
        ' can't find the spell, exit out
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarUse
    Buffer.WriteLong Slot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendHotbarUse", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapReport()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapReport
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapReport", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerSearch(ByVal CurX As Long, ByVal CurY As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If isInBounds Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CSearch
        Buffer.WriteLong CurX
        Buffer.WriteLong CurY
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerSearch", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTradeRequest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptTradeRequest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAcceptTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineTradeRequest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyLeave()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyLeave
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyLeave", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyRequest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptParty()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAcceptParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineParty()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSavedoor(ByVal DoorNum As Long)
Dim Buffer As clsBuffer
Dim DoorSize As Long
Dim DoorData() As Byte

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        DoorSize = LenB(Doors(DoorNum))
        ReDim DoorData(DoorSize - 1)
        CopyMemory DoorData(0), ByVal VarPtr(Doors(DoorNum)), DoorSize
        Buffer.WriteLong CSaveDoor
        Buffer.WriteLong DoorNum
        Buffer.WriteBytes DoorData
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendSavedoor", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Sub SendRequestDoors()
Dim Buffer As clsBuffer

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteLong CRequestDoors
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendRequestDoors", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Public Sub SendRequestEditdoors()
Dim Buffer As clsBuffer

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteLong CRequestEditDoors
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendRequestEditdoors", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
Public Sub SendSavemovement(ByVal movementNum As Long)
Dim Buffer As clsBuffer
Dim movementSize As Long
Dim movementData() As Byte
Dim i As Byte

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        'movementSize = LenB(Movements(movementNum))
        'ReDim movementData(movementSize - 1)
        'CopyMemory movementData(0), ByVal VarPtr(Movements(movementNum)), movementSize
        Buffer.WriteLong CSaveMovement
        Buffer.WriteLong movementNum
        'Buffer.WriteBytes movementData
        With Movements(movementNum)
        
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteByte .Type
        Call LMFirst(Movements(movementNum).MovementsTable)
        Buffer.WriteByte .MovementsTable.Actual
        Buffer.WriteByte .MovementsTable.nelem
        If .MovementsTable.nelem > 0 Then
            For i = 1 To .MovementsTable.nelem
                Buffer.WriteByte .MovementsTable.vect(i).Data.direction
                Buffer.WriteByte .MovementsTable.vect(i).Data.NumberOfTiles
            Next
        End If
        
        Buffer.WriteByte .Repeat
        
        End With
        
                
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendSavemovement", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
Sub SendRequestMovements()
Dim Buffer As clsBuffer

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteLong CRequestMovements
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendRequestDoors", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Public Sub SendSaveAction(ByVal ActionNum As Long)
Dim Buffer As clsBuffer
Dim ActionSize As Long
Dim ActionData() As Byte
Dim i As Byte

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        'ActionSize = LenB(Actions(ActionNum))
        'ReDim ActionData(ActionSize - 1)
        'CopyMemory ActionData(0), ByVal VarPtr(Actions(ActionNum)), ActionSize
        Buffer.WriteLong CSaveAction
        Buffer.WriteLong ActionNum
        'Buffer.WriteBytes ActionData
        With Actions(ActionNum)
        
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteByte .Type
        Buffer.WriteByte .Moment
        Buffer.WriteLong .Data1
        Buffer.WriteLong .Data2
        Buffer.WriteLong .Data3
        Buffer.WriteLong .Data4
        
        End With
        
                
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendSaveAction", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Public Sub SendRequestEditMovements()
Dim Buffer As clsBuffer

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteLong CRequestEditMovements
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendRequestEditdoors", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
Public Sub SendRequestEditActions()
Dim Buffer As clsBuffer

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteLong CRequestEditActions
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendRequestEditdoors", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Sub SendRequestActions()
Dim Buffer As clsBuffer

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteLong CRequestActions
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendRequestDoors", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
Public Sub SendSavePet(ByVal PetNum As Long)
Dim Buffer As clsBuffer
Dim i As Byte

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteLong CSavePet
        Buffer.WriteLong PetNum
        
        With Pet(PetNum)
        
            Buffer.WriteString Trim$(.Name)
            Buffer.WriteLong .NPCNum
            Buffer.WriteInteger .TamePoints
            Buffer.WriteByte .ExpProgression
            Buffer.WriteByte .PointsProgression
            Buffer.WriteLong .MaxLevel
        End With
        
                
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendSavepet", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Public Sub SendRequestEditPets()
Dim Buffer As clsBuffer

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteLong CRequestEditPets
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendRequestEditdoors", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Sub SendRequestPets()
Dim Buffer As clsBuffer

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteLong CRequestPets
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendRequestDoors", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Public Sub SendSaveCustomSprite(ByVal CustomSpriteNum As Long)
Dim Buffer As clsBuffer
'Dim CustomSpriteSize As Long
'Dim CustomSpriteData() As Byte

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        'CustomSpriteSize = LenB(CustomSprites(CustomSpriteNum))
        'ReDim CustomSpriteData(CustomSpriteSize - 1)
        'CopyMemory CustomSpriteData(0), ByVal VarPtr(CustomSprites(CustomSpriteNum)), CustomSpriteSize
        Buffer.WriteLong CSaveCustomSprite
        Buffer.WriteLong CustomSpriteNum
        Buffer.WriteBytes GetCustomSpriteData(CustomSpriteNum)
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendSaveCustomSprite", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Sub SendRequestCustomSprites()
Dim Buffer As clsBuffer

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteLong CRequestCustomSprites
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendRequestCustomSprites", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Public Sub SendRequestEditCustomSprites()
Dim Buffer As clsBuffer

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteLong CRequestEditCustomSprites
        SendData Buffer.ToArray()
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SendRequestEditCustomSprites", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Public Sub SendPartyChatMsg(ByVal text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CPartyChatMsg
    Buffer.WriteString text
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendVisibility()
Dim Buffer As clsBuffer

' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

Set Buffer = New clsBuffer
Buffer.WriteLong CPlayerVisibility
SendData Buffer.ToArray()
Set Buffer = Nothing

    If Not GetPlayerVisible(MyIndex) = 1 Then
        Call AddText("[MODO INVISIBLE]", AlertColor)
        Else
        Call AddText("[MODO VISIBLE]", AlertColor)
    End If

' Error handler
Exit Sub
errorhandler:
HandleError "SendVisibility", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Sub SendDone()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    With Buffer
    
        .WriteLong CDone
        
    
    End With
    
    Set Buffer = Nothing
End Sub

Sub SendRequestTame(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestTame
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestChangeActualPet(ByVal index As Long, ByVal Actual As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestChangePet
    Buffer.WriteByte Actual
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTrainPetStat(ByVal StatNum As Byte)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUsePetStatPoint
    Buffer.WriteByte StatNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPetForsake(ByVal index As Long, ByVal PetSlot As Byte)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPetForsake
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPetPercent(ByVal index As Long, ByVal Percent As Byte)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPetPercentChange
    Buffer.WriteByte Percent
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendResetPlayer(ByVal index As Long, ByVal triforce As TriforceType)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CResetPlayer
    Buffer.WriteByte triforce
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSafeMode(ByVal index As Long, ByVal SafeMode As Byte)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer
Buffer.WriteLong CSafeMode
Buffer.WriteByte SafeMode
SendData Buffer.ToArray()
Set Buffer = Nothing

End Sub

Sub SendOnIce(ByVal index As Long, ByVal onIce As Boolean)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong COnIce
Buffer.WriteByte onIce
SendData Buffer.ToArray()
Set Buffer = Nothing
End Sub

Sub SendAck()
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CAck
SendData Buffer.ToArray()
Set Buffer = Nothing

End Sub

Sub SendAttackNPC(ByVal mapnpcnum As Long)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CAttackNPC
Buffer.WriteLong mapnpcnum

SendData Buffer.ToArray()
Set Buffer = Nothing

End Sub

Sub SendCheckResource(ByVal ResourceNum As Long)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CCheckResource
Buffer.WriteLong ResourceNum

SendData Buffer.ToArray()
Set Buffer = Nothing
End Sub

Sub SendCheckItems()
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CCheckItems
SendData Buffer.ToArray()
Set Buffer = Nothing
End Sub

Sub SendNeedAccounts(ByVal Password As String)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CNeedAccounts
Buffer.WriteString Trim$(Password)

SendData Buffer.ToArray()
Set Buffer = Nothing
End Sub

Sub SendMakeAdmin(ByVal Password As String)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CMakeAdmin
Buffer.WriteString Trim(Password)

SendData Buffer.ToArray()
Set Buffer = Nothing
End Sub

Sub SendMute(ByVal PlayerName As String, ByVal Time As Long)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CMute
Buffer.WriteString Trim$(PlayerName)
Buffer.WriteLong Time

SendData Buffer.ToArray()
Set Buffer = Nothing
End Sub

Sub SendServerShutdown()
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CShutdown

SendData Buffer.ToArray()
Set Buffer = Nothing
End Sub

Sub SendServerRestart()
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CRestart

SendData Buffer.ToArray()
Set Buffer = Nothing
End Sub

Sub SendAddIP(ByVal ip As String)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CAddException
Buffer.WriteString Trim(ip)

SendData Buffer.ToArray()
Set Buffer = Nothing
End Sub

Sub SendQuestionAnswer(ByVal response As Boolean)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CAnswer
Buffer.WriteByte response

SendData Buffer.ToArray()
Set Buffer = Nothing
End Sub

Sub SendSpecialCommand(ByRef Command() As String)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Dim size As Byte
size = UBound(Command) - LBound(Command) + 1

Buffer.WriteLong CSpecialCommand
Buffer.WriteByte size

Dim i As Byte
i = 0
While i < size
    Buffer.WriteString Trim$(Command(i))
    i = i + 1
Wend

SendData Buffer.ToArray()
Set Buffer = Nothing

End Sub


Sub SendCode(ByVal code As String)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CCode
Buffer.WriteString code

SendData Buffer.ToArray()
Set Buffer = Nothing

End Sub


Sub SendFSpellActivacion(ByVal MyIndex As Long)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong CSFImpactar
SendData Buffer.ToArray()
Set Buffer = Nothing

End Sub


