Attribute VB_Name = "modHubTCP"
Public isHubConnected As Boolean
Public hubLastHeard As Long
Public useHubServer As Boolean

Private hubBuffer As clsBuffer

Public Enum HubPackets
    HHello = 1
    HServerInfo
    HShutdown
    HLog
    HGlobalMsg
    HCommand
    HMSG_COUNT
End Enum

Public Enum CommandsType
    Classes = 1
    Maps
    spells
    Shops
    npcs
    Items
    Resources
    Animations
    Language
    SOptions
    SPets
    Weather
End Enum

Public Const MAX_SERVERS As Long = 10
Public HandleDataHub(HMSG_COUNT) As Long

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Sub InitHubMessages()

HandleDataHub(HHello) = GetAddress(AddressOf Hello)
HandleDataHub(HShutdown) = GetAddress(AddressOf HandleShutdownFromHub)
HandleDataHub(HGlobalMsg) = GetAddress(AddressOf HandleForwardGlobalMsg)
HandleDataHub(HCommand) = GetAddress(AddressOf HandleServerCommand)

End Sub

Sub HandleDataHubSub(ByRef Data() As Byte)
'Dim buffer As clsBuffer
Dim MsgType As Long

    'Set buffer = New clsBuffer
    'buffer.WriteBytes Data()

    'MsgType = buffer.ReadLong
    MsgType = ReadHandleDataType(Data)

    If MsgType < 0 Then
        Exit Sub
    End If

    If MsgType >= HMSG_COUNT Then
        Exit Sub
    End If

    'CallWindowProc HandleDataSub(MsgType), index, buffer.ReadBytes(buffer.length), 0, 0
    CallWindowProc HandleDataHub(MsgType), index, Data, 0, 0
'Set buffer = Nothing
End Sub

Sub SendDataHub(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte
If Not useHubServer Then Exit Sub
    If frmServer.hubSocket.state = sckConnected Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        frmServer.hubSocket.SendData Buffer.ToArray()
        
        PacketsSent = PacketsSent + 1
        BytesSent = BytesSent + (UBound(TempData) - LBound(TempData)) + 1
        
    End If
End Sub

Public Sub CheckHubConnection()

If hubLastHeard = 0 Then hubLastHeard = GetRealTickCount

    With frmServer.hubSocket
        If GetRealTickCount > hubLastHeard + 5000 Then
            If .state <> sckConnected Then
                If isHubConnected Then
                    TextAdd "Hub server disconnected."
                    GlobalMsg "Hub Server disconnected.", Green, False
                End If
                
                isHubConnected = False
                'so let's try to connect again..
                .Close
                .Connect
                hubLastHeard = GetRealTickCount
                
                Do While GetRealTickCount < (hubLastHeard + 500)
                    
                    DoEvents
                    Sleep 10
                    
                    If .state = sckConnected Then
                        'we've connected!
                        isHubConnected = True
                        GlobalMsg "Hub server connected.", Green, False
                        Exit Do
                    End If
                    
                Loop
            Else
                isHubConnected = True
            End If
        End If
    End With
    

End Sub

Public Sub HubIncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

If hubBuffer Is Nothing Then Set hubBuffer = New clsBuffer

    frmServer.hubSocket.GetData Buffer, vbUnicode, DataLength
    
    hubBuffer.WriteBytes Buffer()
    
    If hubBuffer.length >= 4 Then pLength = hubBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= hubBuffer.length - 4
        If pLength <= hubBuffer.length - 4 Then
            hubBuffer.ReadLong
            HandleDataHubSub hubBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If hubBuffer.length >= 4 Then pLength = hubBuffer.ReadLong(False)
    Loop
    hubBuffer.Trim

End Sub

Public Sub SendServerInfo()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong HServerInfo
    Buffer.WriteLong TotalPlayersOnline
    Buffer.WriteLong MAX_PLAYERS
    'server name
    Buffer.WriteString IIf(Len(SERVER_NAME) = 0, frmServer.Socket(0).LocalPort, SERVER_NAME)
    Buffer.WriteLong StartTick
    Buffer.WriteLong frmServer.Socket(0).LocalPort
    
    SendDataHub Buffer.ToArray

    Set Buffer = Nothing
End Sub

Public Sub SendHubLog(Text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong HLog
    Buffer.WriteString Text
    
    SendDataHub Buffer.ToArray

    Set Buffer = Nothing
End Sub

Public Sub ForwardGlobalMsg(ByVal msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong HGlobalMsg
    Buffer.WriteString msg
    
    SendDataHub Buffer.ToArray

    Set Buffer = Nothing

End Sub

Private Sub Hello(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next

    TextAdd "Received HELLO from Hub."
    SendServerInfo
    
End Sub

Public Sub SendHubCommand(CommandNum As Long, Data As String)
If Not useHubServer Then Exit Sub
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong HCommand
Buffer.WriteLong CommandNum
Buffer.WriteString Data

SendDataHub Buffer.ToArray

Set Buffer = Nothing

End Sub

Private Sub HandleServerCommand(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim Buffer As New clsBuffer
    Set Buffer = New clsBuffer
    Dim Command As Long
    Dim sData As String
    Buffer.WriteBytes Data()
    Command = Buffer.ReadLong
    sData = Buffer.ReadString
    
        GlobalMsg "Syncing with Hub server..", Green, False
        'DoEvents
        
    Select Case Command
        Case Is = CommandsType.Classes
            frmServer.cmdReloadClasses.Value = True
        Case Is = CommandsType.Maps
            Dim mapnum As Long
            If IsNumeric(sData) Then mapnum = val(sData)
            'frmServer.cmdReloadMaps.Value = True

            Call ClearMap(mapnum)
            Call LoadMap(mapnum)
            
            Call ClearMapWaitingNPCS(mapnum)
            Call SendMapNpcsToMap(mapnum)
            Call SpawnMapNpcs(mapnum)
        
            ' Clear out it all
            For i = 1 To MAX_MAP_ITEMS
                Call SpawnItemSlot(i, 0, 0, mapnum, MapItem(mapnum, i).X, MapItem(mapnum, i).Y)
                Call ClearMapItem(i, mapnum)
            Next
        
            ' Respawn
            Call SpawnMapItems(mapnum)

            Call ClearTempTile(mapnum)
            Call InitTempTile(mapnum)
            Call CacheResources(mapnum)
            Call InitTempMap(mapnum)
        
            ' Refresh map for everyone online
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = mapnum Then
                    AddMapPlayer i, mapnum
                    Call PlayerSpawn(i, mapnum, GetPlayerX(i), GetPlayerY(i))
                End If
            Next i

        Case Is = CommandsType.spells
            'frmServer.CmdReloadSpells.Value = True
            Dim spellnum As Long
            If IsNumeric(sData) Then
                spellnum = val(sData)
                Call LoadSpell(spellnum)
            End If
            
        Case Is = CommandsType.Shops
            'frmServer.cmdReloadShops.Value = True
            Dim shopnum As Long
            If IsNumeric(sData) Then
                shopnum = val(sData)
                Call LoadShop(shopnum)
            End If
            
        Case Is = CommandsType.npcs
            'frmServer.cmdReloadNPCs.Value = True
            Dim npcnum As Long
            If IsNumeric(sData) Then
                npcnum = val(sData)
                Call LoadNpc(npcnum)
            End If
            
        Case Is = CommandsType.Items
            'frmServer.cmdReloadItems.Value = True
            
            Dim ItemNum As Long
            If IsNumeric(sData) Then
                ItemNum = val(sData)
                Call LoadItem(ItemNum)
            End If
        Case Is = CommandsType.Resources
            'frmServer.cmdReloadResources.Value = True
            
            Dim ResourceNum As Long
            If IsNumeric(sData) Then
                ResourceNum = val(sData)
                Call LoadResource(ResourceNum)
            End If
            
        Case Is = CommandsType.Animations
            frmServer.cmdReloadAnimations.Value = True
        Case Is = CommandsType.Language
            frmServer.cmdReloadLang.Value = True
        Case Is = CommandsType.SOptions
            Call TextAdd("Options reloaded.")
            LoadOptions
        Case Is = CommandsType.Weather
            RainOn = val(sData)
            SendWeathertoAll False
            LastWeatherUpdate = GetRealTickCount + WeatherTime
            
        Case Is = CommandsType.SPets
            Dim PetNum As Long
                If IsNumeric(sData) Then
                PetNum = val(sData)
                Call LoadPet(PetNum)
            End If
        
    Case Else
        TextAdd "Unknown command received!"
    
    End Select


    Set Buffer = Nothing

End Sub

Private Sub HandleShutdownFromHub(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next

    TextAdd "ShutdownFromHub received."
        
    isShuttingDown = True
    frmServer.cmdShutDown.Caption = "Cancel"

End Sub

Private Sub HandleForwardGlobalMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As New clsBuffer
    Set Buffer = New clsBuffer
    Dim msg As String, color As Long
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    'color = buffer.ReadLong
    
    Call GlobalMsg(msg, White, False)
    
    Set Buffer = Nothing
End Sub
