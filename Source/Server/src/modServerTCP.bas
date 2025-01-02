Attribute VB_Name = "modServerTCP"
Option Explicit

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Sub UpdateCaption()
    If LenB(SERVER_NAME) <= 0 Then
        frmServer.Caption = "World Server [Port: " & CStr(frmServer.Socket(0).LocalPort) & "] Players Online: " & TotalOnlinePlayers
    Else
        frmServer.Caption = SERVER_NAME & " [Port: " & CStr(frmServer.Socket(0).LocalPort) & "] Players Online: " & TotalOnlinePlayers
    End If
    'UpdateStatFile
    SendServerInfo
    DoEvents
End Sub

Function IsConnected(ByVal index As Long) As Boolean

    If frmServer.Socket(index).state = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal index As Long) As Boolean

    If IsConnected(index) Then
        If TempPlayer(index).InGame Then
            IsPlaying = True
        End If
    End If

End Function

Function IsLoggedIn(ByVal index As Long) As Boolean

    If IsConnected(index) Then
        If LenB(Trim$(player(index).login)) > 0 Then
            IsLoggedIn = True
        End If
    End If

End Function

Function IsMultiAccounts(ByVal login As String) As Boolean
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If LCase$(Trim$(player(i).login)) = LCase$(login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next
    
    CheckAccLockTime (login)
    
    
    If isAccountOnDiffServer(login) = True Then IsMultiAccounts = True

End Function

Public Sub RefreshAccLocks()
Dim i As Long

    For i = 1 To Player_HighIndex

        LockPlayerLogin player(i).login

    Next i
    
End Sub

Public Function CheckAccLockTime(login As String)
Dim i As Long
Dim Tick As Long
Dim Temp As String
If Not usePlayerLock Then Exit Function
    Dim Path As String
    Dim F As Long
    Path = App.Path & "\Data\AccLock\" & Trim$(login) & ".lock"

    If Not FileExist(Path, True) Then Exit Function

    F = FreeFile
    Open Path For Input As #F

    Input #F, Temp
    Tick = CLng(Temp)
    
    Close #F
    
    If GetRealTickCount > Tick + 900000 Then
        'they've been gone for over 15 minutes. That's two cycles of the AccRefresh.
        UnLockPlayerLogin Trim$(login)
        SendHubLog "Player " & Trim$(login) & " has been locked for over 15 mins. Unlocking."
    End If
End Function

Function isAccountOnDiffServer(ByVal login As String) As Boolean
If Not usePlayerLock Then Exit Function
    Dim Path As String
    Path = App.Path & "\Data\AccLock\" & Trim$(login) & ".lock"
    If FileExist(Path, True) Then
        isAccountOnDiffServer = True
        SendHubLog "Player " & Trim$(login) & " is attempting to login while locked!"
    End If
    
End Function

Function LockPlayerLogin(login As String) As Boolean
'On Error Resume Next
If Not usePlayerLock Then Exit Function
    If LenB(Trim$(login)) = 0 Then Exit Function
    Dim Path As String
    Dim DidExist As Boolean
    Dim F As Long
    Path = App.Path & "\Data\AccLock\" & Trim$(login) & ".lock"

    If FileExist(Path, True) Then DidExist = True

    F = FreeFile
    Open Path For Output As F
    Write #F, GetRealTickCount
    Close F

    If FileExist(Path, True) Then LockPlayerLogin = True Else LockPlayerLogin = False
    If DidExist = True Then SendHubLog "Player " & Trim$(login) & " lock has been refreshed." Else SendHubLog "Player " & Trim$(login) & " has been locked."
End Function

Public Sub UpdateStatFile(Optional ShutDown As Boolean = False)
On Error Resume Next
'Disabled due to not using/not testing..
Exit Sub
    Dim Path As String
    Dim F As Long
    Path = App.Path & "\Data\Servers\" & IIf(Options.OverridePort, Options.OverridePort, Options.Port) & ".dat"

    If ShutDown = True Then Kill Path: Exit Sub
    
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong TotalOnlinePlayers
    
    Dim i As Integer
    For i = 1 To Player_HighIndex
        Buffer.WriteString Trim$(player(i).Name)
        Buffer.WriteString Trim$(Class(player(i).Class).Name)
        Buffer.WriteString Trim$(player(i).level)
    Next
    
    F = FreeFile
    Open Path For Binary As #F
    Put #F, , Buffer.ToArray
    Close #F

End Sub

Function UnLockPlayerLogin(login As String)
On Error Resume Next
If Not usePlayerLock Then Exit Function
    Dim Path As String
    If LenB(Trim$(login)) = 0 Then Exit Function
    Path = App.Path & "\Data\AccLock\" & Trim$(login) & ".lock"
    
    If FileExist(Path, True) Then Kill Path
    
    SendHubLog "UnLockPlayerLogin: " & Trim$(login)
    
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim i As Long
    Dim N As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = IP Then
                N = N + 1

                If (N > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Function IsBanned(ByVal IP As String) As Boolean
    Dim FileName As String
    Dim fIP As String
    Dim fName As String
    Dim F As Long
    FileName = App.Path & "\data\banlist.txt"

    ' Check if file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If

    F = FreeFile
    Open FileName For Input As #F

    Do While Not EOF(F)
        Input #F, fIP
        Input #F, fName

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

Function IsRangeBanned(ByVal IP As String) As Boolean
    Dim FileName As String
    FileName = App.Path & "\data\RangeBan.ini"

    IsRangeBanned = False
    ' Check if file exists
    If Not FileExist("data\RangeBan.ini") Then
        Exit Function
    End If

    Dim NIPS As Long
    NIPS = GetVar(FileName, "IPS", "Number")
    While NIPS > 0
        Dim BannedIP As String
        Dim Digits As Long
        BannedIP = Trim$(GetVar(FileName, "IPS", "IP" & NIPS))
        Digits = CInt(GetVar(FileName, "IPS", "Digits" & NIPS))
        
        If left$(BannedIP, Digits) = left$(IP, Digits) Then
            Dim NExceptions As Long
            NExceptions = GetVar(FileName, "EXCEPTIONS", "Number")
            
            While NExceptions > 0
                Dim CheckIP As String
                CheckIP = Trim$(GetVar(FileName, "EXCEPTIONS", "IP" & NExceptions))
                If CheckIP = IP + "." Then
                    Exit Function
                End If
                    
                NExceptions = NExceptions - 1
            Wend
            
            IsRangeBanned = True
            Exit Function
        End If
        NIPS = NIPS - 1
    Wend
    
End Function

Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte, Optional ForceSend As Boolean = False)
Dim Buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(index) Or ForceSend = True Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()

        frmServer.Socket(index).SendData Buffer.ToArray()
        
        PacketsSent = PacketsSent + 1
        BytesSent = BytesSent + (UBound(TempData) - LBound(TempData)) + 1
        
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If

    Next

End Sub

Sub SendDataToAllBut(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMap(ByVal mapnum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal mapnum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                If i <> index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If

    Next

End Sub

Sub SendDataToParty(ByVal partynum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To Party(partynum).MemberCount
        If Party(partynum).Member(i) > 0 Then
            Call SendDataTo(Party(partynum).Member(i), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal msg As String, ByVal color As Byte, Optional Forward As Boolean = False)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    
    Buffer.WriteLong SGlobalMsg
    Buffer.WriteString msg
    Buffer.WriteLong color
    SendDataToAll Buffer.ToArray
    If Forward Then ForwardGlobalMsg "[" & SERVER_NAME & "] " & msg
    
    Set Buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    
    
    Buffer.WriteLong SAdminMsg
    Buffer.WriteString msg
    Buffer.WriteLong color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess_Mode(i) > 0 Then
            SendDataTo i, Buffer.ToArray
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal msg As String, ByVal color As Byte, Optional ByVal IsSystem As Boolean = True)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
        Debug.Assert (index > 0)
        If index = 0 Then Exit Sub
    
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString msg
    Buffer.WriteLong color
    Buffer.WriteByte IsSystem
    SendDataTo index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub MapMsg(ByVal mapnum As Long, ByVal msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    

    Buffer.WriteLong SMapMsg
    Buffer.WriteString msg
    Buffer.WriteLong color
    SendDataToMap mapnum, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteString msg
    SendDataTo index, Buffer.ToArray
    DoEvents
    Call CloseSocket(index)
    
    Set Buffer = Nothing
End Sub

Public Sub ServerFullMsg(ByVal index As Long, ByVal msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SFullMsg
    Buffer.WriteString msg
    SendDataTo index, Buffer.ToArray, True
    DoEvents
    
    Set Buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal partynum As Long, ByVal msg As String, ByVal color As Byte)
Dim i As Long
    ' send message to all people
    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partynum).Member(i) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partynum).Member(i)) And IsPlaying(Party(partynum).Member(i)) Then
                PlayerMsg Party(partynum).Member(i), msg, color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)

    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been kicked for (" & Reason & ")", White, False)
        End If

        Call AlertMsg(index, "You have been kicked from " & Options.Game_Name & ".")
    End If

End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        Else
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            ServerFullMsg index, "Too many connections, sorry!"
            DoEvents
            frmServer.Socket(0).Close
            frmServer.Socket(0).Listen
        End If
    End If

End Sub

Sub SocketConnected(ByVal index As Long)
Dim i As Long

    If index <> 0 Then
        ' make sure they're not banned
        If Not IsBanned(GetPlayerIP(index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(index) & ".")
        Else
            Call AlertMsg(index, "You have been banned from" & Options.Game_Name & ", and you can't play.")
        End If
        ' re-set the high index
        Player_HighIndex = 0
        For i = MAX_PLAYERS To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
        ' send the new highindex to all logged in players
        SendHighIndex
    End If
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

If index = 0 Then Exit Sub

    If GetPlayerAccess_Mode(index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(index).DataBytes > 1000 Then
            Exit Sub
        End If
      
        ' Check for packet flooding
        If TempPlayer(index).DataPackets > 80 Then
            Exit Sub
        End If
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    If GetRealTickCount >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = GetRealTickCount + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(index).Buffer.length >= 4 Then
        pLength = TempPlayer(index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(index).Buffer.length - 4
        If pLength <= TempPlayer(index).Buffer.length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).Buffer.ReadLong
            HandleData index, TempPlayer(index).Buffer.ReadBytes(pLength)
            
            PacketsReceived = PacketsReceived + 1
        End If
        
        pLength = 0
        If TempPlayer(index).Buffer.length >= 4 Then
            pLength = TempPlayer(index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
    
    BytesReceived = BytesReceived + DataLength
            
    TempPlayer(index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal index As Long)
Dim i As Integer
    

    If index > 0 Then
        'Pet check
        If TempPlayer(index).TempPet.TempPetSlot > 0 Then
            Call PetDisband(index, GetPlayerMap(index), False)
            For i = 1 To Player_HighIndex
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    SendMap i, GetPlayerMap(index)
                End If
            Next
        End If
        
        Call LeftGame(index)
        Call TextAdd("Connection from " & GetPlayerIP(index) & " has been terminated.")
        frmServer.Socket(index).Close
        Call UpdateCaption
        Call ClearPlayer(index)
        'UnLockPlayerLogin player(index).login
    End If

End Sub

Public Sub MapCache_Create(ByVal mapnum As Long, ByRef map As MapRec)
    Dim MapData As String
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Call Buffer.WriteBytes(Compress(GetMapData(map)))

    MapCache(mapnum).Data = Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim N As Long
    Dim i As Long
    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            'If i <> index Then 'And (GetPlayerAccess_Mode(index) >= ADMIN_MAPPER Or GetPlayerAccess_Mode(i) = 0) Then
                s = s & GetPlayerName(i) & ", "
                N = N + 1
            'End If
        End If

    Next

    If N = 0 Then
        s = "You are alone."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & N & " adventurers online: " & s & "."
    End If
    
    Dim sSend As String
    Dim sTemp As String
    Dim comma As Long
    Dim ii As Integer
    
    sTemp = s
    
    ii = Len(s) / 60
    If ii > 1 Then
        Do While Len(sTemp) > 0
            comma = InStr(60, sTemp, ",")
            If comma = 0 Then comma = Len(sTemp) Else comma = comma - 1
            sSend = left(sTemp, comma)
            If Len(sTemp) <> Len(sSend) Then sTemp = right(sTemp, Len(sTemp) - Len(sSend) - 2) Else sTemp = vbNullString
            Call PlayerMsg(index, sSend, WhoColor)
            'DoEvents
        Loop
    Else
        Call PlayerMsg(index, s, WhoColor)
    End If
End Sub

Function PlayerData(ByVal index As Long) As Byte()
    Dim Buffer As clsBuffer, i As Long

    If index > MAX_PLAYERS Then Exit Function
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong index
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerLevel(index)
    Buffer.WriteLong GetPlayerPOINTS(index)
    Buffer.WriteLong GetPlayerSprite(index)
    Buffer.WriteLong GetPlayerMap(index)
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    Buffer.WriteLong GetPlayerAccess_Mode(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteLong GetPlayerClass(index)
    Buffer.WriteLong GetPlayerVisible(index)
    'Kill Counter
    Buffer.WriteLong player(index).Kill
    Buffer.WriteLong player(index).Dead
    Buffer.WriteLong player(index).NpcKill
    Buffer.WriteLong player(index).NpcDead
    Buffer.WriteLong player(index).EnviroDead
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(index, i)
    Next
    
    If player(index).GuildFileId > 0 Then
        If TempPlayer(index).tmpGuildSlot > 0 Then
            Buffer.WriteByte 1
            Buffer.WriteString GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name
            Buffer.WriteLong player(index).GuildMemberId
        End If
    Else
        Buffer.WriteByte 0
    End If
    
    'Triforces
    For i = 1 To TriforceType.TriforceType_Count - 1
        Buffer.WriteByte player(index).triforce(i)
    Next
    
    'Ice
    Buffer.WriteByte player(index).onIce
    Buffer.WriteByte player(index).IceDir
    
    'Rupee System
    Buffer.WriteByte player(index).RupeeBags
    
    'Custom Sprite
    Buffer.WriteByte player(index).CustomSprite
    
    Buffer.WriteLong GetPlayerSpeed(index, MOVING_WALKING)
    Buffer.WriteLong GetPlayerSpeed(index, MOVING_RUNNING)
    
    Buffer.WriteByte GetPlayerState(index)

    PlayerData = Buffer.ToArray()
    Set Buffer = Nothing
End Function




Sub SendJoinMap(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> index Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    SendDataTo index, PlayerData(i)
                End If
            End If
        End If
    Next
    
    
    
    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    Set Buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal mapnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeft
    Buffer.WriteLong index
    SendDataToMapBut index, mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerData(ByVal index As Long)
    Dim packet As String
    SendDataToMap GetPlayerMap(index), PlayerData(index)
End Sub

Sub SendMap(ByVal index As Long, ByVal mapnum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(MapCache(mapnum).Data) - LBound(MapCache(mapnum).Data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteLong mapnum
    Buffer.WriteBytes MapCache(mapnum).Data()
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData
    
    Dim ItemHighIndex As Long
    ItemHighIndex = TempMap(mapnum).Item_highindex
    Buffer.WriteLong ItemHighIndex

    For i = 1 To ItemHighIndex
        Buffer.WriteString MapItem(mapnum, i).playerName
        Buffer.WriteLong MapItem(mapnum, i).Num
        Buffer.WriteLong MapItem(mapnum, i).Value
        Buffer.WriteByte MapItem(mapnum, i).X
        Buffer.WriteByte MapItem(mapnum, i).Y
    Next

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsToAll(ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData
    
    Dim ItemHighIndex As Long
    ItemHighIndex = TempMap(mapnum).Item_highindex
    Buffer.WriteLong ItemHighIndex

    For i = 1 To ItemHighIndex
        Buffer.WriteString MapItem(mapnum, i).playerName
        Buffer.WriteLong MapItem(mapnum, i).Num
        Buffer.WriteLong MapItem(mapnum, i).Value
        Buffer.WriteLong MapItem(mapnum, i).X
        Buffer.WriteLong MapItem(mapnum, i).Y
    Next

    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemTo(ByVal index As Long, ByVal mapnum As Long, ByVal ItemIndex As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapSingularItemData
    
    Buffer.WriteLong ItemIndex

    Buffer.WriteString MapItem(mapnum, ItemIndex).playerName
    Buffer.WriteLong MapItem(mapnum, ItemIndex).Num
    Buffer.WriteLong MapItem(mapnum, ItemIndex).Value
    Buffer.WriteLong MapItem(mapnum, ItemIndex).X
    Buffer.WriteLong MapItem(mapnum, ItemIndex).Y
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemToAll(ByVal mapnum As Long, ByVal ItemIndex As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapSingularItemData
    
    Buffer.WriteString MapItem(mapnum, ItemIndex).playerName
    Buffer.WriteLong MapItem(mapnum, ItemIndex).Num
    Buffer.WriteLong MapItem(mapnum, ItemIndex).Value
    Buffer.WriteLong MapItem(mapnum, ItemIndex).X
    Buffer.WriteLong MapItem(mapnum, ItemIndex).Y
    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal mapnum As Long, ByVal mapnpcnum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcVitals
    Buffer.WriteLong mapnpcnum
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong MapNpc(mapnum).NPC(mapnpcnum).vital(i)
    Next

    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Function MapNpcData(ByVal mapnum As Long, ByVal mapnpcnum As Long, Optional ByVal WriteMapNPCNum As Boolean = False) As Byte()
Dim Buffer As clsBuffer, i As Long
Set Buffer = New clsBuffer

If WriteMapNPCNum Then
    Buffer.WriteLong mapnpcnum
End If
Buffer.WriteLong MapNpc(mapnum).NPC(mapnpcnum).Num
Buffer.WriteByte MapNpc(mapnum).NPC(mapnpcnum).X
Buffer.WriteByte MapNpc(mapnum).NPC(mapnpcnum).Y
Buffer.WriteByte MapNpc(mapnum).NPC(mapnpcnum).dir
Buffer.WriteLong MapNpc(mapnum).NPC(mapnpcnum).vital(HP)
Buffer.WriteLong MapNpc(mapnum).NPC(mapnpcnum).PetData.Owner

MapNpcData = Buffer.ToArray()
Set Buffer = Nothing
End Function

Sub SendClearMapNpcTo(ByVal index As Long, ByVal mapnum As Long, ByVal mapnpcnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapSingularNpcData
    
    Buffer.WriteLong mapnpcnum
    Buffer.WriteLong 0
    Buffer.WriteByte 0
    Buffer.WriteByte 0
    Buffer.WriteByte 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub


Sub SendMapNpcTo(ByVal index As Long, ByVal mapnum As Long, ByVal mapnpcnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapSingularNpcData
    
    Buffer.WriteBytes MapNpcData(mapnum, mapnpcnum, True)

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendClearMapNpcToMap(ByVal mapnum As Long, ByVal mapnpcnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapSingularNpcData
    
    Buffer.WriteLong mapnpcnum
    Buffer.WriteLong 0
    Buffer.WriteByte 0
    Buffer.WriteByte 0
    Buffer.WriteByte 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0

    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcToMap(ByVal mapnum As Long, ByVal mapnpcnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapSingularNpcData
    
    Buffer.WriteBytes MapNpcData(mapnum, mapnpcnum, True)

    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData
    
    Dim npc_highindex As Long
    npc_highindex = TempMap(mapnum).npc_highindex
    Buffer.WriteLong npc_highindex
    

    For i = 1 To npc_highindex
        Buffer.WriteBytes MapNpcData(mapnum, i)
    Next

    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData
    
    Dim npc_highindex As Long
    npc_highindex = TempMap(mapnum).npc_highindex
    Buffer.WriteLong npc_highindex
    
    For i = 1 To npc_highindex
        Buffer.WriteBytes MapNpcData(mapnum, i)
    Next

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKeysTo(ByVal index As Long, ByVal mapnum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim state As Byte
    Dim X As Long
    Dim Y As Long
    
    Dim i As Long
    With TempTile(mapnum)
    For i = 1 To .NumDoors
        If CanRenderTempDoor(mapnum, i) Then
            X = .Door(i).X
            Y = .Door(i).Y
            state = .Door(i).state
            
            SendMapKey index, X, Y, state
        End If
    Next
        
    End With
    
End Sub


Sub SendItems(ByVal index As Long)
    Dim i As Long
    Call SendUpdateItemsTo(index)
    Exit Sub
    For i = 1 To MAX_ITEMS

        If ItemExists(i) Then
            Call SendUpdateItemTo(index, i)
        End If

    Next

End Sub

Sub SendAnimations(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If AnimationExists(i) Then
            Call SendUpdateAnimationTo(index, i)
        End If

    Next

End Sub

Sub SendNpcs(ByVal index As Long)

    Call SendUpdateNPCSTo(index)
    Exit Sub
    Dim i As Long

    For i = 1 To MAX_NPCS

        If NPCExists(i) Then
            Call SendUpdateNpcTo(index, i)
        End If

    Next

End Sub

Sub SendResources(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If ResourceExists(i) Then
            Call SendUpdateResourceTo(index, i)
        End If

    Next

End Sub

Sub SendInventory(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(index, i)
        Buffer.WriteLong GetPlayerInvItemValue(index, i)
    Next

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal invSlot As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteLong invSlot
    Buffer.WriteLong GetPlayerInvItemNum(index, invSlot)
    Buffer.WriteLong GetPlayerInvItemValue(index, invSlot)
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerWornEq
    Buffer.WriteLong GetPlayerEquipment(index, Armor)
    Buffer.WriteLong GetPlayerEquipment(index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(index, helmet)
    Buffer.WriteLong GetPlayerEquipment(index, Shield)
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerEquipment(index, Armor)
    Buffer.WriteLong GetPlayerEquipment(index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(index, helmet)
    Buffer.WriteLong GetPlayerEquipment(index, Shield)
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong PlayerNum
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Armor)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Weapon)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, helmet)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Shield)
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendVital(ByVal index As Long, ByVal vital As Vitals)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Select Case vital
        Case HP
            Buffer.WriteLong SPlayerHp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.MP)
    End Select

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendEXP(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong GetPlayerExp(index)
    Buffer.WriteLong GetPlayerNextLevel(index)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendStats(ByVal index As Long)
Dim i As Long
Dim packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(index, i)
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendStat(ByVal index As Long, ByVal stat As Stats)
Dim i As Long
Dim packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerStat
    Buffer.WriteByte stat
    Buffer.WriteInteger GetPlayerStat(index, stat)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendWelcome(ByVal index As Long)
Dim RGB As Integer
' Send visibility message
If GetPlayerAccess_Mode(index) > ADMIN_MONITOR Then
If GetPlayerVisible(index) = 1 Then
Call PlayerMsg(index, "[INVISIBLE]", AlertColor)
End If
End If
    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Dim splitData() As String
        Dim i As Integer
        splitData = Split(Options.MOTD, "\r")
    
        For i = 0 To UBound(splitData)
            Call PlayerMsg(index, splitData(i), Cyan)
        Next i
    
        
    End If

    ' Send whos online
    Call SendWhosOnline(index)
End Sub

Sub SendClasses(ByVal index As Long)
    Dim packet As String
    Dim i As Long, N As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClassesData
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        N = UBound(Class(i).MaleSprite)
        
        ' send array size
        Buffer.WriteLong N
        
        ' loop around sending each sprite
        For q = 0 To N
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        N = UBound(Class(i).FemaleSprite)
        
        ' send array size
        Buffer.WriteLong N
        
        ' loop around sending each sprite
        For q = 0 To N
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        Buffer.WriteLong Class(i).Face
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).stat(q)
        Next
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal index As Long)
    Dim packet As String
    Dim i As Long, N As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewCharClasses
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        N = UBound(Class(i).MaleSprite)
        ' send array size
        Buffer.WriteLong N
        ' loop around sending each sprite
        For q = 0 To N
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        N = UBound(Class(i).FemaleSprite)
        ' send array size
        Buffer.WriteLong N
        ' loop around sending each sprite
        For q = 0 To N
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).stat(q)
        Next
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLeftGame(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong index
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    SendDataToAllBut index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXY
    Buffer.WriteByte GetPlayerX(index)
    Buffer.WriteByte GetPlayerY(index)
    Buffer.WriteByte GetPlayerDir(index)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXYToMap(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXYMap
    Buffer.WriteLong index
    Buffer.WriteByte GetPlayerX(index)
    Buffer.WriteByte GetPlayerY(index)
    Buffer.WriteByte GetPlayerDir(index)
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(item(ItemNum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(item(ItemNum)), ItemSize
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteByte False 'no useful data, all sent
    Buffer.WriteBytes ItemData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    
    Dim usefuldata As Boolean
    usefuldata = SendUsefulDataToPlayer(index)
    If usefuldata Then
        ItemData = CompressData(GetItemUsefulData(ItemNum), 2)
    Else
        ItemSize = LenB(item(ItemNum))
        ReDim ItemData(ItemSize - 1)
        CopyMemory ItemData(0), ByVal VarPtr(item(ItemNum)), ItemSize
    End If
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteByte usefuldata
    Buffer.WriteBytes ItemData
    SendDataTo index, Buffer.ToArray()
    
    ByteCounter = ByteCounter + Buffer.length
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemsTo(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim nItems As Long
    Dim ItemData() As Byte
    
    Set Buffer = New clsBuffer
    
    Dim usefuldata As Boolean
    usefuldata = SendUsefulDataToPlayer(index)
    
    Buffer.WriteLong SUpdateItems
    Buffer.WriteByte usefuldata
    
    Dim NewItemData() As Byte
    
    Dim CompressedBuffer As clsBuffer
    Set CompressedBuffer = New clsBuffer
    
    Dim i As Long
    If usefuldata Then
        For i = 1 To MAX_ITEMS
            If ItemExists(i) Then
                ItemData = GetItemUsefulData(i)
                CompressedBuffer.WriteByte True 'item exists
                CompressedBuffer.WriteLong UBound(ItemData) - LBound(ItemData) + 1
                CompressedBuffer.WriteBytes ItemData
            Else
                CompressedBuffer.WriteByte False 'item does not exist
            End If
        Next
    Else
        For i = 1 To MAX_ITEMS
            ItemSize = LenB(item(i))
            ReDim ItemData(ItemSize - 1)
            CopyMemory ItemData(0), ByVal VarPtr(item(i)), ItemSize
            CompressedBuffer.WriteBytes ItemData
        Next
    End If
    
    Buffer.WriteBytes CompressData(CompressedBuffer.ToArray, 3)
    SendDataTo index, Buffer.ToArray()
    
    ByteCounter = ByteCounter + Buffer.length
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteByte False 'no useful data, all sent
    Buffer.WriteBytes AnimationData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal index As Long, ByVal AnimationNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    
    Dim usefuldata As Boolean
    usefuldata = SendUsefulDataToPlayer(index)
    If usefuldata Then
        AnimationData = GetAnimationUseFulData(AnimationNum)
    Else
        AnimationSize = LenB(Animation(AnimationNum))
        ReDim AnimationData(AnimationSize - 1)
        CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    End If
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteByte usefuldata
    Buffer.WriteBytes AnimationData
    SendDataTo index, Buffer.ToArray()
    
    ByteCounter = ByteCounter + Buffer.length
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal npcnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim npcSize As Long
    Dim npcData() As Byte
    Set Buffer = New clsBuffer
    npcSize = LenB(NPC(npcnum))
    ReDim npcData(npcSize - 1)
    CopyMemory npcData(0), ByVal VarPtr(NPC(npcnum)), npcSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong npcnum
    Buffer.WriteByte False 'no useful data, all sent
    Buffer.WriteBytes npcData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal npcnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim npcSize As Long
    Dim npcData() As Byte
    Set Buffer = New clsBuffer
    
    Dim usefuldata As Boolean
    usefuldata = SendUsefulDataToPlayer(index)
    If usefuldata Then
        npcData = GetNPCUsefulData(npcnum)
    Else
        npcSize = LenB(NPC(npcnum))
        ReDim npcData(npcSize - 1)
        CopyMemory npcData(0), ByVal VarPtr(NPC(npcnum)), npcSize
    End If
   
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong npcnum
    Buffer.WriteByte usefuldata
    Buffer.WriteBytes npcData
    SendDataTo index, Buffer.ToArray()
    
    ByteCounter = ByteCounter + Buffer.length
    Set Buffer = Nothing
End Sub

Sub SendUpdateNPCSTo(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim npcSize As Long
    Dim nnpcs As Long
    Dim npcData() As Byte
    
    Set Buffer = New clsBuffer
    
    Dim usefuldata As Boolean
    usefuldata = SendUsefulDataToPlayer(index)
    
    Buffer.WriteLong SUpdateNPCS
    Buffer.WriteByte usefuldata
    
    Dim NewnpcData() As Byte
    
    Dim CompressedBuffer As clsBuffer
    Set CompressedBuffer = New clsBuffer
    
    Dim i As Long
    If usefuldata Then
        For i = 1 To MAX_NPCS
            If NPCExists(i) Then
                npcData = GetNPCUsefulData(i)
                CompressedBuffer.WriteByte True 'npc exists
                CompressedBuffer.WriteLong UBound(npcData) - LBound(npcData) + 1
                CompressedBuffer.WriteBytes npcData
            Else
                CompressedBuffer.WriteByte False 'npc does not exist
            End If
        Next
    Else
        For i = 1 To MAX_NPCS
            npcSize = LenB(NPC(i))
            ReDim npcData(npcSize - 1)
            CopyMemory npcData(0), ByVal VarPtr(NPC(i)), npcSize
            CompressedBuffer.WriteBytes npcData
        Next
    End If
    
    Buffer.WriteBytes CompressData(CompressedBuffer.ToArray, 3)
    SendDataTo index, Buffer.ToArray()
    
    ByteCounter = ByteCounter + Buffer.length
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteByte False 'no useful data, all sent
    Buffer.WriteBytes ResourceData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal index As Long, ByVal ResourceNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    Dim usefuldata As Boolean
    usefuldata = SendUsefulDataToPlayer(index)
    
    If usefuldata Then
        ResourceData = GetResourceUsefulData(ResourceNum)
    Else
        ResourceSize = LenB(Resource(ResourceNum))
        ReDim ResourceData(ResourceSize - 1)
        CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    End If
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteByte usefuldata
    Buffer.WriteBytes ResourceData
    
    
    ByteCounter = ByteCounter + Buffer.length
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendShops(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If ShopExists(i) Then
            Call SendUpdateShopTo(index, i)
        End If

    Next

End Sub

Sub SendUpdateShopToAll(ByVal shopnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopnum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopnum
    Buffer.WriteBytes CompressData(ShopData, 2)

    SendDataToAll Buffer.ToArray()
    
    
    Set Buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal shopnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopnum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopnum
    Buffer.WriteBytes CompressData(ShopData, 2)
    
    SendDataTo index, Buffer.ToArray()
    
    ByteCounter = ByteCounter + Buffer.length
    Set Buffer = Nothing
End Sub

Sub SendSpells(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If SpellExists(i) Then
            Call SendUpdateSpellTo(index, i)
        End If

    Next

End Sub

Sub SendUpdateSpellToAll(ByVal spellnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong spellnum
    Buffer.WriteByte False 'no useful data, all sent
    Buffer.WriteBytes SpellData
    
    SendDataToAll Buffer.ToArray()
    
    ByteCounter = ByteCounter + Buffer.length
    Set Buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal spellnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    Dim usefuldata As Boolean
    usefuldata = SendUsefulDataToPlayer(index)
    If usefuldata Then
        SpellData = GetSpellUsefulData(spellnum)
    Else
        SpellSize = LenB(Spell(spellnum))
        ReDim SpellData(SpellSize - 1)
        CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    End If
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong spellnum
    Buffer.WriteByte usefuldata
    Buffer.WriteBytes SpellData
    
    SendDataTo index, Buffer.ToArray()
    
    ByteCounter = ByteCounter + Buffer.length
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong GetPlayerSpell(index, i)
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(GetPlayerMap(index)).Resource_Count

    If ResourceCache(GetPlayerMap(index)).Resource_Count > 0 Then

        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            Buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(i).ResourceState
            Buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(i).X
            Buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(i).Y
        Next

    End If

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal mapnum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(mapnum).Resource_Count

    If ResourceCache(mapnum).Resource_Count > 0 Then

        For i = 0 To ResourceCache(mapnum).Resource_Count
            Buffer.WriteByte ResourceCache(mapnum).ResourceData(i).ResourceState
            Buffer.WriteByte ResourceCache(mapnum).ResourceData(i).X
            Buffer.WriteByte ResourceCache(mapnum).ResourceData(i).Y
        Next

    End If

    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSingleResourceCacheToMap(ByVal mapnum As Long, ByVal Resource_Num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    If Resource_Num <= 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSingleResourceCache
    
    Buffer.WriteLong Resource_Num
    Buffer.WriteByte ResourceCache(mapnum).ResourceData(Resource_Num).ResourceState
    Buffer.WriteByte ResourceCache(mapnum).ResourceData(Resource_Num).X
    Buffer.WriteByte ResourceCache(mapnum).ResourceData(Resource_Num).Y

    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Sub SendSingleResourceCacheTo(ByVal index As Long, ByVal Resource_Num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSingleResourceCache

    Buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).ResourceState
    Buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).X
    Buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).Y

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Sub SendDoorAnimation(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SDoorAnimation
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendActionMsg(ByVal mapnum As Long, ByVal message As String, ByVal color As Long, ByVal MsgType As Long, ByVal X As Long, ByVal Y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer
        
    'If IsNumeric(message) And val(message) < 0 And val(message) > -2 Then Exit Sub
    
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString (message)
    Buffer.WriteLong color
    Buffer.WriteLong MsgType
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap mapnum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendBlood(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBlood
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendAnimation(ByVal mapnum As Long, ByVal Anim As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimation
    Buffer.WriteLong Anim
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteByte LockType
    Buffer.WriteLong LockIndex
    
    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCooldown(ByVal index As Long, ByVal slot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong slot
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendClearSpellBuffer(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearSpellBuffer
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal mapnum As Long, ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess_Mode(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong saycolour
    Buffer.WriteLong MapChat
    
    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess_Mode(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong saycolour
    Buffer.WriteLong GlobalChat
    
    SendDataToAll Buffer.ToArray()
    ForwardGlobalMsg "[" & SERVER_NAME & "] " & GetPlayerName(index) & ": " & message
    
    Set Buffer = Nothing
End Sub

Sub ResetShopAction(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendBlockedAction(ByVal index As Long, ByVal PlayerAction As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    Buffer.WriteByte PlayerAction
    Buffer.WriteByte IsActionBlocked(index, PlayerAction)
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendBank(ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        Buffer.WriteLong Bank(index).item(i).Num
        Buffer.WriteLong Bank(index).item(i).Value
    Next
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKey(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteByte Value
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKeyToMap(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteByte Value
    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendOpenShop(ByVal index As Long, ByVal shopnum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    Buffer.WriteLong shopnum
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal index As Long, ByVal Movement As Long, ByVal dir As Byte, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    Buffer.WriteLong index
    'buffer.WriteLong GetPlayerX(index)
    'buffer.WriteLong GetPlayerY(index)
    Buffer.WriteByte dir
    Buffer.WriteLong Movement
    
    If Not sendToSelf Then
        SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendTrade(ByVal index As Long, ByVal tradeTarget As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    Buffer.WriteLong tradeTarget
    Buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal index As Long, ByVal dataType As Byte)
Dim Buffer As clsBuffer
Dim i As Long
Dim tradeTarget As Long
Dim totalWorth As Long
    
    tradeTarget = TempPlayer(index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    Buffer.WriteByte dataType
    
If tradeTarget > 0 Then

    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).Num
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(index).TradeOffer(i).Num > 0 Then
                ' currency?
                If isItemStackable(TempPlayer(index).TradeOffer(i).Num) Then
                    totalWorth = totalWorth + (item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).Price * TempPlayer(index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If isItemStackable(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)) Then
                    totalWorth = totalWorth + (item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    Buffer.WriteLong totalWorth
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End If
End Sub

Sub SendTradeStatus(ByVal index As Long, ByVal Status As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    Buffer.WriteByte Status
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTarget(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    
    Dim Target As Long
    Target = TempPlayer(index).Target
    Buffer.WriteLong Target
    Buffer.WriteLong TempPlayer(index).TargetType
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHotbar(ByVal index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        Buffer.WriteLong player(index).Hotbar(i).slot
        Buffer.WriteByte player(index).Hotbar(i).sType
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLoginOk(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLoginOk
    Buffer.WriteLong index
    Buffer.WriteLong Player_HighIndex
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInGame(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SInGame
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHighIndex()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHighIndex
    Buffer.WriteLong Player_HighIndex
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSound(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapSound(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSoundToMap(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer
   
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub



Sub SendTradeRequest(ByVal index As Long, ByVal TradeRequest As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    Buffer.WriteString Trim$(player(TradeRequest).Name)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal index As Long, ByVal targetPlayer As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    Buffer.WriteString Trim$(player(targetPlayer).Name)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal partynum As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    Buffer.WriteByte 1
    Buffer.WriteLong Party(partynum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(partynum).Member(i)
    Next
    Buffer.WriteLong Party(partynum).MemberCount
    SendDataToParty partynum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal index As Long)
Dim Buffer As clsBuffer, i As Long, partynum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partynum = TempPlayer(index).inParty
    If partynum > 0 Then
        ' send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(partynum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(partynum).Member(i)
        Next
        Buffer.WriteLong Party(partynum).MemberCount
    Else
        ' send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal partynum As Long, ByVal index As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    Buffer.WriteLong index
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(index, i)
        Buffer.WriteLong player(index).vital(i)
    Next
    SendDataToParty partynum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpawnItemToMap(ByVal mapnum As Long, ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpawnItem
    Buffer.WriteLong index
    Buffer.WriteString MapItem(mapnum, index).playerName
    Buffer.WriteLong MapItem(mapnum, index).Num
    Buffer.WriteLong MapItem(mapnum, index).Value
    Buffer.WriteLong MapItem(mapnum, index).X
    Buffer.WriteLong MapItem(mapnum, index).Y
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendProjectileToMap(ByVal index As Long, ByVal PlayerProjectile As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SHandleProjectile
    Buffer.WriteLong PlayerProjectile
    Buffer.WriteLong index
    With TempPlayer(index).ProjecTile(PlayerProjectile)
        Buffer.WriteLong .Direction
        Buffer.WriteLong .Pic
        Buffer.WriteLong .range
        Buffer.WriteLong .Damage
        Buffer.WriteLong .Speed
    End With
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDoors(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_DOORS

        If LenB(Trim$(Doors(i).Name)) > 0 Then
            Call SendUpdateDoorsTo(index, i)
        End If

    Next

End Sub

Sub SendUpdateDoorToAll(ByVal DoorNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim DoorSize As Long
    Dim DoorData() As Byte
    Set Buffer = New clsBuffer
    
    DoorSize = LenB(Doors(DoorNum))
    ReDim DoorData(DoorSize - 1)
    CopyMemory DoorData(0), ByVal VarPtr(Doors(DoorNum)), DoorSize
    
    Buffer.WriteLong SUpdateDoors
    Buffer.WriteLong DoorNum
    Buffer.WriteBytes DoorData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateDoorsTo(ByVal index As Long, ByVal DoorNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim DoorSize As Long
    Dim DoorData() As Byte
    Set Buffer = New clsBuffer
    
    DoorSize = LenB(Doors(DoorNum))
    ReDim DoorData(DoorSize - 1)
    CopyMemory DoorData(0), ByVal VarPtr(Doors(DoorNum)), DoorSize
    
    Buffer.WriteLong SUpdateDoors
    Buffer.WriteLong DoorNum
    Buffer.WriteBytes DoorData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub PartyChatMsg(ByVal index As Long, ByVal msg As String, ByVal color As Byte)
Dim i As Long
Dim Member As Integer
Dim partynum As Long

partynum = TempPlayer(index).inParty

    ' not in a party?
    If TempPlayer(index).inParty = 0 Then
        Call PlayerMsg(index, "No est�s en un equipo.", BrightRed)
        Exit Sub
    End If
        
    SayMsg_Party index, msg, QBColor(White)
    
    Call AddLog(index, "Party #" & partynum & " map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " says, '" & msg & "'", PLAYER_LOG)

    'For i = 1 To MAX_PARTY_MEMBERS
        'Member = Party(partyNum).Member(i)
        ' is online, does exist?
        'If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
        ' yep, send the message!
            'Call PlayerMsg(Member, "[Party] " & GetPlayerName(index) & ": " & Msg, color)
            
        'End If
   ' Next
End Sub

Sub SendChatBubble(ByVal mapnum As Long, ByVal Target As Long, ByVal TargetType As Long, ByVal message As String, ByVal colour As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SChatBubble
    Buffer.WriteLong Target
    Buffer.WriteLong TargetType
    Buffer.WriteString message
    Buffer.WriteLong colour
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Sub SendLoad(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLoad
   
    SendDataTo index, Buffer.ToArray
    Set Buffer = Nothing
    TempPlayer(index).IsLoading = True
End Sub

Sub SendDone(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SDone
   
    SendDataTo index, Buffer.ToArray
    Set Buffer = Nothing
End Sub

Sub SendWeather(ByVal index As Long)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer
Buffer.WriteLong SSendWeather
If RainOn = True Then
     Buffer.WriteLong 1
Else
     Buffer.WriteLong 0
End If

SendDataTo index, Buffer.ToArray()
Set Buffer = Nothing


End Sub

Sub SendWeathertoAll(Optional SendToHub As Boolean = True)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer
Buffer.WriteLong SSendWeather
If RainOn = True Then
     Buffer.WriteLong 1
Else
     Buffer.WriteLong 0
End If

SendDataToAll Buffer.ToArray()
Set Buffer = Nothing

If SendToHub Then Call SendHubCommand(CommandsType.Weather, CStr(RainOn))


End Sub

Sub SendMovements(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_MOVEMENTS

        If LenB(Trim$(Movements(i).Name)) > 0 Then
            Call SendUpdateMovementsTo(index, i)
        End If

    Next

End Sub

Sub SendUpdateMovementToAll(ByVal MovementNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Byte
    
    
    Buffer.WriteLong SUpdateMovements
    Buffer.WriteLong MovementNum
    'buffer.WriteBytes movementData
    With Movements(MovementNum)
        
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteByte .Type
        Buffer.WriteByte .MovementsTable.Actual
        Buffer.WriteByte .MovementsTable.nelem
        If .MovementsTable.nelem > 0 Then
            For i = 1 To .MovementsTable.nelem
                Buffer.WriteByte .MovementsTable.vect(i).Data.Direction
                Buffer.WriteByte .MovementsTable.vect(i).Data.NumberOfTiles
            Next
        End If
        
        Buffer.WriteByte .Repeat
        
    End With
    

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateMovementsTo(ByVal index As Long, ByVal MovementNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Byte
    
    
    Buffer.WriteLong SUpdateMovements
    Buffer.WriteLong MovementNum
    With Movements(MovementNum)
        
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteByte .Type
        Buffer.WriteByte .MovementsTable.Actual
        Buffer.WriteByte .MovementsTable.nelem
        If .MovementsTable.nelem > 0 Then
            For i = 1 To .MovementsTable.nelem
                Buffer.WriteByte .MovementsTable.vect(i).Data.Direction
                Buffer.WriteByte .MovementsTable.vect(i).Data.NumberOfTiles
            Next
        End If
        
        Buffer.WriteByte .Repeat
        
    End With
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendActions(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ACTIONS

        If LenB(Trim$(Actions(i).Name)) > 0 Then
            Call SendUpdateActionsTo(index, i)
        End If

    Next

End Sub

Sub SendUpdateActionToAll(ByVal ActionNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Byte

    
    Buffer.WriteLong SUpdateActions
    Buffer.WriteLong ActionNum
    With Actions(ActionNum)
        
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteByte .Type
        Buffer.WriteByte .Moment
        Buffer.WriteLong .Data1
        Buffer.WriteLong .Data2
        Buffer.WriteLong .Data3
        Buffer.WriteLong .Data4
        
    End With
    

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateActionsTo(ByVal index As Long, ByVal ActionNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Byte
    
    
    Buffer.WriteLong SUpdateActions
    Buffer.WriteLong ActionNum
    With Actions(ActionNum)
        
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteByte .Type
        Buffer.WriteByte .Moment
        Buffer.WriteLong .Data1
        Buffer.WriteLong .Data2
        Buffer.WriteLong .Data3
        Buffer.WriteLong .Data4
        
    End With
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub



Sub NPCCache_Create_SendToAll(ByVal mapnum As Long, ByVal npcnum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SNPCCache
    Buffer.WriteLong mapnum
    Buffer.WriteLong npcnum
    Buffer.WriteLong map(mapnum).NPC(npcnum)
    Buffer.WriteLong MapNpc(mapnum).NPC(npcnum).Num
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If player(i).map = mapnum Then
                SendDataTo i, Buffer.ToArray()
            End If
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Sub NPCCache_Create(ByVal index As Long, ByVal mapnum As Long, ByVal npcnum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SNPCCache
    Buffer.WriteLong mapnum
    Buffer.WriteLong npcnum
    Buffer.WriteLong map(mapnum).NPC(npcnum)
    Buffer.WriteLong MapNpc(mapnum).NPC(npcnum).Num
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub


Sub SendPets(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_PETS

        If PetExists(i) Then
            Call SendUpdatePetsTo(index, i)
        End If

    Next

End Sub

Sub SendUpdatePetToAll(ByVal PetNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Byte

    
    Buffer.WriteLong SUpdatePets
    Buffer.WriteLong PetNum
    With Pet(PetNum)
        
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteLong .npcnum
        Buffer.WriteInteger .TamePoints
        Buffer.WriteByte .ExpProgression
        Buffer.WriteByte .pointsprogression
        Buffer.WriteLong .MaxLevel
        
    End With
    

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdatePetsTo(ByVal index As Long, ByVal PetNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    
    Buffer.WriteLong SUpdatePets
    Buffer.WriteLong PetNum
    With Pet(PetNum)
        
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteLong .npcnum
        Buffer.WriteInteger .TamePoints
        Buffer.WriteByte .ExpProgression
        Buffer.WriteByte .pointsprogression
        Buffer.WriteLong .MaxLevel
        
    End With
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPetData(ByVal index As Long, ByVal PlayerPetSlot As Byte)
    Dim Buffer As clsBuffer
    Dim i As Byte
    Set Buffer = New clsBuffer
    
    If player(index).Pet(PlayerPetSlot).NumPet < 1 Or player(index).Pet(PlayerPetSlot).NumPet > MAX_PETS Then GoTo SendNullData
    If Pet(player(index).Pet(PlayerPetSlot).NumPet).npcnum < 1 Or Pet(player(index).Pet(PlayerPetSlot).NumPet).npcnum > MAX_NPCS Then GoTo SendNullData
    
    Buffer.WriteLong SPetData
    Buffer.WriteByte PlayerPetSlot
    
    'buffer.WriteString NPC(Pet(Player(index).Pet(PlayerPetSlot).NumPet).NPCNum).Name
    Buffer.WriteInteger player(index).Pet(PlayerPetSlot).points
    Buffer.WriteLong CLng(GetPetExpPercent(index))
    Buffer.WriteLong player(index).Pet(PlayerPetSlot).level
    Buffer.WriteByte player(index).Pet(PlayerPetSlot).NumPet
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteByte player(index).Pet(PlayerPetSlot).StatsAdd(i)
    Next
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
    Exit Sub
    
SendNullData:

    Buffer.WriteLong SPetData
    Buffer.WriteByte PlayerPetSlot
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteByte 0
    Next
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
    Exit Sub
    
End Sub

Sub SendOpenTriforce(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SOpenTriforce
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendOnIce(ByVal index As Long, ByVal Ice As Boolean)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SOnIce
    Buffer.WriteByte Ice
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendIceDir(ByVal index As Long, ByVal dir As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SIceDir
    Buffer.WriteByte dir
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendBags(ByVal index As Long, ByVal Bags As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SBags
    Buffer.WriteByte Bags
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPoints(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPoints
    Buffer.WriteLong GetPlayerPOINTS(index)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLevel(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLevel
    Buffer.WriteLong GetPlayerLevel(index)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendJustice(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SJustice
    Buffer.WriteLong index
    Buffer.WriteByte GetPlayerPK(index)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendJusticeToMap(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SJustice
    Buffer.WriteLong index
    Buffer.WriteByte GetPlayerPK(index)
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapUpdate(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong GetPlayerMap(index)
    Buffer.WriteLong map(GetPlayerMap(index)).Revision
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerAttack(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerAttack
    Buffer.WriteLong index
    SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub




Sub SendCustomSprites(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_CUSTOM_SPRITES

        If LenB(Trim$(CustomSprites(i).Name)) > 0 Then
            Call SendUpdateCustomSpritesTo(index, i)
        End If

    Next

End Sub

Sub SendUpdateCustomSpriteToAll(ByVal CustomSpriteNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    'Dim CustomSpriteSize As Long
    'Dim CustomSpriteData() As Byte
    
    'CustomSpriteSize = LenB(CustomSprites(CustomSpriteNum))
    'ReDim CustomSpriteData(CustomSpriteSize - 1)
    'CopyMemory CustomSpriteData(0), ByVal VarPtr(CustomSprites(CustomSpriteNum)), CustomSpriteSize
    
    Buffer.WriteLong SUpdateCustomSprites
    Buffer.WriteLong CustomSpriteNum
    Buffer.WriteBytes GetCustomSpriteData(CustomSpriteNum)
    'buffer.WriteBytes CustomSpriteData
    'With CustomSprites(CustomSpriteNum)
            'buffer.WriteString .Name
            'buffer.WriteByte .NLayers
            'Dim i As Byte
            'For i = 1 To .NLayers
                'buffer.WriteLong .Layers(i).Sprite
                'buffer.WriteByte .Layers(i).UseCenterPosition
                'buffer.WriteByte .Layers(i).UsePlayerSprite
                'Dim j As Byte
                'For j = 0 To MAX_SPRITE_ANIMS - 1
                    'buffer.WriteByte .Layers(i).fixed.EnabledAnims(j)
                'Next
                'For j = 0 To MAX_DIRECTIONS - 1
                    'buffer.WriteInteger .Layers(i).CentersPositions(j).X
                    'buffer.WriteInteger .Layers(i).CentersPositions(j).Y
                'Next
            'Next
    'End With

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateCustomSpritesTo(ByVal index As Long, ByVal CustomSpriteNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    'Dim CustomSpriteSize As Long
    'Dim CustomSpriteData() As Byte
    
    'CustomSpriteSize = LenB(CustomSprites(CustomSpriteNum))
    'ReDim CustomSpriteData(CustomSpriteSize - 1)
    'CopyMemory CustomSpriteData(0), ByVal VarPtr(CustomSprites(CustomSpriteNum)), CustomSpriteSize
    
    Buffer.WriteLong SUpdateCustomSprites
    Buffer.WriteLong CustomSpriteNum
    Buffer.WriteBytes GetCustomSpriteData(CustomSpriteNum)
    'buffer.WriteBytes CustomSpriteData
    'With CustomSprites(CustomSpriteNum)
            'buffer.WriteString .Name
            'buffer.WriteByte .NLayers
            'Dim i As Byte
            'For i = 1 To .NLayers
                'buffer.WriteLong .Layers(i).Sprite
                'buffer.WriteByte .Layers(i).UseCenterPosition
                'buffer.WriteByte .Layers(i).UsePlayerSprite
                'Dim j As Byte
                'For j = 0 To MAX_SPRITE_ANIMS - 1
                    'buffer.WriteByte .Layers(i).fixed.EnabledAnims(j)
                'Next
                'For j = 0 To MAX_DIRECTIONS - 1
                    'buffer.WriteInteger .Layers(i).CentersPositions(j).X
                    'buffer.WriteInteger .Layers(i).CentersPositions(j).Y
                'Next
            'Next
    'End With

    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub


Sub SendPlayerSprite(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerSprite
    
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerSprite(index)
    Buffer.WriteByte GetPlayerCustomSprite(index)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpriteToMap(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerSprite
    
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerSprite(index)
    Buffer.WriteByte GetPlayerCustomSprite(index)
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMaxWeight(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMaxWeight
    
    Buffer.WriteLong GetPlayerMaxWeight(index)
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdate(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdate
    
    Buffer.WriteString Options.Update
    Buffer.WriteString Options.Instructions
    
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub


    


