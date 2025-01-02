Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = Options.Game_Name & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Sub CreateFullMapCache()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next

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
        If LenB(Trim$(Player(index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If

End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next

End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim i As Long
    Dim n As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = IP Then
                n = n + 1

                If (n > 1) Then
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

Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        frmServer.Socket(index).SendData Buffer.ToArray()
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

Sub SendDataToParty(ByVal partyNum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To Party(partyNum).MemberCount
        If Party(partyNum).Member(i) > 0 Then
            Call SendDataTo(Party(partyNum).Member(i), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataToAll Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, Buffer.ToArray
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataTo index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub MapMsg(ByVal mapnum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataToMap mapnum, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteString Msg
    SendDataTo index, Buffer.ToArray
    DoEvents
    Call CloseSocket(index)
    
    Set Buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal partyNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim i As Long
    ' send message to all people
    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partyNum).Member(i) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
                PlayerMsg Party(partyNum).Member(i), Msg, Color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)

    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " he has been banned for (..." & Reason & ")", White)
        End If

        Call AlertMsg(index, "You've lost your connection to" & Options.Game_Name & ".")
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
            Call AlertMsg(index, "Has sido baneado de " & Options.Game_Name & ", and you can't play.")
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

If GetPlayerAccess(index) <= 0 Then
' Check for data flooding
If TempPlayer(index).DataBytes > 1000 Then
Exit Sub
End If

' Check for packet flooding
If TempPlayer(index).DataPackets > 25 Then
Exit Sub
End If
End If
            
    ' Check if elapsed time has passed
    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = GetTickCount + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(index).Buffer.Length >= 4 Then
        pLength = TempPlayer(index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(index).Buffer.Length - 4
        If pLength <= TempPlayer(index).Buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).Buffer.ReadLong
            HandleData index, TempPlayer(index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(index).Buffer.Length >= 4 Then
            pLength = TempPlayer(index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal index As Long)
Dim i As Integer
    

    If index > 0 Then
        'Pet check
        If TempPlayer(index).TempPetSlot > 0 Then
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
    End If

End Sub

Public Sub MapCache_Create(ByVal mapnum As Long)
    Dim MapData As String
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong mapnum
    Buffer.WriteString Trim$(Map(mapnum).Name)
    Buffer.WriteString Trim$(Map(mapnum).Music)
    Buffer.WriteLong Map(mapnum).Revision
    Buffer.WriteByte Map(mapnum).Moral
    Buffer.WriteLong Map(mapnum).Up
    Buffer.WriteLong Map(mapnum).Down
    Buffer.WriteLong Map(mapnum).Left
    Buffer.WriteLong Map(mapnum).Right
    Buffer.WriteLong Map(mapnum).BootMap
    Buffer.WriteByte Map(mapnum).BootX
    Buffer.WriteByte Map(mapnum).BootY
    Buffer.WriteByte Map(mapnum).MaxX
    Buffer.WriteByte Map(mapnum).MaxY

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            With Map(mapnum).Tile(x, y)
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).x
                    Buffer.WriteLong .Layer(i).y
                    Buffer.WriteLong .Layer(i).Tileset
                Next
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteByte .DirBlock
            End With

        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(mapnum).NPC(x)
    Next
    
    For x = 1 To MAX_MAP_NPCS
        Buffer.WriteByte Map(mapnum).NPCSProperties(x).Movement
        Buffer.WriteByte Map(mapnum).NPCSProperties(x).Action
    Next
    
    Buffer.WriteLong Map(mapnum).Weather

    MapCache(mapnum).Data = Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        s = "There are no more players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "Have " & n & " Online players: " & s & "."
    End If

    Call PlayerMsg(index, s, WhoColor)
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
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteLong GetPlayerClass(index)
    Buffer.WriteLong GetPlayerVisible(index)
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(index, i)
    Next
    
    If Player(index).GuildFileId > 0 Then
        If TempPlayer(index).tmpGuildSlot > 0 Then
            Buffer.WriteByte 1
            Buffer.WriteString GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name
            Buffer.WriteLong Player(index).GuildMemberId
        End If
    Else
        Buffer.WriteByte 0
    End If
    
    'Triforces
    For i = 1 To TriforceType.TriforceType_Count - 1
        Buffer.WriteByte Player(index).triforce(i)
    Next
    
    'Send doors
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
    
    Call SendDone(index)
    
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

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(mapnum, i).playerName
        Buffer.WriteLong MapItem(mapnum, i).Num
        Buffer.WriteLong MapItem(mapnum, i).Value
        Buffer.WriteLong MapItem(mapnum, i).x
        Buffer.WriteLong MapItem(mapnum, i).y
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

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(mapnum, i).playerName
        Buffer.WriteLong MapItem(mapnum, i).Num
        Buffer.WriteLong MapItem(mapnum, i).Value
        Buffer.WriteLong MapItem(mapnum, i).x
        Buffer.WriteLong MapItem(mapnum, i).y
    Next

    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal mapnum As Long, ByVal MapNPCNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcVitals
    Buffer.WriteLong MapNPCNum
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong mapnpc(mapnum).NPC(MapNPCNum).Vital(i)
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

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong mapnpc(mapnum).NPC(i).Num
        Buffer.WriteLong mapnpc(mapnum).NPC(i).x
        Buffer.WriteLong mapnpc(mapnum).NPC(i).y
        Buffer.WriteLong mapnpc(mapnum).NPC(i).Dir
        Buffer.WriteLong mapnpc(mapnum).NPC(i).Vital(HP)
        Buffer.WriteByte mapnpc(mapnum).NPC(i).IsPet
        Buffer.WriteLong mapnpc(mapnum).NPC(i).PetData.owner
    Next

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong mapnpc(mapnum).NPC(i).Num
        Buffer.WriteLong mapnpc(mapnum).NPC(i).x
        Buffer.WriteLong mapnpc(mapnum).NPC(i).y
        Buffer.WriteLong mapnpc(mapnum).NPC(i).Dir
        Buffer.WriteLong mapnpc(mapnum).NPC(i).Vital(HP)
        Buffer.WriteByte mapnpc(mapnum).NPC(i).IsPet
        Buffer.WriteLong mapnpc(mapnum).NPC(i).PetData.owner
    Next

    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendItems(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(index, i)
        End If

    Next

End Sub

Sub SendAnimations(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(index, i)
        End If

    Next

End Sub

Sub SendNpcs(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS

        If LenB(Trim$(NPC(i).Name)) > 0 Then
            Call SendUpdateNpcTo(index, i)
        End If

    Next

End Sub

Sub SendResources(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).Name)) > 0 Then
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
    Buffer.WriteLong GetPlayerEquipment(index, Helmet)
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
    Buffer.WriteLong GetPlayerEquipment(index, Helmet)
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
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Helmet)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Shield)
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendVital(ByVal index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.HP)
        Case mp
            Buffer.WriteLong SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.mp)
            Buffer.WriteLong GetPlayerVital(index, Vitals.mp)
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

Sub SendWelcome(ByVal index As Long)
Dim RGB As Integer
' Send visibility message
If GetPlayerAccess(index) > ADMIN_MONITOR Then
If GetPlayerVisible(index) = 1 Then
Call PlayerMsg(index, "[INVISIBLE]", AlertColor)
End If
End If
    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(index, Options.MOTD, BrightBlue)
    End If

    ' Send whos online
    Call SendWhosOnline(index)
End Sub

Sub SendClasses(ByVal index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClassesData
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.mp)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        
        ' send array size
        Buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        
        ' send array size
        Buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
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
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewCharClasses
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.mp)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
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
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXYToMap(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXYMap
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
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
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    SendDataTo index, Buffer.ToArray()
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
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal npcnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set Buffer = New clsBuffer
    NPCSize = LenB(NPC(npcnum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(npcnum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong npcnum
    Buffer.WriteBytes NPCData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal npcnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set Buffer = New clsBuffer
    NPCSize = LenB(NPC(npcnum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(npcnum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong npcnum
    Buffer.WriteBytes NPCData
    SendDataTo index, Buffer.ToArray()
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
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendShops(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(index, i)
        End If

    Next

End Sub

Sub SendUpdateShopToAll(ByVal shopNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal shopNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpells(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(index, i)
        End If

    Next

End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes SpellData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes SpellData
    
    SendDataTo index, Buffer.ToArray()
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

Sub SendResourceCacheTo(ByVal index As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(GetPlayerMap(index)).Resource_Count

    If ResourceCache(GetPlayerMap(index)).Resource_Count > 0 Then

        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            Buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(i).ResourceState
            Buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(i).x
            Buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(i).y
        Next

    End If

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal mapnum As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(mapnum).Resource_Count

    If ResourceCache(mapnum).Resource_Count > 0 Then

        For i = 0 To ResourceCache(mapnum).Resource_Count
            Buffer.WriteByte ResourceCache(mapnum).ResourceData(i).ResourceState
            Buffer.WriteLong ResourceCache(mapnum).ResourceData(i).x
            Buffer.WriteLong ResourceCache(mapnum).ResourceData(i).y
        Next

    End If

    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDoorAnimation(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SDoorAnimation
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendActionMsg(ByVal mapnum As Long, ByVal Message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString Message
    Buffer.WriteLong Color
    Buffer.WriteLong MsgType
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap mapnum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendBlood(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBlood
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendAnimation(ByVal mapnum As Long, ByVal Anim As Long, ByVal x As Long, ByVal y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimation
    Buffer.WriteLong Anim
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte LockType
    Buffer.WriteLong LockIndex
    
    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCooldown(ByVal index As Long, ByVal Slot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong Slot
    
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

Sub SayMsg_Map(ByVal mapnum As Long, ByVal index As Long, ByVal Message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString Message
    Buffer.WriteString "[Map]"
    Buffer.WriteLong saycolour
    
    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal index As Long, ByVal Message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString Message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong saycolour
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub ResetShopAction(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendStunned(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    Buffer.WriteLong TempPlayer(index).StunDuration
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendBank(ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        Buffer.WriteLong Bank(index).Item(i).Num
        Buffer.WriteLong Bank(index).Item(i).Value
    Next
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKey(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte Value
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKeyToMap(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte Value
    SendDataToMap mapnum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendOpenShop(ByVal index As Long, ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    Buffer.WriteLong shopNum
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal index As Long, ByVal Movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
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
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).price * TempPlayer(index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).price
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
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).price
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
    Buffer.WriteLong TempPlayer(index).target
    Buffer.WriteLong TempPlayer(index).targetType
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHotbar(ByVal index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        Buffer.WriteLong Player(index).Hotbar(i).Slot
        Buffer.WriteByte Player(index).Hotbar(i).sType
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

Sub SendPlayerSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSoundToMap(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer
   
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub



Sub SendTradeRequest(ByVal index As Long, ByVal TradeRequest As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    Buffer.WriteString Trim$(Player(TradeRequest).Name)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal index As Long, ByVal targetPlayer As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    Buffer.WriteString Trim$(Player(targetPlayer).Name)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal partyNum As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    Buffer.WriteByte 1
    Buffer.WriteLong Party(partyNum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(partyNum).Member(i)
    Next
    Buffer.WriteLong Party(partyNum).MemberCount
    SendDataToParty partyNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal index As Long)
Dim Buffer As clsBuffer, i As Long, partyNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(partyNum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(partyNum).Member(i)
        Next
        Buffer.WriteLong Party(partyNum).MemberCount
    Else
        ' send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal partyNum As Long, ByVal index As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    Buffer.WriteLong index
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(index, i)
        Buffer.WriteLong Player(index).Vital(i)
    Next
    SendDataToParty partyNum, Buffer.ToArray()
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
    Buffer.WriteLong MapItem(mapnum, index).x
    Buffer.WriteLong MapItem(mapnum, index).y
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
        Buffer.WriteLong .Range
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


Sub SendUpdateDoorToAll(ByVal doornum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim DoorSize As Long
    Dim DoorData() As Byte
    SetStatus ("sending doors")
    Set Buffer = New clsBuffer
    
    DoorSize = LenB(Doors(doornum))
    ReDim DoorData(DoorSize - 1)
    CopyMemory DoorData(0), ByVal VarPtr(Doors(doornum)), DoorSize
    
    Buffer.WriteLong SUpdateDoors
    Buffer.WriteLong doornum
    Buffer.WriteBytes DoorData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateDoorsTo(ByVal index As Long, ByVal doornum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim DoorSize As Long
    Dim DoorData() As Byte
    'SetStatus ("sending doors")
    Set Buffer = New clsBuffer
    
    DoorSize = LenB(Doors(doornum))
    ReDim DoorData(DoorSize - 1)
    CopyMemory DoorData(0), ByVal VarPtr(Doors(doornum)), DoorSize
    
    Buffer.WriteLong SUpdateDoors
    Buffer.WriteLong doornum
    Buffer.WriteBytes DoorData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub PartyChatMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim i As Long
Dim Member As Integer
Dim partyNum As Long

partyNum = TempPlayer(index).inParty

    ' not in a party?
    If TempPlayer(index).inParty = 0 Then
        Call PlayerMsg(index, "You are not in a group.", BrightRed)
        Exit Sub
    End If

    For i = 1 To MAX_PARTY_MEMBERS
        Member = Party(partyNum).Member(i)
        ' is online, does exist?
        If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
        ' yep, send the message!
            Call PlayerMsg(Member, "[Party] " & GetPlayerName(index) & ": " & Msg, Color)
        End If
    Next
End Sub

Sub SendChatBubble(ByVal mapnum As Long, ByVal target As Long, ByVal targetType As Long, ByVal Message As String, ByVal Colour As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SChatBubble
    Buffer.WriteLong target
    Buffer.WriteLong targetType
    Buffer.WriteString Message
    Buffer.WriteLong Colour
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
    TempPlayer(index).IsLoading = False
End Sub

Sub SendWeather()
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer
Buffer.WriteLong SSendWeather
If Rainon = True Then
     Buffer.WriteLong 1
Else
     Buffer.WriteLong 0
End If

SendDataToAll Buffer.ToArray()
Set Buffer = Nothing

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
    SetStatus ("sending movements")
    Set Buffer = New clsBuffer
    Dim i As Byte
    
    
    Buffer.WriteLong SUpdateMovements
    Buffer.WriteLong MovementNum
    'Buffer.WriteBytes movementData
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
        
    End With
    

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateMovementsTo(ByVal index As Long, ByVal MovementNum As Long)
    Dim Buffer As clsBuffer
    'SetStatus ("sending movements")
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
    SetStatus ("sending Actions")
    Set Buffer = New clsBuffer
    Dim i As Byte

    
    Buffer.WriteLong SUpdateActions
    Buffer.WriteLong ActionNum
    With Actions(ActionNum)
        
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
    'SetStatus ("sending Actions")
    Set Buffer = New clsBuffer
    Dim i As Byte
    
    
    Buffer.WriteLong SUpdateActions
    Buffer.WriteLong ActionNum
    With Actions(ActionNum)
        
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

Sub SendPlayerDoor(ByVal index As Long, ByVal doornum As Byte)

    Dim packet As String
    Dim Buffer As clsBuffer
    SetStatus ("sending Actions")
    Set Buffer = New clsBuffer

    
    If Not (doornum > 0 And doornum <= MAX_DOORS) Then Exit Sub
     
    Buffer.WriteLong SUpdatePlayerDoor
    Buffer.WriteByte doornum
    Buffer.WriteByte Player(index).PlayerDoors(doornum).state
    
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
    Buffer.WriteLong Map(mapnum).NPC(npcnum)
    Buffer.WriteLong mapnpc(mapnum).NPC(npcnum).Num
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Map = mapnum Then
                SendDataTo i, Buffer.ToArray()
            End If
        End If
    Next
End Sub

Sub NPCCache_Create(ByVal index As Long, ByVal mapnum As Long, ByVal npcnum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SNPCCache
    Buffer.WriteLong mapnum
    Buffer.WriteLong npcnum
    Buffer.WriteLong Map(mapnum).NPC(npcnum)
    Buffer.WriteLong mapnpc(mapnum).NPC(npcnum).Num
    
    SendDataTo index, Buffer.ToArray()
End Sub


Sub SendPets(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_PETS

        If LenB(Trim$(Pet(i).Name)) > 0 Then
            Call SendUpdatePetsTo(index, i)
        End If

    Next

End Sub

Sub SendUpdatePetToAll(ByVal PetNum As Long)
    Dim Buffer As clsBuffer
    SetStatus ("sending Pets")
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
    'SetStatus ("sending Pets")
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
    SetStatus ("sending Pets")
    Set Buffer = New clsBuffer
    
    If Player(index).Pet(PlayerPetSlot).NumPet < 1 Or Player(index).Pet(PlayerPetSlot).NumPet > MAX_PETS Then GoTo SendNullData
    If Pet(Player(index).Pet(PlayerPetSlot).NumPet).npcnum < 1 Or Pet(Player(index).Pet(PlayerPetSlot).NumPet).npcnum > MAX_NPCS Then GoTo SendNullData
    
    Buffer.WriteLong SPetData
    Buffer.WriteByte PlayerPetSlot
    
    Buffer.WriteString NPC(Pet(Player(index).Pet(PlayerPetSlot).NumPet).npcnum).Name
    Buffer.WriteInteger Player(index).Pet(PlayerPetSlot).points
    Buffer.WriteLong Player(index).Pet(PlayerPetSlot).Experience
    Buffer.WriteLong Player(index).Pet(PlayerPetSlot).Level
    Buffer.WriteByte Player(index).Pet(PlayerPetSlot).NumPet
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteByte Player(index).Pet(PlayerPetSlot).StatsAdd(i)
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


    

