Attribute VB_Name = "modGameLogic"
Option Explicit

Private Type trio
    first As Byte
    Second As Long
    third As Long
End Type

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal mapnum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(mapnum, i).Num = 0 And Not IsItemMapped(mapnum, i) Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long
    
    Name = Trim$(Name)

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            'If Name & "*" Like GetPlayerName(i) Then Exit Function
            If UCase$(left$(GetPlayerName(i), 1)) = UCase$(left$(Name, 1)) Then 'optimizing
                ' Make sure we dont try to check a name thats to small
                If UCase$(GetPlayerName(i)) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
                'test short len
                Dim lenght As Integer
                lenght = Len(Name)
                If lenght > 0 Then
                    If UCase$(left$(GetPlayerName(i), lenght)) = UCase$(Name) Then
                        FindPlayer = i
                        Exit Function
                    End If
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal itemval As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal isDrop As Boolean = True)
    Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(mapnum)
    Call SpawnItemSlot(i, ItemNum, itemval, mapnum, X, Y, playerName, isDrop)
    CheckMapItemHighIndex mapnum, i, True
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal itemval As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal isDrop As Boolean = True)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
                 
            MapItem(mapnum, i).playerName = playerName
            MapItem(mapnum, i).playerTimer = GetRealTickCount + ITEM_SPAWN_TIME
            MapItem(mapnum, i).isDrop = isDrop
            MapItem(mapnum, i).Timer = GetRealTickCount + ITEM_DESPAWN_TIME
            MapItem(mapnum, i).Num = ItemNum
            MapItem(mapnum, i).Value = itemval
            MapItem(mapnum, i).X = X
            MapItem(mapnum, i).Y = Y
            ' send to map
            SendSpawnItemToMap mapnum, i
            
            'check upper bound
            CheckMapItemHighIndex mapnum, i, ItemNum > 0
            
            
        End If
    End If

    Set Buffer = Nothing
End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal mapnum As Long)
    Dim X As Long
    Dim Y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For X = 0 To map(mapnum).MaxX
        For Y = 0 To map(mapnum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (map(mapnum).Tile(X, Y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If isItemStackable(map(mapnum).Tile(X, Y).Data1) And map(mapnum).Tile(X, Y).Data2 <= 0 Then
                    Call SpawnItem(map(mapnum).Tile(X, Y).Data1, 1, mapnum, X, Y, , False)
                Else
                    Call SpawnItem(map(mapnum).Tile(X, Y).Data1, map(mapnum).Tile(X, Y).Data2, mapnum, X, Y, , False)
                End If
            End If

        Next
    Next

End Sub

Public Function Random(ByVal Low As Long, ByVal high As Long) As Long
    Random = ((high - Low + 1) * Rnd) + Low
End Function

'here
Public Sub SpawnNpc(ByVal mapnpcnum As Long, ByVal mapnum As Long, Optional ByVal SetX As Long, Optional ByVal SetY As Long, Optional ByVal npcnum As Long = 0, Optional nearOwner As Boolean = False)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim spawned As Boolean
    Dim PetOwner As Long
    Dim HPSpawn As Long, MPSpawn As Long

    ' Check for subscript out of range
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Sub
    
    If npcnum = 0 Then
        npcnum = map(mapnum).NPC(mapnpcnum)
    End If

    If npcnum > 0 Then
    
        MapNpc(mapnum).NPC(mapnpcnum).Num = npcnum
        MapNpc(mapnum).NPC(mapnpcnum).mapnpcnum = mapnpcnum
        MapNpc(mapnum).NPC(mapnpcnum).Target = 0
        MapNpc(mapnum).NPC(mapnpcnum).TargetType = 0 ' clear
        
        
        'Pet Hp Recovery
        PetOwner = GetMapPetOwner(mapnum, mapnpcnum)
        
        If PetOwner > 0 Then
            For i = 1 To Vitals.Vital_Count - 1
                If player(PetOwner).Pet(TempPlayer(PetOwner).TempPet.ActualPet).CurVital(i) > 0 Then
                    MapNpc(mapnum).NPC(mapnpcnum).vital(i) = player(PetOwner).Pet(TempPlayer(PetOwner).TempPet.ActualPet).CurVital(i)
                Else
                    MapNpc(mapnum).NPC(mapnpcnum).vital(i) = GetNpcMaxVital(mapnum, mapnpcnum, i)
                End If
            Next
        Else
        'here
            For i = 1 To Vitals.Vital_Count - 1
                MapNpc(mapnum).NPC(mapnpcnum).vital(i) = GetNpcMaxVital(mapnum, mapnpcnum, i)
            Next
        End If
        
        MapNpc(mapnum).NPC(mapnpcnum).dir = Int(Rnd * 4)
        
    If PetOwner > 0 Then
        X = SetX
        Y = SetY
        If nearOwner = True Then
            For i = 1 To 50
            'we should have had x/y passed.

                If X > map(mapnum).MaxX Then X = map(mapnum).MaxX
                If Y > map(mapnum).MaxY Then Y = map(mapnum).MaxY
                If X <= 0 Then X = 1
                If Y <= 0 Then Y = 1

                'possibly try to look in a circle/square around the player
                'to figure out where we are able to attempt a spawn..
                'I'll need to do some proper math for that :o
                If Not NpcTileIsOpen(mapnum, X, Y) Then
                X = SetX
                Y = SetY
                    Select Case GetPlayerDir(PetOwner)
                        Case DIR_UP
                            Y = Y + Random(-5, 1)
                            X = X + Random(-5, 5)
                        Case DIR_DOWN
                            Y = Y - Random(-5, 1)
                            X = X + Random(-5, 5)
                        Case DIR_LEFT
                            X = X + Random(-5, 1)
                            Y = Y - Random(-5, 5)
                        Case DIR_RIGHT
                            X = X - Random(-5, 1)
                            Y = Y - Random(-5, 5)
                    End Select
                Else
                    MapNpc(mapnum).NPC(mapnpcnum).X = X
                    MapNpc(mapnum).NPC(mapnpcnum).Y = Y
                    spawned = True
                    Exit For
                End If
            Next
        End If
    End If
    
    'If spawned = True Then
    '    If PetOwner > 0 Then
    '        PlayerMsg PetOwner, "Your pet has followed you.", White, , False
    '    End If
    'End If
    
        If Not spawned Then
        For i = 1 To TempTile(mapnum).NumSpawnSites
            With TempTile(mapnum).NPCSpawnSite(i)
            If map(mapnum).Tile(.X, .Y).Data1 = mapnpcnum Then
                MapNpc(mapnum).NPC(mapnpcnum).X = .X
                MapNpc(mapnum).NPC(mapnpcnum).Y = .Y
                MapNpc(mapnum).NPC(mapnpcnum).dir = map(mapnum).Tile(.X, .Y).Data2
                spawned = True
                Exit For
            End If
            End With
        Next
        'Check if theres a spawn tile for the specific npc
        'For X = 0 To map(mapnum).MaxX
            'For Y = 0 To map(mapnum).MaxY
                'If map(mapnum).Tile(X, Y).Type = TILE_TYPE_NPCSPAWN Then
                    'If map(mapnum).Tile(X, Y).Data1 = mapnpcnum Then
                        'mapnpc(mapnum).NPC(mapnpcnum).X = X
                        'mapnpc(mapnum).NPC(mapnpcnum).Y = Y
                        'mapnpc(mapnum).NPC(mapnpcnum).dir = map(mapnum).Tile(X, Y).Data2
                        'spawned = True
                        'Exit For
                    'End If
                'End If
            'Next Y
        'Next X
        End If
        
        If Not spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                
                If SetX = 0 And SetY = 0 Then
                    X = Random(0, map(mapnum).MaxX)
                    Y = Random(0, map(mapnum).MaxY)
                Else
                    X = SetX
                    Y = SetY
                End If
    
                If X > map(mapnum).MaxX Then X = map(mapnum).MaxX
                If Y > map(mapnum).MaxY Then Y = map(mapnum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(mapnum, X, Y) Then
                    MapNpc(mapnum).NPC(mapnpcnum).X = X
                    MapNpc(mapnum).NPC(mapnpcnum).Y = Y
                    spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not spawned Then

            For X = 0 To map(mapnum).MaxX
                For Y = 0 To map(mapnum).MaxY

                    If NpcTileIsOpen(mapnum, X, Y) Then
                        MapNpc(mapnum).NPC(mapnpcnum).X = X
                        MapNpc(mapnum).NPC(mapnpcnum).Y = Y
                        spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong mapnpcnum
            Buffer.WriteLong MapNpc(mapnum).NPC(mapnpcnum).Num
            Buffer.WriteLong MapNpc(mapnum).NPC(mapnpcnum).X
            Buffer.WriteLong MapNpc(mapnum).NPC(mapnpcnum).Y
            Buffer.WriteLong MapNpc(mapnum).NPC(mapnpcnum).dir
            Buffer.WriteLong MapNpc(mapnum).NPC(mapnpcnum).PetData.Owner
            SendDataToMap mapnum, Buffer.ToArray()
            
            
            Call ResetMapNPCMovement(mapnum, mapnpcnum)
            
            If mapnpcnum > TempMap(mapnum).npc_highindex Then
                Call SetMapNPCHighIndex(mapnum, mapnpcnum)
            End If
            
            AddNPCToMapRef mapnum, mapnpcnum
        End If
        
        SendMapNpcVitals mapnum, mapnpcnum
    End If
    
    Set Buffer = Nothing

End Sub

Public Function NpcTileIsOpen(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If ArePlayersOnMap(mapnum) > 0 Then
        If FindPlayerByPos(mapnum, X, Y) > 0 Then
            NpcTileIsOpen = False
            Exit Function
        End If
    End If
        'For LoopI = 1 To Player_HighIndex

            'If GetPlayerMap(LoopI) = mapnum Then
                'If GetPlayerX(LoopI) = x Then
                    'If GetPlayerY(LoopI) = y Then
                        'NpcTileIsOpen = False
                        'Exit Function
                    'End If
                'End If
            'End If

        'Next

    'End If

    LoopI = GetMapNPCNumByTile(mapnum, X, Y)
    If LoopI > 0 Then
        NpcTileIsOpen = False
        Exit Function
    End If
    'For LoopI = 1 To MAX_MAP_NPCS

        'If mapnpc(mapnum).NPC(LoopI).Num > 0 Then
            'If mapnpc(mapnum).NPC(LoopI).x = x Then
                'If mapnpc(mapnum).NPC(LoopI).y = y Then
                    'NpcTileIsOpen = False
                    'Exit Function
                'End If
            'End If
        'End If

    'Next

    If map(mapnum).Tile(X, Y).Type <> TILE_TYPE_WALKABLE Then
        If map(mapnum).Tile(X, Y).Type <> TILE_TYPE_NPCSPAWN Then
            If map(mapnum).Tile(X, Y).Type <> TILE_TYPE_ITEM Then
                    NpcTileIsOpen = False
            End If
        End If
    End If
    
    If map(mapnum).Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        If isWalkableResource(mapnum, X, Y) Then
            NpcTileIsOpen = True
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal mapnum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, mapnum)
    Next
    


End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub



Function GetTotalMapPlayers(ByVal mapnum As Long) As Long
    Dim i As Long
    Dim N As Long
    N = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = mapnum Then
            N = N + 1
        End If

    Next

    GetTotalMapPlayers = N
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal mapnum As Long)
    TempTile(mapnum).NumDoors = 0
    ReDim TempTile(mapnum).Door(1 To 1)
End Sub

Sub InitTempTiles()
    Dim i As Long
    For i = 1 To MAX_MAPS
        InitTempTile (i)
    Next
End Sub

Public Sub InitTempTile(ByVal mapnum As Long)
    Dim Y As Long
    Dim X As Long
    For X = 0 To map(mapnum).MaxX
        For Y = 0 To map(mapnum).MaxY
            If map(mapnum).Tile(X, Y).Type = TILE_TYPE_DOOR Then
                TempTile(mapnum).NumDoors = TempTile(mapnum).NumDoors + 1
                ReDim Preserve TempTile(mapnum).Door(1 To TempTile(mapnum).NumDoors)
                With TempTile(mapnum).Door(TempTile(mapnum).NumDoors)
                    .DoorNum = map(mapnum).Tile(X, Y).Data1
                    .DoorTimer = 0
                    .state = GetInitialDoorState(.DoorNum)
                    .X = X
                    .Y = Y
                End With
            ElseIf map(mapnum).Tile(X, Y).Type = TILE_TYPE_KEY Then
                TempTile(mapnum).NumDoors = TempTile(mapnum).NumDoors + 1
                ReDim Preserve TempTile(mapnum).Door(1 To TempTile(mapnum).NumDoors)
                With TempTile(mapnum).Door(TempTile(mapnum).NumDoors)
                    .DoorNum = -1 'use this encode
                    .DoorTimer = 0
                    .state = False
                    .X = X
                    .Y = Y
                End With
            ElseIf map(mapnum).Tile(X, Y).Type = TILE_TYPE_NPCSPAWN Then
                TempTile(mapnum).NumSpawnSites = TempTile(mapnum).NumSpawnSites + 1
                ReDim Preserve TempTile(mapnum).NPCSpawnSite(1 To TempTile(mapnum).NumSpawnSites)
                With TempTile(mapnum).NPCSpawnSite(TempTile(mapnum).NumSpawnSites)
                .X = X
                .Y = Y
                End With
            End If
        Next
    Next

End Sub

Public Sub CacheResources(ByVal mapnum As Long)
    Dim X As Long, Y As Long, Resource_Count As Long
    Resource_Count = 0

    For X = 0 To map(mapnum).MaxX
        For Y = 0 To map(mapnum).MaxY

            If map(mapnum).Tile(X, Y).Type = TILE_TYPE_RESOURCE And map(mapnum).Tile(X, Y).Data1 > 0 Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mapnum).ResourceData(0 To Resource_Count)
                ResourceCache(mapnum).ResourceData(Resource_Count).X = X
                ResourceCache(mapnum).ResourceData(Resource_Count).Y = Y
                ResourceCache(mapnum).ResourceData(Resource_Count).cur_health = Resource(map(mapnum).Tile(X, Y).Data1).health
            End If

        Next
    Next

    ResourceCache(mapnum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(index, oldSlot)
    OldValue = GetPlayerBankItemValue(index, oldSlot)
    NewNum = GetPlayerBankItemNum(index, newSlot)
    NewValue = GetPlayerBankItemValue(index, newSlot)
    
    SetPlayerBankItemNum index, newSlot, OldNum
    SetPlayerBankItemValue index, newSlot, OldValue
    
    SetPlayerBankItemNum index, oldSlot, NewNum
    SetPlayerBankItemValue index, oldSlot, NewValue
        
    SendBank index
End Sub

Sub PlayerSwitchInvSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(index, oldSlot)
    OldValue = GetPlayerInvItemValue(index, oldSlot)
    NewNum = GetPlayerInvItemNum(index, newSlot)
    NewValue = GetPlayerInvItemValue(index, newSlot)
    SetPlayerInvItemNum index, newSlot, OldNum
    SetPlayerInvItemValue index, newSlot, OldValue
    SetPlayerInvItemNum index, oldSlot, NewNum
    SetPlayerInvItemValue index, oldSlot, NewValue
    SendInventory index
End Sub

Sub PlayerSwitchSpellSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(index, oldSlot)
    NewNum = GetPlayerSpell(index, newSlot)
    SetPlayerSpell index, oldSlot, NewNum
    SetPlayerSpell index, newSlot, OldNum
    SendPlayerSpells index
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    Dim i As Long
    i = FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot))
    If i > 0 Then
        SwapInvEquipment index, i, EqSlot
        SendInventoryUpdate index, i
        SendEquipmentUpdate index
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, EqSlot)
    Else
        PlayerMsg index, "Tu inventario está lleno.", BrightRed
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(left$(Word, 1))
   
    'Word = GetTranslation(Word)
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function IsinRange(ByVal range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    IsinRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= range Then IsinRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef dir As Byte) As Boolean
    If Not blockvar And (2 ^ dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function RAND(ByVal Low As Long, ByVal high As Long) As Long
    Randomize
    RAND = Int((high - Low + 1) * Rnd) + Low
End Function

Public Function RAND2(ByVal side1 As Long, ByVal side2 As Long) As Long
    If Rnd < 0.5 Then
        RAND2 = side1
    Else
        RAND2 = side2
    End If
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal index As Long)
Dim partynum As Long, i As Long

    partynum = TempPlayer(index).inParty
    If partynum > 0 Then
        ' find out how many members we have
        Party_CountMembers partynum
        ' make sure there's more than 2 people
        If Party(partynum).MemberCount > 2 Then
            ' check if leader
            If Party(partynum).Leader = index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(i) > 0 And Party(partynum).Member(i) <> index Then
                        Party(partynum).Leader = Party(partynum).Member(i)
                        PartyMsg partynum, GetPlayerName(i) & " " & GetTranslation("es ahora el líder del equipo"), Cyan
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partynum, GetPlayerName(Party(partynum).Leader) & " " & GetTranslation("ha dejado el equipo"), BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(i) = index Then
                        Party(partynum).Member(i) = 0
                        Exit For
                    End If
                Next
                
                TempPlayer(index).inParty = 0
                TempPlayer(index).partyInvite = 0
                
                ' recount party
                Party_CountMembers partynum
                ' set update to all
                SendPartyUpdate partynum
                ' send clear to player
                SendPartyUpdateTo index
            Else
                ' not the leader, just leave
                PartyMsg partynum, GetPlayerName(index) & GetTranslation("ha dejado el equipo"), BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(i) = index Then
                        Party(partynum).Member(i) = 0
                        Exit For
                    End If
                Next
                
                TempPlayer(index).inParty = 0
                TempPlayer(index).partyInvite = 0
                
                ' recount party
                Party_CountMembers partynum
                ' set update to all
                SendPartyUpdate partynum
                ' send clear to player
                SendPartyUpdateTo index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partynum
            ' only 2 people, disband
            PartyMsg partynum, GetTranslation("Equipo disuelto."), BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                index = Party(partynum).Member(i)
                ' player exist?
                If index > 0 Then
                    ' remove them
                    TempPlayer(index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo index
                End If
            Next
            ' clear out the party itself
            ClearParty partynum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal index As Long, ByVal targetPlayer As Long)
Dim partynum As Long, i As Long

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Or TempPlayer(targetPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg index, "El jugador está ocupado.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg index, "Éste jugador ya está en el equipo.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(index).inParty > 0 Then
        partynum = TempPlayer(index).inParty
        ' make sure we're the leader
        If Party(partynum).Leader = index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partynum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = index
                    ' let them know
                    PlayerMsg index, "Invitación enviada.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg index, "El equipo está lleno.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg index, "No eres el líder del equipo.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = index
        ' let them know
        PlayerMsg index, "Invitación enviada.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal index As Long, ByVal targetPlayer As Long)
Dim partynum As Long, i As Long, X As Long

    ' check if already in a party
    If TempPlayer(index).inParty > 0 Then
        ' get the partynumber
        partynum = TempPlayer(index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partynum).Member(i) = 0 Then
                'add to the party
                Party(partynum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partynum
                ' send update to all - including new player
                SendPartyUpdate partynum
                For X = 1 To MAX_PARTY_MEMBERS
                If Party(partynum).Member(X) > 0 Then
                    SendPartyVitals partynum, Party(partynum).Member(X)
                End If
            Next
                ' let everyone know they've joined
                PartyMsg partynum, GetPlayerName(targetPlayer) & " " & GetTranslation("se ha unido al equipo."), Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partynum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg index, "El equipo está lleno.", BrightRed
        PlayerMsg targetPlayer, "El equipo está lleno.", BrightRed
        TempPlayer(targetPlayer).partyInvite = 0
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partynum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partynum).MemberCount = 2
        Party(partynum).Leader = index
        Party(partynum).Member(1) = index
        Party(partynum).Member(2) = targetPlayer
        SendPartyUpdate partynum
        SendPartyVitals partynum, index
        SendPartyVitals partynum, targetPlayer
        ' let them know it's created
        PartyMsg partynum, GetTranslation("Equipo creado."), BrightGreen
        PartyMsg partynum, GetPlayerName(index) & " " & GetTranslation("se ha unido al equipo."), Pink, False
        PartyMsg partynum, GetPlayerName(targetPlayer) & " " & GetTranslation("se ha unido al equipo."), Pink, False
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(index).inParty = partynum
        TempPlayer(targetPlayer).inParty = partynum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal index As Long, ByVal targetPlayer As Long)
    PlayerMsg index, GetPlayerName(targetPlayer) & " " & GetTranslation("ha rechazado unirse al equipo."), BrightRed, , False
    PlayerMsg targetPlayer, "Has rechazado unirte al equipo.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partynum As Long)
Dim i As Long, highIndex As Long, X As Long
    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partynum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partynum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For X = i To MAX_PARTY_MEMBERS - 1
                    Party(partynum).Member(X) = Party(partynum).Member(X + 1)
                    Party(partynum).Member(X + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partynum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partynum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partynum
End Sub

Public Sub CheckPlayerPartyTasks(ByVal index As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)

CheckTasks index, TaskType, TargetIndex

Dim partynum As Byte
partynum = TempPlayer(index).inParty

If partynum > 0 Then
    Dim i As Byte
    For i = 1 To MAX_PARTY_MEMBERS
        Dim tmpIndex As Long
        tmpIndex = Party(partynum).Member(i)
        If IsPlaying(tmpIndex) Then
        If tmpIndex <> index Then
            If GetPlayerMap(tmpIndex) = GetPlayerMap(index) Then
                Dim lvldiff As Integer
                lvldiff = GetLevelDifference(index, tmpIndex)
                If lvldiff <= 10 Then
                    CheckTasks tmpIndex, TaskType, TargetIndex
                End If
            End If
        End If
        End If
    Next
End If

End Sub

Public Sub HandleProjecTile(ByVal index As Long, ByVal PlayerProjectile As Long)
Dim X As Long, Y As Long, i As Long
Dim Damage As Long
Dim npcnum As Long
Dim mapnum As Long
Dim blockAmount As Long

Damage = 0

    ' check for subscript out of range
    If index < 1 Or index > MAX_PLAYERS Or PlayerProjectile < 1 Or PlayerProjectile > MAX_PLAYER_PROJECTILES Then Exit Sub
        
    ' check to see if it's time to move the Projectile
    If GetRealTickCount > TempPlayer(index).ProjecTile(PlayerProjectile).TravelTime Then
        With TempPlayer(index).ProjecTile(PlayerProjectile)
            ' set next travel time and the current position and then set the actual direction based on RMXP arrow tiles.
            Select Case .Direction
                ' down
                Case DIR_DOWN
                    .Y = .Y + 1
                    ' check if they reached maxrange
                    If .Y = (GetPlayerY(index) + .range) + 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
                ' up
                Case DIR_UP
                    .Y = .Y - 1
                    ' check if they reached maxrange
                    If .Y = (GetPlayerY(index) - .range) - 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
                ' right
                Case DIR_RIGHT
                    .X = .X + 1
                    ' check if they reached max range
                    If .X = (GetPlayerX(index) + .range) + 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
                ' left
                Case DIR_LEFT
                    .X = .X - 1
                    ' check if they reached maxrange
                    If .X = (GetPlayerX(index) - .range) - 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
            End Select
            .TravelTime = GetRealTickCount + .Speed
        End With
    Else
        Exit Sub
    End If
    
    X = TempPlayer(index).ProjecTile(PlayerProjectile).X
    Y = TempPlayer(index).ProjecTile(PlayerProjectile).Y
    
    ' check if left map
    If X > map(GetPlayerMap(index)).MaxX Or Y > map(GetPlayerMap(index)).MaxY Or X < 0 Or Y < 0 Then
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
    If TempPlayer(index).ProjecTile(PlayerProjectile).Depth = 0 Then
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
    mapnum = GetPlayerMap(index)
    
'Projectile scaling formula
'For i = 1 To Player_HighIndex
    'If i <> index Then
        'If IsPlaying(i) Then
            ' check coordinates
            'If GetPlayerMap(i) = mapnum Then
                'If x = Player(i).x And y = GetPlayerY(i) Then
                    ' check if player can attack
    i = FindPlayerByPos(mapnum, X, Y)
    If i > 0 And i <> index Then
        If CanPlayerAttackPlayer(index, i, True) = True Then
            ' attack the player and kill the project tile
            If CanPlayerDodge(i) Then
                SendActionMsg mapnum, "¡Esquivado!", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32), , True
                Exit Sub
            End If
        
            Damage = GetPlayerProjectileDamageAgainstPlayer(index, i)
            Damage = Damage - GetPlayerDefenseAgainstPlayer(index, Damage)
            
            If (Damage < 0) Then
                Damage = Damage * -1
            End If
            
            
            
            Damage = RAND(Damage / 2, Damage)
            
            ' * 1.5 if it's a crit!
            If CanPlayerCriticalHit(index) Then
                Damage = Damage * 1.5
                SendActionMsg mapnum, "¡Crítico!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32), , True
            End If
            If Damage > 0 Then
                Call PlayerAttackPlayer(index, i, Damage)
                Call Impactar(index, i, Damage, TempPlayer(index).TargetType)
            Else
                Call PlayerMsg(index, "¡No acertastes al objetivo!", BrightRed)
            End If
        Else
            'PlayerMsg index, "imposible atacar", BrightRed
        End If
        'tried to attack, clear projectile
        CheckClearProjectile index, PlayerProjectile
        Exit Sub
    End If


Dim npc_highindex As Long

npc_highindex = TempMap(mapnum).npc_highindex

i = GetMapRefNPCNumByTile(GetMapRef(mapnum), X, Y)
If i > 0 Then

' check for npc hit
'For i = 1 To npc_highindex
    npcnum = MapNpc(mapnum).NPC(i).Num
    If npcnum > 0 Then
        'If x = mapnpc(mapnum).NPC(i).x And y = mapnpc(mapnum).NPC(i).y Then
            ' they're hit, remove it and deal that damage ;)
            If CanPlayerAttackNpc(index, i, True) Then
                If CanNpcDodge(index, npcnum) Then
                    SendActionMsg mapnum, "¡Esquivado!", Pink, 1, (MapNpc(mapnum).NPC(i).X * 32), (MapNpc(mapnum).NPC(i).Y * 32), , True
                    Exit Sub
                End If
                If CanNpcParry(index, npcnum) Then
                    SendActionMsg mapnum, "¡Desviado!", Pink, 1, (MapNpc(mapnum).NPC(i).X * 32), (MapNpc(mapnum).NPC(i).Y * 32), , True
                    
                    Exit Sub
                End If


                Damage = GetPlayerProjectileDamageAgainstNPC(index, PlayerProjectile)
                Damage = Damage - GetNpcDefense(mapnum, i, Damage)
                Damage = RAND(Damage / 2, Damage)
                               
                ' * 1.5 if it's a crit!
                If CanPlayerCriticalHit(index) Then
                    Damage = Damage * 1.5
                    SendActionMsg mapnum, "¡Crítico!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32), , True
                End If
                If Damage > 0 Then
                    Call Impactar(index, i, Damage, TempPlayer(index).TargetType)
                    Call PlayerAttackNpc(index, i, Damage)
                Else
                    Call PlayerMsg(index, "¡No acertastes al objetivo!", BrightRed)
                End If
            End If
            'clear and exit
            CheckClearProjectile index, PlayerProjectile
            'ClearProjectile index, PlayerProjectile
            Exit Sub
        'End If
    End If
End If

' hit a block
If map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
    ' hit a block, clear it.
    ClearProjectile index, PlayerProjectile
    Exit Sub
End If

' hit a resource
If map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
    If CheckResource(index, X, Y) Then 'resource was cut
        If Not isWalkableResource(mapnum, X, Y) Then
            ' resource is not cut and can
            ClearProjectile index, PlayerProjectile
            Exit Sub
        Else
            CheckClearProjectile index, PlayerProjectile
            Exit Sub
        End If
    Else 'resource was not cut, so stop
        If Not isWalkableResource(mapnum, X, Y) Then
            ' resource is not cut and can
            ClearProjectile index, PlayerProjectile
            Exit Sub
        Else
            'CheckClearProjectile index, PlayerProjectile
            Exit Sub
        End If
    End If
End If

' hit a switch
If map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_DOOR Then
    ' hit a switch, clear it.
    CheckDoor index, X, Y
    CheckClearProjectile index, PlayerProjectile
    Exit Sub
End If
    
End Sub

Sub CheckClearProjectile(ByVal index As Long, ByVal PlayerProjectile As Long)
    If TempPlayer(index).ProjecTile(PlayerProjectile).Depth > 0 Then
        TempPlayer(index).ProjecTile(PlayerProjectile).Depth = TempPlayer(index).ProjecTile(PlayerProjectile).Depth - 1
        'Dim dir As Byte
        'dir = TempPlayer(index).ProjecTile(PlayerProjectile).Direction
        'GetNextPositionByRef dir, GetPlayerMap(index), TempPlayer(index).ProjecTile(PlayerProjectile).X, TempPlayer(index).ProjecTile(PlayerProjectile).Y
    End If
    
    If TempPlayer(index).ProjecTile(PlayerProjectile).Depth = 0 Then
        ClearProjectile index, PlayerProjectile
    End If
End Sub
Sub SpeechWindow(ByVal index As Long, ByVal msg As String, ByVal npcnum As Long)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong SSpeechWindow
Buffer.WriteString msg
Buffer.WriteLong npcnum
SendDataTo index, Buffer.ToArray()
Set Buffer = Nothing
End Sub


Public Sub ScriptTilePresses(ByVal index As Long, ByVal scriptnum As Long)
    If scriptnum > 0 Then
        Select Case scriptnum
        Case 1
            Call PlayerNavigation(index)
        Case 2
            Call RespawnMapSlideNPCS(GetPlayerMap(index))
            MapMsg GetPlayerMap(index), "Los bloques han vuelto a su posición", BrightGreen
        Case 3
            Call Exchange(index)
        Case 4 'event, temporal
        Case 5
            Call ComputePlayerOnFlag(index)
        Case 6
            Call SpecialLevelup(index)

        End Select
        'mensaje de prueba, borrar cuando convenga
        'Call PlayerMsg(index, "Has pisado el script tile n." & Scriptnum, 15)
    End If
    
End Sub

Public Sub ScriptTileLeave(ByVal index As Long, ByVal scriptnum As Long)
    If scriptnum > 0 Then
        Select Case scriptnum
        Case 1
            Call PlayerNavigation(index)
        End Select
    End If
    
End Sub

Sub SpecialLevelup(ByVal index As Long)
    Dim p As Variant
    Dim CanLevelUp As Boolean
    CanLevelUp = False
    For Each p In GetMapPlayerCollection(GetPlayerMap(index))
        If p <> index Then
            If GetPlayerAccess_Mode(p) > ADMIN_MAPPER Then
                If GetPlayerIP(p) = GetPlayerIP(index) Then
                    CanLevelUp = True
                End If
            End If
        End If
    Next
    If CanLevelUp Then
        GivePlayerEXP index, GetPlayerNextLevel(index) - GetPlayerExp(index)
    End If
End Sub

Public Sub UpdateWeather()
If Rainon = True Then
     Rainon = False
         Call SendWeathertoAll
Else
     Rainon = True
         Call SendWeathertoAll
End If
End Sub

Public Sub ActivateWeather()
    Rainon = True
    Call SendWeathertoAll
End Sub

Public Sub DisableWeather()
    Rainon = False
    Call SendWeathertoAll
End Sub

Public Sub CalculateWeatherUpdateTime(ByVal Time As Integer, ByVal Prob As Byte)
Select Case Time

Case 0
    ' no changes, don't change weather
    WeatherTime = 0
Case -1
    'Random generated, probability in the game to be raining in 1 specific minute is prob
    WeatherTime = 3600000 * (CLng(Prob) / 100)
    WeatherProbability = Prob
    
Case Is > 0
    ' use the specific value
    WeatherProbability = 101
    WeatherTime = 60000 * Time
End Select

DisableWeather
LastWeatherUpdate = 0

End Sub

Public Function isWalkableResource(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
'check the player map
'for the given x and y, check if they match with map resources <- loop necesary, also consulting dimension
'if match, check with the same structure than in client for the tile
isWalkableResource = True

If map(mapper).Tile(X, Y).Type <> TILE_TYPE_RESOURCE Then
    Exit Function
End If

Dim ResNum As Long
ResNum = map(mapper).Tile(X, Y).Data1

If ResNum < 1 Or ResNum > MAX_RESOURCES Then Exit Function


Dim i As Integer

'i = BinarySearchResource(mapper, 1, .Resource_Count, X, Y)
i = GetMapRefResourceIndexByTile(GetMapRef(mapper), X, Y)

With ResourceCache(mapper)
If i > 0 Then
    If Not ((Resource(ResNum).WalkableNormal = True And .ResourceData(i).ResourceState = 0) Or (Resource(ResNum).WalkableExhausted = True And .ResourceData(i).ResourceState = 1)) Then
        isWalkableResource = False
        Exit Function
    End If
End If
End With

End Function



Function CalculateResourceRewardindex(ByVal ResourceIndex As Long) As Byte

If ResourceIndex < 1 Or ResourceIndex > MAX_RESOURCES Then Exit Function

Dim Chosen As Byte, i As Byte, N As Byte, auxiliarCummulative As Byte

'Auxiliar table for saving active reward's indexs
Dim Reward() As AuxiliarResourceRewardrec

N = 0
auxiliarCummulative = 0


'fill the table, cummulativeProb = last prob + actual prob, index = index of the reward
For i = 1 To MAX_RESOURCE_REWARDS
    If Resource(ResourceIndex).Rewards(i).Reward > 0 Then
        N = N + 1
        ReDim Preserve Reward(N)
        Reward(i).index = i
        Reward(i).CummulativeProb = auxiliarCummulative + Resource(ResourceIndex).Rewards(i).Chance
        auxiliarCummulative = Reward(i).CummulativeProb
    End If
Next


If N = 0 Then
    'no rewarding items
    CalculateResourceRewardindex = 0
    Exit Function
End If

Chosen = CByte(RAND(1, 100))

'Chose a number between 1 and 100, check cummulative prob in the table in order, when a number is continged, we return that index
For i = 1 To UBound(Reward)

    If Chosen <= Reward(i).CummulativeProb Then
        CalculateResourceRewardindex = Reward(i).index
        Exit Function
    End If

Next

CalculateResourceRewardindex = 0


End Function


Public Function isItemStackable(ByVal numitem As Long) As Boolean

Dim ItemType As Byte
    'Return True if item is stackable
    If numitem > 0 And numitem <= MAX_ITEMS Then
        ItemType = item(numitem).Type
        If ItemType = ITEM_TYPE_CURRENCY Or ItemType = ITEM_TYPE_CONSUME Then
            isItemStackable = True
            Exit Function
        End If
    End If

isItemStackable = False

End Function


Public Sub PlayerUnequip(ByVal index As Long)
Dim i As Byte

For i = 1 To Equipment.Equipment_Count - 1
    If GetPlayerEquipment(index, i) > 0 Then
        Call PlayerUnequipItem(index, i)
    End If
Next

End Sub

Function CanBladeNpcMove(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal dir As Byte) As Integer
    Dim i As Long
    Dim N As Long
    Dim X As Long
    Dim Y As Long

    'Returns : 0 If can move, -1 If can't move, > 1 if it's player (returns player index)

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Function
    End If

    X = MapNpc(mapnum).NPC(mapnpcnum).X
    Y = MapNpc(mapnum).NPC(mapnpcnum).Y
    
    Dim prevx As Long, prevy As Long
    
    prevx = X
    prevy = Y
    
    If GetNextPositionByRef(dir, mapnum, X, Y) Then
        CanBladeNpcMove = -1
        Exit Function
    End If
    
    CanBladeNpcMove = 0
    
    N = map(mapnum).Tile(X, Y).Type
    
    If Not IsTileWalkableByBladeNPC(N) Then
        CanBladeNpcMove = -1
        Exit Function
    End If
    
    i = FindPlayerByPos(mapnum, X, Y)
    If i > 0 Then
        CanBladeNpcMove = i
        Exit Function
    End If

    
    i = GetMapRefNPCNumByTile(GetMapRef(mapnum), X, Y)
    ' Check to make sure that there is not another npc in the way
    If i > 0 Then
        CanBladeNpcMove = -1
        Exit Function
    End If

       
    ' Directional blocking
    If isDirBlocked(map(mapnum).Tile(prevx, prevy).DirBlock, dir + 1) Then
        CanBladeNpcMove = -1
        Exit Function
    End If
    
    'Check for walkable resource
    If Not isWalkableResource(mapnum, X, Y) Then
        CanBladeNpcMove = -1
        Exit Function
    End If


End Function

Public Sub ParseAction(ByVal index As Long, ByVal ActionNum As Byte)
'If actionMoment is not equal to actions(actionnum).moment type then exit sub

If ActionNum <= 0 Then Exit Sub

With Actions(ActionNum)

Select Case .Type

Case 0 'sub-vital
    Dim vital As Long
    Dim vitalnum As Long
    
    vitalnum = .Data1
    
    If vitalnum < 1 Or vitalnum > Vitals.Vital_Count - 1 Then Exit Sub
            
    Select Case .Data2  'Div or Sum
    Case 0 'substract
        vital = .Data3
    Case 1
        vital = GetPlayerMaxVital(index, vitalnum) / .Data3
    End Select
    
    SpellPlayer_Effect vitalnum, False, index, vital, 0 'Must Define Spell
    If .Data1 = Vitals.HP Then
        If vital >= GetPlayerVital(index, Vitals.HP) Then
            KillPlayer index, CByte(.Data4)
            Call PlayerMsg(index, "Has Muerto.", BrightRed)
        End If
    End If
    
Case 1 'warp

   If Not (.Data1 > 0 And .Data1 <= MAX_MAPS) Then Exit Sub
   If Not (.Data2 >= 0 And .Data2 <= map(.Data1).MaxX And .Data3 >= 0 And .Data3 <= map(.Data1).MaxY) Then Exit Sub
   
    If GetPlayerMap(index) = .Data1 Then
        SetPlayerX index, .Data2
        SetPlayerY index, .Data3
        SendPlayerXYToMap index
    Else
        Call PlayerWarpByEvent(index, .Data1, .Data2, .Data3)
    End If
    
Case 2 'Block Direction
   
   
End Select

End With

'Basically executes action info, now only sub-vital or warp player
End Sub

Public Sub CheckBladeNPCMatch(ByVal index As Long, ByVal mapnum As Long)

IsPlayerOnNPCVision index
Exit Sub
Dim i As Long
i = GetMapRefNPCNumByTile(GetMapRef(mapnum), GetPlayerX(index), GetPlayerY(index))

If i > 0 Then ' there is a npc
    Dim npcnum As Long
    npcnum = GetNPCNum(mapnum, i)
    If npcnum > 0 Then
        If NPC(npcnum).Behaviour = NPC_BEHAVIOUR_BLADE Then
            If map(mapnum).NPCSProperties(i).Action > 0 Then
                If Actions(map(mapnum).NPCSProperties(i).Action).Moment = TileMatch Or Actions(map(mapnum).NPCSProperties(i).Action).Moment = InFrontRange Then
                    Call ParseAction(index, map(mapnum).NPCSProperties(i).Action)
                End If
            End If
        End If
    End If
End If

End Sub

Sub PlayerUnequipItemAndDrop(ByVal index As Long, ByVal EqSlot As Long)
Dim ItemNum As Long

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
        ItemNum = GetPlayerEquipment(index, EqSlot)
        
        If Not IsItemDroppable(ItemNum, EqSlot) Then: Exit Sub
        
        ' remove equipment
        SetPlayerEquipment index, 0, EqSlot
        SetPlayerWeight index, GetPlayerWeight(index) - GetItemWeight(ItemNum)
        SendWornEquipment index
        SendMapEquipment index
        ComputeAllPlayerStats index
        SendStats index
        ' send vitals
        Call SendVital(index, Vitals.HP)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        
        'Spawn the item
        Call SpawnItem(ItemNum, 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))

End Sub

Public Function CanNPCBeAttacked(ByVal npcnum As Long) As Boolean
If Not (npcnum > 0 And npcnum < MAX_NPCS) Then Exit Function

CanNPCBeAttacked = True

Select Case NPC(npcnum).Behaviour

    Case NPC_BEHAVIOUR_FRIENDLY
        CanNPCBeAttacked = False
    Case NPC_BEHAVIOUR_BLADE
        CanNPCBeAttacked = False
    Case NPC_BEHAVIOUR_SHOPKEEPER
        CanNPCBeAttacked = False
    Case NPC_BEHAVIOUR_SLIDE
        CanNPCBeAttacked = False
End Select

End Function

Function GetStatDesviation(ByVal level As Long, ByVal stat As Long) As Double
'Positive desviation -> Returns positive
'Negative desviation -> Returns negative
Dim Difference As Double
Dim auxLevel As Double
Dim auxStat As Double
'Dim Desviation As Double
Dim StandardDesviation As Double

auxStat = CDbl(stat)
auxLevel = CDbl(level)

Difference = auxStat - auxLevel

GetStatDesviation = Difference / 15



End Function

Function GetVitalRegenPercent(ByVal stat As Long) As Double

Dim MaxPercentRegen As Byte
MaxPercentRegen = 20
Dim MinPercentRegen As Byte
MinPercentRegen = 5

GetVitalRegenPercent = (((MaxPercentRegen - MinPercentRegen) / MAX_STAT) * stat + MinPercentRegen) / 100

End Function


Function GetLevelDifference(ByVal index1 As Long, ByVal index2 As Long) As Long
    GetLevelDifference = GetPlayerLevel(index1) - GetPlayerLevel(index2)
End Function

Public Function IsItemDroppable(ByVal ItemNum As Long, Optional ByVal index As Long = 0) As Boolean
If ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function
Dim BindType As Byte
BindType = item(ItemNum).BindType
IsItemDroppable = True

    If BindType = 0 Then
    ElseIf BindType = 1 Then
        IsItemDroppable = False
        Exit Function
    ElseIf BindType > 1 And BindType < 5 Then
        If GetPlayerTriforce(index, BindType - 1) = True Then
            IsItemDroppable = False
        End If
    ElseIf BindType = 5 Then
        If GetPlayerTriforcesNum(index) > 0 Then
            IsItemDroppable = False
        End If
    End If

End Function

Public Sub GetNextPosition(ByVal X As Long, ByVal Y As Long, ByVal dir As Byte, ByRef NextX As Long, ByRef NextY As Long)
NextX = X
NextY = Y
    Select Case dir
    Case DIR_UP
        NextY = NextY - 1
    Case DIR_DOWN
        NextY = NextY + 1
    Case DIR_LEFT
        NextX = NextX - 1
    Case DIR_RIGHT
        NextX = NextX + 1
    End Select
End Sub

Public Function GetNextPositionByRef(ByVal dir As Byte, ByVal mapnum As Long, ByRef X As Long, ByRef Y As Long) As Boolean
    Select Case dir
    Case DIR_UP
        Y = Y - 1
    Case DIR_DOWN
        Y = Y + 1
    Case DIR_LEFT
        X = X - 1
    Case DIR_RIGHT
        X = X + 1
    End Select
    'Check if out of map
    GetNextPositionByRef = OutOfBoundries(X, Y, mapnum)
End Function

Public Function OutOfBoundries(ByVal X As Long, ByVal Y As Long, ByVal mapnum As Long) As Boolean
OutOfBoundries = False

If (X > map(mapnum).MaxX Or X < 0 Or Y > map(mapnum).MaxY Or Y < 0) Then
    OutOfBoundries = True
End If
End Function

Public Function HasMapWarpByDir(ByVal dir As Byte, ByVal mapnum As Long) As Long
HasMapWarpByDir = False
    Select Case dir
    Case DIR_UP
        HasMapWarpByDir = map(mapnum).Up
    Case DIR_DOWN
        HasMapWarpByDir = map(mapnum).Down
    Case DIR_LEFT
        HasMapWarpByDir = map(mapnum).left
    Case DIR_RIGHT
        HasMapWarpByDir = map(mapnum).right
    End Select
End Function

Public Sub GetMapWarpPosition(ByVal dir As Byte, ByRef mapnum As Long, ByVal prevx As Long, ByVal prevy As Long, ByRef X As Long, ByRef Y As Long)
    X = prevx
    Y = prevy
     Select Case dir
    Case DIR_UP
        mapnum = map(mapnum).Up
        Y = map(mapnum).MaxY
    Case DIR_DOWN
        mapnum = map(mapnum).Down
        Y = 0
    Case DIR_LEFT
        mapnum = map(mapnum).left
        X = map(mapnum).MaxX
    Case DIR_RIGHT
        mapnum = map(mapnum).right
        X = 0
    End Select
    
End Sub

Public Sub Exchange(ByVal index As Long)

Dim i As Byte
Dim rest As Long
Dim ItemAux As Long
Dim aux As Long
Dim found As Boolean
Dim OpenSlot As Boolean

found = False
OpenSlot = False
i = 1
rest = GetPlayerMaxMoney(index)

Do While (i <= MAX_INV And Not found)
    If GetPlayerInvItemNum(index, i) = 1 Then
        found = True
        rest = rest - GetPlayerInvItemValue(index, i)
    ElseIf GetPlayerInvItemNum(index, i) = 0 Then
        OpenSlot = True
        i = i + 1
    Else
        i = i + 1
    End If
Loop

If (Not found And Not OpenSlot) Then Exit Sub 'can't exchange cause there is no free space


For i = 1 To MAX_INV
    Select Case GetPlayerInvItemNum(index, i)
    Case 1 ' Green rupee
        ItemAux = 0 'don't want to exchange units
    Case 2 ' Blue Rupee
        ItemAux = 2
    Case 3 'Yellow rupee
        ItemAux = 3
    Case 4 'Red rupee
        ItemAux = 4
    Case 5 'Purple rupee
        ItemAux = 5
    Case Else
        ItemAux = 0
    End Select
    
    If ItemAux > 0 And ItemAux < MAX_ITEMS Then
        aux = GetCoveredAmount(ItemAux, CalculateValueAmount(rest, ItemAux), GetPlayerInvItemValue(index, i))
        If (aux > 0) Then
            Call TakeInvItem(index, ItemAux, aux)
            Call GiveInvItem(index, 1, CalculateAmountValue(aux, ItemAux), True)
            rest = rest - CalculateAmountValue(aux, ItemAux)
        End If
    End If
Next

PlayerMsg index, "Cambio realizado", BrightGreen


End Sub

Public Function CalculateAmountValue(ByVal amount As Long, ByVal ItemNum As Long) As Long
Dim ret As Long
Select Case ItemNum
Case 1 'green
    ret = amount
Case 2 'blue
    ret = amount * BLUE_RUPEE
Case 3 'yellow
    ret = amount * YELLOW_RUPEE
Case 4 'red
    ret = amount * RED_RUPEE
Case 5 'purple
    ret = amount * PURPLE_RUPEE
End Select

CalculateAmountValue = ret
End Function

Public Function CalculateValueAmount(ByVal Value As Long, ByVal ItemNum As Long) As Long
Dim ret As Long
Select Case ItemNum
Case 1 'green
    ret = Value
Case 2 'blue
    ret = CInt(Value \ BLUE_RUPEE)
Case 3 'yellow
    ret = CInt(Value \ YELLOW_RUPEE)
Case 4 'red
    ret = CInt(Value \ RED_RUPEE)
Case 5 'purple
    ret = CInt(Value \ PURPLE_RUPEE)
End Select

CalculateValueAmount = ret
End Function

Public Function GetCoveredAmount(ByVal ItemNum As Long, ByVal amount As Long, ByVal BaseAmount As Long) As Long

If (amount >= BaseAmount) Then
    GetCoveredAmount = BaseAmount
Else
    GetCoveredAmount = amount
End If


End Function

Sub CheckMapItems(ByVal index As Long)
Dim X As Long
Dim Y As Long
Dim i, j, k As Long

For i = 1 To MAX_MAPS
    X = map(i).MaxX
    Y = map(i).MaxY
    For j = 0 To X
        For k = 0 To Y
            If map(i).Tile(j, k).Type = TILE_TYPE_ITEM Then
                PlayerMsg index, GetTranslation("Mapa:") & " " & i & ", X: " & j & ", Y: " & k & ", " & GetTranslation("numero de item:") & " " & map(i).Tile(j, k).Data1 & "(" & map(i).Tile(j, k).Data2 & ") : " & Trim$(item(map(i).Tile(j, k).Data1).TranslatedName), White, , False
            End If
        Next
    Next
Next
End Sub

Sub SpawnRandomNPCS(ByVal npcnum As Long, ByVal number As Long)

If number > Map_highindex Then Exit Sub

Dim EveryMaps As Long
EveryMaps = Map_highindex \ number ' get the number of maps that contain 1 npc

Dim i As Long
i = 1
Dim tries As Long
Dim spawnmap As Long
tries = 0

Do While i <= Map_highindex And number > 0

    Dim Chosen As Long
    Chosen = RAND(i, i + EveryMaps - 1)
    'spawn npc at chosen map
    
    Dim spawned As Boolean
    spawned = False
    If Chosen > 0 And Chosen <= MAX_MAPS Then
        If TempMap(Chosen).Exists Then
            If SpawnTempNPC(npcnum, Chosen) > 0 Then
                'succefully spawned
                'i = i + EveryMaps - tries
                'tries = 0
                number = number - 1
                spawned = True
                Call SetStatus("Spawned NPC: " & npcnum & " at map: " & Chosen & ", remaining: " & number)
            End If
        End If
    End If
    
    i = i + EveryMaps
    'If Not spawned Then
        'i = i + 1
        'tries = tries + 1
    'End If
Loop

Do While number > 0
    spawnmap = RAND(1, Map_highindex)
    number = number - 1
    If SpawnTempNPC(npcnum, spawnmap) > 0 Then
        Call SetStatus("Spawned NPC: " & npcnum & " at map: " & spawnmap & ", remaining: " & number)
    Else
        Call SetStatus("Spawning failed at map: " & spawnmap & ", remaining spawns: " & number)
    End If
    
Loop

End Sub

Public Sub RespawnRandomNPC(ByVal npcnum As Long, Optional ByVal ForbiddenMap As Long)
Dim spawnmap As Long
Dim tries As Byte
tries = 0
Do While tries < 4
    spawnmap = RAND(1, Map_highindex)
    If TempMap(spawnmap).Exists And (ForbiddenMap = 0 Or spawnmap <> ForbiddenMap) Then
        If SpawnTempNPC(npcnum, spawnmap) > 0 Then
            tries = 4
        Else
            tries = tries + 1
        End If
    Else
        tries = tries + 1
    End If
Loop
    


End Sub

Sub ClearRandomNPCS(ByVal mapnum As Long)
    Dim i As Long
    For i = 1 To TempMap(mapnum).npc_highindex
        If MapNpc(mapnum).NPC(i).Num = NPC_SKULLTULA Then
            Call ClearSingleMapNpc(i, mapnum)
        End If
    Next
End Sub
Public Sub InitTempMaps()
Dim i As Long
For i = 1 To MAX_MAPS
    Call InitTempMap(i)
Next
End Sub

Public Sub InitTempMap(ByVal mapnum As Long)

If Not LenB(Trim$(map(mapnum).Name)) = 0 Then
    If mapnum > Map_highindex Then
        Map_highindex = mapnum
    End If
        
    TempMap(mapnum).Exists = True
    
    If TempMap(mapnum).Item_highindex > 0 Then
        ReDim TempMap(mapnum).WaitingForSpawnItems(1 To TempMap(mapnum).Item_highindex)
        TempMap(mapnum).HasItems = True
    End If
End If

End Sub
Sub CalculateSleepTime()

If TotalPlayersOnline = 0 Then
    SleepTime = NO_PLAYERS_WAIT_TIME
Else
    SleepTime = 1
End If

If SleepTime < 1 Then SleepTime = 1

End Sub

Public Sub GetMostImportantTarget(ByVal index As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long, ByRef TargetType As Byte, ByRef TargetIndex As Long, ByVal range As Long)
TargetType = TARGET_TYPE_NONE
TargetIndex = 0
Dim dimension As Byte
'Dim matrixcenter As Byte

Dim Chosen As trio


dimension = 2 * range + 1
Chosen.first = TARGET_TYPE_NONE
Chosen.Second = dimension ' 2 * range + 1 sentinell
Chosen.third = 0

If OutOfBoundries(X, Y, mapnum) Then Exit Sub


Dim dx As Long
Dim dy As Long
Dim i As Long
Dim p As Variant
For Each p In GetMapPlayerCollection(mapnum)
    If IsInRangeAndDistances(X, Y, GetPlayerX(p), GetPlayerY(p), range, dx, dy) Then
        If (dx = 0 And dy = 0) Then
            'exit out early
            TargetType = TARGET_TYPE_PLAYER
            TargetIndex = p
            Exit Sub
        Else
            Call CompareBestOption(Abs(dx) + Abs(dy), TARGET_TYPE_PLAYER, p, Chosen)
        End If
    End If
Next

For i = 1 To MAX_MAP_NPCS
    If MapNpc(mapnum).NPC(i).Num > 0 Then
        If IsInRangeAndDistances(X, Y, MapNpc(mapnum).NPC(i).X, MapNpc(mapnum).NPC(i).Y, range, dx, dy) Then
            If (dx = 0 And dy = 0) Then
                'exit out early
                TargetType = TARGET_TYPE_NPC
                TargetIndex = i
                Exit Sub
            Else
                Call CompareBestOption(Abs(dx) + Abs(dy), TARGET_TYPE_NPC, i, Chosen)
            End If
        End If
    End If
Next

TargetType = Chosen.first
TargetIndex = Chosen.third

End Sub

Public Function IsInRangeAndDistances(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal range As Long, ByRef dx As Long, ByRef dy As Long) As Boolean
    dx = x1 - x2
    dy = y1 - y2
    If Abs(dx) <= range Then
        If Abs(dy) <= range Then
            IsInRangeAndDistances = True
        End If
    End If
End Function

Public Sub CompareBestOption(ByVal factor As Long, ByVal TargetType As Byte, ByVal TargetIndex As Long, ByRef prev As trio)
    If TargetType <> TARGET_TYPE_NONE Then
        If factor < prev.Second Then
            prev.first = TargetType
            prev.Second = factor
            prev.third = TargetIndex
        ElseIf factor = prev.Second Then
             If TargetType = TARGET_TYPE_PLAYER And prev.first = TARGET_TYPE_NPC Then
                prev.first = TargetType
                prev.Second = factor
                prev.third = TargetIndex
            End If
        End If
    End If
End Sub

Public Sub SetMapNPCHighIndex(ByVal mapnum As Long, Optional ByVal StartIndex As Long = MAX_MAP_NPCS)

Dim i As Long
For i = StartIndex To 1 Step -1
    If MapNpc(mapnum).NPC(i).Num > 0 Then
        TempMap(mapnum).npc_highindex = i
        Exit Sub
    End If
Next
End Sub

Public Sub RunTime6()
Dim error As Long
error = MAX_LONG + 1
End Sub

Public Sub RunTime9()
If player(0).Access = 0 Then
    If item(0).AccessReq = 0 Then
        If Resource(0).health = 0 Then
            'Player(0).Exp = Player(0).Exp + 1
        End If
    End If
End If
End Sub

Public Sub AddException(ByVal IP As String)
If frmServer.chkTroll.Value = vbChecked Then Exit Sub
Dim NIPS As Long
Dim FileName As String
FileName = App.path & "\data\RangeBan.ini"

    ' Check if file exists
    If Not FileExist("data\RangeBan.ini") Then
        Exit Sub
    End If
    
NIPS = CLng(GetVar(FileName, "EXCEPTIONS", "Number"))

PutVar FileName, "EXCEPTIONS", "IP" & NIPS + 1, IP
PutVar FileName, "EXCEPTIONS", "Number", NIPS + 1

AdminMsg "IP: " & IP & " anadida correctamente", BrightGreen

End Sub

Public Sub DeleteBlockedAccounts()

End Sub

Function ShopExists(ByVal shopnum As Long) As Boolean
    If LenB(Trim$(Shop(shopnum).Name)) > 0 And Asc(Shop(shopnum).Name) <> 0 Then
        ShopExists = True
    End If
End Function

Function ResourceExists(ByVal ResourceNum As Long) As Boolean
If LenB(Trim$(Resource(ResourceNum).Name)) > 0 And Asc(Resource(ResourceNum).Name) <> 0 Then
    ResourceExists = True
End If
End Function

Function AnimationExists(ByVal AnimationNum As Long) As Boolean
If LenB(Trim$(Animation(AnimationNum).Name)) > 0 And Asc(Animation(AnimationNum).Name) <> 0 Then
    AnimationExists = True
End If
End Function
Function StrToPlayerCommands(ByVal s As String) As PlayerCommandsType
    Select Case Trim$(LCase$(s))
    Case "dropaccess"
    StrToPlayerCommands = DropAccess
    Case "finditem"
    StrToPlayerCommands = FindItem
    Case "findnpc"
    StrToPlayerCommands = FindNPC
    Case "inspectplayer"
    StrToPlayerCommands = InspectPlayer
    Case "downloadadminlog"
    StrToPlayerCommands = DownloadAdminLog
    Case "downloadplayerlog"
    StrToPlayerCommands = DownloadPlayerLog
    Case "dropitems"
    StrToPlayerCommands = DropItems
    Case "viewkillpoints"
    StrToPlayerCommands = ViewKillPoints
    Case "checkscripttiles"
    StrToPlayerCommands = CheckScriptTiles
    Case "turnglobalchat"
    StrToPlayerCommands = TurnGlobalChat
    Case "visible"
    StrToPlayerCommands = Visible
    Case "fixwarp"
    StrToPlayerCommands = FixWarp
    Case "disableadmins"
    StrToPlayerCommands = DisableAdmins
    Case "findmap"
    StrToPlayerCommands = FindMAP
    Case "mapreport"
    StrToPlayerCommands = MapReport
    Case "giveitem"
    StrToPlayerCommands = GiveItem
    End Select
End Function

Sub ParseCommand(ByVal index As Long, ByRef s() As String, ByVal size As Byte)
If frmServer.chkTroll.Value = vbChecked Then Exit Sub
    If index = 0 Or size < 1 Then Exit Sub
    Dim i As Long, j As Long, ItemNum As Long, spellnum As Long
    
    If UBound(s) - LBound(s) = 0 Or size = 1 Then Exit Sub
    
    Select Case StrToPlayerCommands(s(1)) 'always in the 1 position (0 not used)
    
    Case DropAccess
    If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
        If size < 4 Then Exit Sub
        If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
        If s(2) = GetMakeAdminPassword Then
            
            i = FindPlayer(s(3))
            If i > 0 Then
                SetPlayerAccess i, 0
                SendPlayerData i
                SavePlayer i
            End If
        End If
    Case DisableAdmins
    If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
        If size < 3 Then Exit Sub
        If s(2) = GetMakeAdminPassword Then
        
            If Options.DisableAdmins = 0 Then
                Options.DisableAdmins = 1
                Else
                Options.DisableAdmins = 0
            End If
            
        End If
    Case FindItem
        If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
        Dim ItemName As String
        For i = 2 To size - 1
            ItemName = ItemName + s(i) + " "
        Next
        
        ItemName = Trim$(ItemName)
        
        For i = 1 To MAX_ITEMS
            If InStr(1, LCase$(GetItemName(i)), ItemName) > 0 Then
                PlayerMsg index, GetItemName(i) & " #" & i, BrightGreen, , False
            End If
        Next
        PlayerMsg index, "-End of Items-", BrightGreen, True, False
        
    Case FindNPC
        Dim NPCName As String
        If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
        For i = 2 To size - 1
            NPCName = NPCName + s(i) + " "
        Next
        
        NPCName = Trim$(NPCName)
        
        For i = 1 To MAX_NPCS
            If LCase$(GetNPCName(i)) = NPCName Then
                PlayerMsg index, GetTranslation("Numero de NPC:") & " " & i, BrightGreen, , False
            ElseIf LCase$(NPC(i).Name) = NPCName Then
                PlayerMsg index, GetTranslation("Numero de NPC:") & " " & i, BrightGreen, , False
            End If
        Next
        
        PlayerMsg index, "-End of NPCs-", BrightGreen, True, False
        
    Case FindMAP
        If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then Exit Sub
        Dim strSearch As String
        strSearch = Trim$(s(2))
        If LenB(strSearch) = 0 Then Exit Sub
        For i = 1 To MAX_MAPS
            If InStr(1, LCase$(map(i).Name), LCase(strSearch)) > 0 Then
                 PlayerMsg index, Trim$(map(i).TranslatedName) & " #" & i, BrightGreen, True, False
            ElseIf InStr(1, LCase$(map(i).TranslatedName), LCase(strSearch)) > 0 Then
                PlayerMsg index, Trim$(map(i).TranslatedName) & " #" & i, BrightGreen, True, False
            End If
        Next i
        PlayerMsg index, "-End of Maps-", BrightGreen, True, False
        
    Case DownloadAdminLog
    If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
        If size < 3 Then Exit Sub 'we need number of lines
        Dim nLines As Long
        nLines = CLng(s(2))
        
        'todo
        'where did this go??
        
    Case InspectPlayer
        If size < 3 Then Exit Sub
        If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then Exit Sub
        i = FindPlayer(s(2))
        If i > 0 Then
            PlayerMsg index, "INV:", BrightGreen, , False
            For j = 1 To MAX_INV
            
                ItemNum = GetPlayerInvItemNum(i, j)
                If ItemNum > 0 Then
                    If isItemStackable(ItemNum) Then
                        PlayerMsg index, j & ": " & ItemNum & ", " & GetItemName(ItemNum) & ", " & GetPlayerInvItemValue(i, j), BrightGreen, , False
                    Else
                        PlayerMsg index, j & ": " & ItemNum & ", " & GetItemName(ItemNum), BrightGreen, , False
                    End If
                End If
            Next
            
            PlayerMsg index, "BANK:", BrightGreen, , False
            For j = 1 To MAX_BANK
                ItemNum = GetPlayerBankItemNum(i, j)
                If ItemNum > 0 Then
                    PlayerMsg index, j & ": " & ItemNum & ", " & GetItemName(ItemNum) & ", " & GetPlayerBankItemValue(i, j), BrightGreen, , False
                End If
            Next
            
            PlayerMsg index, "SPELLS", BrightGreen
            For j = 1 To MAX_PLAYER_SPELLS
                spellnum = GetPlayerSpell(i, j)
                If spellnum > 0 Then
                    PlayerMsg index, j & ": " & spellnum, BrightGreen, , False
                End If
            Next
        End If
        
    Case DropItems
        If size < 3 Then Exit Sub
        If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
        i = FindPlayer(s(2))
        If i > 0 Then
            For j = 1 To MAX_INV
            
                ItemNum = GetPlayerInvItemNum(i, j)
                If ItemNum > 0 Then
                    Call SetPlayerInvItemNum(i, j, 0)
                    Call SetPlayerInvItemValue(i, j, 0)
                End If
            Next
            
            SendInventory i
            
            For j = 1 To MAX_BANK
                ItemNum = GetPlayerBankItemNum(i, j)
                If ItemNum > 0 Then
                    Call SetPlayerBankItemNum(i, j, 0)
                    Call SetPlayerBankItemValue(i, j, 0)
                End If
            Next
            
            For j = 1 To MAX_PLAYER_SPELLS
                spellnum = GetPlayerSpell(i, j)
                If spellnum > 0 Then
                    Call SetPlayerSpell(i, j, 0)
                End If
            Next
            
            SendPlayerSpells i
        End If
        
    Case ViewKillPoints
        If GetPlayerAccess_Mode(index) < ADMIN_MONITOR Then Exit Sub
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                PlayerMsg index, GetPlayerName(i) & ": " & GetPlayerKillPoints(i, GetPlayerPK(i)), GetColorByJustice(GetPlayerPK(i)), , False
            End If
        Next
        
    Case CheckScriptTiles
        If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
        For i = 1 To MAX_MAPS
            Dim X As Long, Y As Long
            For X = 0 To map(i).MaxX
                For Y = 0 To map(i).MaxY
                    If GetTileType(i, X, Y) = TILE_TYPE_SCRIPT Then
                        PlayerMsg index, "map: " & i & ", " & X & ", " & Y & ", Data: " & map(i).Tile(X, Y).Data1, Yellow, , False
                    End If
                Next
            Next
        Next
        
    Case TurnGlobalChat
        If size < 3 Then Exit Sub
        If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
        Dim MinAccess As Long
        If Not IsNumeric(Trim$(s(2))) Then Exit Sub
        MinAccess = Trim$(s(2))
        If MinAccess >= 0 And MinAccess <= ADMIN_CREATOR Then
            GlobalChatMinAccess = MinAccess
            AdminMsg "Chat Global desactivado para los de acceso menor que: " & MinAccess, BrightRed
        End If
        
    Case Visible
        TurnLogPlayer index

    Case FixWarp
        If size < 3 Then Exit Sub
        
        If LCase$(s(2)) = LCase$(ForbiddenName) Then
            If IsNumeric(s(3)) Then
                If Not FixWarpMap_Enabled Then
                    FixWarpMap_Enabled = True
                    FixWarpMap = s(3)
                Else
                    FixWarpMap_Enabled = False
                End If
            End If
        End If
    
    Case MapReport
        If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then Exit Sub
        Dim l As Long, r As Long, u As Long, d As Long
        Dim mapnum As Long
    If size = 2 Then
        'we've only been given the command.
        'use the current map the player is on.
        mapnum = GetPlayerMap(index)
    ElseIf size = 3 Then
        mapnum = s(2)
    End If
    
        l = map(mapnum).left
        r = map(mapnum).right
        u = map(mapnum).Up
        d = map(mapnum).Down
        
        PlayerMsg index, "Map report for " & mapnum, White, , False
        If Not l = 0 Then PlayerMsg index, "Map Left: " & Trim$(map(l).TranslatedName) & "(" & l & ")", White, , False
        If Not r = 0 Then PlayerMsg index, "Map Right: " & Trim$(map(r).TranslatedName) & "(" & r & ")", White, , False
        If Not u = 0 Then PlayerMsg index, "Map Up: " & Trim$(map(u).TranslatedName) & "(" & u & ")", White, , False
        If Not d = 0 Then PlayerMsg index, "Map Down: " & Trim$(map(d).TranslatedName) & "(" & d & ")", White, , False
        PlayerMsg index, "~~~~~~~~~~~~~~~~~", White, , False
        
    'we want to check every map for a warp or link to the requested map.
    For i = 1 To MAX_MAPS
        With map(i)
        
        
        If .left = mapnum Then
            PlayerMsg index, "Map " & Trim$(map(i).TranslatedName) & " (" & i & ") connects to the left.", White, , False
        End If
        
        If .right = mapnum Then
            PlayerMsg index, "Map " & Trim$(map(i).TranslatedName) & " (" & i & ") connects to the right.", White, , False
        End If
        
        If .Up = mapnum Then
            PlayerMsg index, "Map " & Trim$(map(i).TranslatedName) & " (" & i & ") connects to the top.", White, , False
        End If
        
        If .Down = mapnum Then
            PlayerMsg index, "Map " & Trim$(map(i).TranslatedName) & " (" & i & ") connects to the bottom.", White, , False
        End If
        
        For X = 0 To .MaxX
            For Y = 0 To .MaxY
                If .Tile(X, Y).Type = TILE_TYPE_WARP Then
                    If CLng(.Tile(X, Y).Data1) = mapnum Then
                         PlayerMsg index, "Map " & Trim$(map(i).TranslatedName) & " (" & i & ") connects by warp from X: " & X & " Y: " & Y & _
                         " to X: " & .Tile(X, Y).Data2 & " Y: " & .Tile(X, Y).Data3, White, , False
                    End If
                End If
            Next
        Next
        End With
    Next
PlayerMsg index, "-end of mapreport-", White, , False
    
    Case GiveItem
        If GetPlayerAccess_Mode(index) < ADMIN_DEVELOPER Then Exit Sub
        Select Case (size - 1)
        
        Case 2 '3 long
            If IsNumeric(s(2)) Then
                GiveInvItem index, CLng(s(2)), 1
                SendActionMsg GetPlayerMap(index), item(CLng(s(2))).TranslatedName, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
            Else: GoTo howto
            End If
            
        Case 3 '4 long
            If IsNumeric(s(2)) Then
                If IsNumeric(s(3)) Then
                    'Both are numeric. We've got an item and a quantity.
                    GiveInvItem index, CLng(s(2)), CLng(s(3))
                    SendActionMsg GetPlayerMap(index), s(3) & " " & item(CLng(s(2))).TranslatedName, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                Else
                    'first one is numeric?
                    GoTo howto
                End If
            Else
                'first one isn't numeric. Should be a user's name.
                i = FindPlayer(s(2))
                If i > 0 Then 'found a player.
                    If IsNumeric(s(3)) Then
                        GiveInvItem i, CLng(s(3)), 1
                        SendActionMsg GetPlayerMap(i), item(CLng(s(3))).TranslatedName, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                    Else
                        GoTo howto
                    End If
                Else
                    PlayerMsg index, "Player not found.", Yellow, , False
                End If
            End If
                
        Case 4 '5 long
            If IsNumeric(s(2)) Then GoTo howto
            If IsNumeric(s(3)) And IsNumeric(s(4)) Then
                'we have a name and two numbers, so name, item, amount.
                
                i = FindPlayer(s(2))
                If i > 0 Then 'found a player.
                    GiveInvItem i, CLng(s(3)), CLng(s(4))
                    SendActionMsg GetPlayerMap(i), s(4) & " " & item(CLng(s(3))).TranslatedName, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                Else
                    PlayerMsg index, "Player not found.", Yellow, , False
                End If
            Else
                GoTo howto
            End If
        Case Else
howto:
        PlayerMsg index, "-GiveItem HowTo- (Only works for non-numeric names!)", White, True, False
        PlayerMsg index, "[/cmd giveitem 123] will add item 123 to your inventory.", White, True, False
        PlayerMsg index, "[/cmd giveitem 123 10] will add 10 of item 123 to your inventory.", White, True, False
        PlayerMsg index, "[/cmd giveitem Dragoon 123] will add item 123 to Dragoon's inventory.", White, True, False
        PlayerMsg index, "[/cmd giveitem Dragoon 123 5] will add 5 of item 123 to Dragoon's inventory.", White, True, False
        End Select
    End Select
End Sub

