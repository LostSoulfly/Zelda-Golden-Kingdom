Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal index As Long)
    If Not IsPlaying(index) Then
        Call JoinGame(index)
        Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal index As Long)
    Dim i As Long, j As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
    frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
    frmServer.lvwInfo.ListItems(index).SubItems(3) = GetPlayerName(index)
    
    ' send the login ok
    SendLoginOk index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call SendDoors(index)
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendAnimations(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendResources(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendMapEquipment(index)
    Call SendPlayerSpells(index)
    Call SendHotbar(index)
    Call SendQuests(index)
    Call SendWeather
    Call SendMovements(index)
    Call SendActions(index)
    Call SendPets(index)
    'Call SendMap(index, GetPlayerMap(index))
    
    
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    SendEXP index
    Call SendStats(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(index) & "has connected", BrightGreen)
    Else
        Call GlobalMsg(GetPlayerName(index) & " has connected", BrightGreen)
    End If
    
    ' Send welcome messages
    Call SendWelcome(index)
    
    'Do all the guild start up checks
    Call GuildLoginCheck(index)

    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, i
    Next
    
    'miscellanious
    Call InitPlayerPets(index)
    Call SendPetData(index, TempPlayer(index).ActualPet)
    
    ' Send the flag so they know they can start doing stuff
    SendInGame index
End Sub

Sub LeftGame(ByVal index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    
    If TempPlayer(index).InGame Then
        TempPlayer(index).InGame = False

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' cancel any trade they're in
        If TempPlayer(index).InTrade > 0 Then
            tradeTarget = TempPlayer(index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & "has refused the trade.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' leave party.
        Party_PlayerLeave index

If Player(index).GuildFileId > 0 Then
'Set player online flag off
GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(Player(index).GuildMemberId).Online = False
Call CheckUnloadGuild(TempPlayer(index).tmpGuildSlot)
End If
        ' save and clear data.
        Call SavePlayer(index)
        Call SaveBank(index)
        Call ClearBank(index)

        ' Send a global message that he/she left
        If GetPlayerAccess(index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(index) & " has disconnected", BrightRed)
        Else
            Call GlobalMsg(GetPlayerName(index) & " has disconnected", BrightRed)
        End If

        Call TextAdd(GetPlayerName(index) & " has disconnected" & Options.Game_Name & ".")
        Call SendLeftGame(index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(index)
End Sub

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(index, Armor)
    Helm = GetPlayerEquipment(index, Helmet)
    GetPlayerProtection = (GetPlayerStat(index, Stats.endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If

End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(index, Weapon) > 0 Then
        n = (Rnd) * 1.333

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Strength) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(index, Shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.endurance) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Sub PlayerWarp(ByVal index As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(mapnum).MaxX Then x = Map(mapnum).MaxX
    If y > Map(mapnum).MaxY Then y = Map(mapnum).MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    ' if same map then just send their co-ordinates
    If mapnum = GetPlayerMap(index) Then
        SendPlayerXYToMap index
        'Alatar v1.2
        'Call CheckTasks(index, QUEST_TYPE_GOREACH, mapnum)
        '/Alatar v1.2
    End If
    
    ' clear target
    TempPlayer(index).target = 0
    TempPlayer(index).targetType = TARGET_TYPE_NONE
    SendTarget index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)

    If OldMap <> mapnum Then
        Call SendLeaveMap(index, OldMap)
    End If

    Call SetPlayerMap(index, mapnum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
    
    'If 'refreshing' map
    If (OldMap <> mapnum) And TempPlayer(index).TempPetSlot > 0 Then
        'switch maps
        PetDisband index, OldMap, False
        'SpawnPet index, MapNum
        'PetFollowOwner index

        If PetMapCache(OldMap).UpperBound > 0 Then
            For i = 1 To PetMapCache(OldMap).UpperBound
                If PetMapCache(OldMap).Pet(i) = TempPlayer(index).TempPetSlot Then
                    PetMapCache(OldMap).Pet(i) = 0
                End If
            Next
        Else
            PetMapCache(OldMap).Pet(1) = 0
        End If
    End If

    'View Current Pets on Map
    If PetMapCache(Player(index).Map).UpperBound > 0 Then
        For i = 1 To PetMapCache(Player(index).Map).UpperBound
            Call NPCCache_Create(index, Player(index).Map, PetMapCache(Player(index).Map).Pet(i))
        Next
    End If
    
    ' send player's equipment to new map
    SendMapEquipment index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(mapnum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = mapnum Then
                    SendMapEquipmentTo i, index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO

        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS

            If mapnpc(OldMap).NPC(i).Num > 0 Then
                mapnpc(OldMap).NPC(i).vital(Vitals.HP) = GetNpcMaxVital(mapnpc(OldMap).NPC(i).Num, Vitals.HP)
            End If

        Next

    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES
    TempPlayer(index).GettingMap = YES
    
    'ALATAR
    Call CheckTasks(index, QUEST_TYPE_GOREACH, mapnum)
    '/ALATAR
    
    Call SendMapNpcsTo(index, mapnum)
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong mapnum
    Buffer.WriteLong Map(mapnum).Revision
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, mapnum As Long
    Dim x As Long, y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, amount As Long, Scriptnum As Long, i As Integer
    Dim doornum As Long
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Moved = NO
    mapnum = GetPlayerMap(index)
    
Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                        'If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_RESOURCE Then
                        If isWalkableResource(GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index) - 1) Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_DOOR Then
                                    If Player(index).PlayerDoors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).state = 0 Then
                                            
                                            
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).DoorType = 0 Then
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).UnlockType = 0 Then
                                                    PlayerMsg index, "You need the right key to open this door. (" & Trim$(Item(Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).key).Name) & ")", BrightRed
                                                ElseIf Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).UnlockType = 1 Then
                                                    PlayerMsg index, "You have to flip a switch to open this door.", BrightRed
                                                Else
                                                    mapnum = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).WarpMap
                                                    x = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).WarpX
                                                    y = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).WarpY
                                                    Call PlayerWarp(index, mapnum, x, y)
                                                    Moved = YES
                                                End If
                                                End If
                                            
                                            PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                            Exit Sub
                                    Else
                                            If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).DoorType = 1 Then
                                                PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                                Exit Sub
                                            End If
                                    End If
                                End If
                                
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = YES) Then
                                    Call SetPlayerY(index, GetPlayerY(index) - 1)
                                    SendPlayerMove index, Movement, sendToSelf
                                    Moved = YES
                                End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).Up).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < Map(mapnum).MaxY Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                        'If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_RESOURCE Then
                        If isWalkableResource(GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index) + 1) Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_DOOR Then
                                    If Player(index).PlayerDoors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).state = 0 Then
                                            
                                            
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).DoorType = 0 Then
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).UnlockType = 0 Then
                                                    PlayerMsg index, "You need the right key to open this door. (" & Trim$(Item(Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).key).Name) & ")", BrightRed
                                                ElseIf Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).UnlockType = 1 Then
                                                    PlayerMsg index, "You have to flip a switch to open this door.", BrightRed
                                                Else
                                                    mapnum = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).WarpMap
                                                    x = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).WarpX
                                                    y = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).WarpY
                                                    Call PlayerWarp(index, mapnum, x, y)
                                                    Moved = YES
                                                End If
                                                End If
                                            
                                            PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                            Exit Sub
                                    Else
                                            If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).DoorType = 1 Then
                                                PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                                Exit Sub
                                            End If
                                    End If
                                End If
                                
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = YES) Then
                                    Call SetPlayerY(index, GetPlayerY(index) + 1)
                                    SendPlayerMove index, Movement, sendToSelf
                                    Moved = YES
                                End If
                            
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Down > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).Down).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                        'If map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
                        If isWalkableResource(GetPlayerMap(index), GetPlayerX(index) - 1, GetPlayerY(index)) Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
                                    If Player(index).PlayerDoors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).state = 0 Then
                                            
                                            
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).DoorType = 0 Then
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).UnlockType = 0 Then
                                                    PlayerMsg index, "You need the right key to open this door. (" & Trim$(Item(Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).key).Name) & ")", BrightRed
                                                ElseIf Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).UnlockType = 1 Then
                                                    PlayerMsg index, "You have to flip a switch to open this door.", BrightRed
                                                Else
                                                    mapnum = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).WarpMap
                                                    x = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).WarpX
                                                    y = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).WarpY
                                                    Call PlayerWarp(index, mapnum, x, y)
                                                    Moved = YES
                                                End If
                                                End If
                                            
                                            PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                            Exit Sub
                                    Else
                                            If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).DoorType = 1 Then
                                                PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                                Exit Sub
                                            End If
                                    End If
                                End If
                                
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = YES) Then
                                    Call SetPlayerX(index, GetPlayerX(index) - 1)
                                    SendPlayerMove index, Movement, sendToSelf
                                    Moved = YES
                                End If
                            
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).Left).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, NewMapX, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
            
        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < Map(mapnum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                        'If map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
                        If isWalkableResource(GetPlayerMap(index), GetPlayerX(index) + 1, GetPlayerY(index)) Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
                                    If Player(index).PlayerDoors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).state = 0 Then
                                            
                                            
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).DoorType = 0 Then
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).UnlockType = 0 Then
                                                    PlayerMsg index, "You need the right key to open this door. (" & Trim$(Item(Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).key).Name) & ")", BrightRed
                                                ElseIf Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).UnlockType = 1 Then
                                                    PlayerMsg index, "You have to flip a switch to open this door. ", BrightRed
                                                Else
                                                    mapnum = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).WarpMap
                                                    x = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).WarpX
                                                    y = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).WarpY
                                                    Call PlayerWarp(index, mapnum, x, y)
                                                    Moved = YES
                                                End If
                                                End If
                                            
                                            PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                            Exit Sub
                                    Else
                                            If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).DoorType = 1 Then
                                                PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                                Exit Sub
                                            End If
                                    End If
                                End If
                                
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = YES) Then
                                    Call SetPlayerX(index, GetPlayerX(index) + 1)
                                    SendPlayerMove index, Movement, sendToSelf
                                    Moved = YES
                                End If
                            
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Right > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).Right).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
    End Select
    
    With Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index))
    
        Call CheckBladeNPCMatch(index, GetPlayerMap(index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            mapnum = .Data1
            x = .Data2
            y = .Data3
            Call PlayerWarp(index, mapnum, x, y)
            Moved = YES
        End If
    
        ' Check to see if the tile is a door tile, and if so warp them
        If .Type = TILE_TYPE_DOOR Then
            doornum = .Data1
            
            If Player(index).PlayerDoors(doornum).state = 1 Then
                mapnum = Doors(doornum).WarpMap
                x = Doors(doornum).WarpX
                y = Doors(doornum).WarpY
                Call PlayerWarp(index, mapnum, x, y)
                Moved = YES
            End If
            
        End If
    
        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            x = .Data1
            y = .Data2
    
            If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                'Send to all players on the map
                SendMapKeyToMap GetPlayerMap(index), x, y, 1
                'SendMapKey index, X, Y, 1
                Call MapMsg(GetPlayerMap(index), "The door has been opened.", White)
                SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSwitchFloor, 1
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                    SendOpenShop index, x
                    TempPlayer(index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank index
            TempPlayer(index).InBank = True
            Moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            VitalType = .Data1
            amount = .Data2
            If Not GetPlayerVital(index, VitalType) = GetPlayerMaxVital(index, VitalType) Then
                If VitalType = Vitals.HP Then
                    Colour = BrightGreen
                Else
                    Colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(index), "+" & amount, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                SetPlayerVital index, VitalType, GetPlayerVital(index, VitalType) + amount
                PlayerMsg index, "You feel some forces that rejuvenate your body.", BrightGreen
                Call SendVital(index, VitalType)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            amount = .Data1
            SendActionMsg GetPlayerMap(index), "-" & amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
            If GetPlayerVital(index, HP) - amount <= 0 Then
                KillPlayer index
                PlayerMsg index, "you Have Died.", BrightRed
            Else
                SetPlayerVital index, HP, GetPlayerVital(index, HP) - amount
                PlayerMsg index, "You have been damaged.", BrightRed
                Call SendVital(index, HP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
        ' Slide
        If .Type = TILE_TYPE_SLIDE Then
            ForcePlayerMove index, MOVING_WALKING, .Data1
            Moved = YES
        End If
        
        ' Script Tile
        If .Type = TILE_TYPE_SCRIPT Then
            Scriptnum = .Data1
            Call ScriptTile(index, Scriptnum)
        End If
    End With

    ' They tried to hack
    If Moved = NO Then
        PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    End If

End Sub

Sub ForcePlayerMove(ByVal index As Long, ByVal Movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If Movement < 1 Or Movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove index, Direction, Movement, True
End Sub

Sub CheckEquippedItems(ByVal index As Long)
    Dim Slot As Long
    Dim itemnum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        itemnum = GetPlayerEquipment(index, i)

        If itemnum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If Item(itemnum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment index, 0, i
                Case Equipment.Armor

                    If Item(itemnum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment index, 0, i
                Case Equipment.Helmet

                    If Item(itemnum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment index, 0, i
                Case Equipment.Shield

                    If Item(itemnum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment index, 0, i
            End Select

        Else
            SetPlayerEquipment index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    If isItemStackable(itemnum) Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(index, i) = itemnum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    If Not IsPlaying(index) Then Exit Function
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = itemnum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function HasItem(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If isItemStackable(itemnum) Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If isItemStackable(itemnum) Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal index As Long, ByVal invSlot As Byte, ByVal ItemVal As Long, Optional ByVal Update As Boolean = False) As Boolean
    Dim i As Long
    Dim n As Long
    Dim itemnum As Integer
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then Exit Function
    
    itemnum = GetPlayerInvItemNum(index, invSlot)

    ' Prevent subscript out of range
    If itemnum < 1 Then Exit Function
    
    If isItemStackable(itemnum) Then
        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, invSlot, GetPlayerInvItemValue(index, invSlot) - ItemVal)
            
            ' Send the inventory update
            If Update Then
                Call SendInventoryUpdate(index, invSlot)
            End If
            Exit Function
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(index, invSlot, 0)
        Call SetPlayerInvItemValue(index, invSlot, 0)
        
        ' Send the inventory update
        If Update Then
            Call SendInventoryUpdate(index, invSlot)
        End If
    End If
End Function

Function GiveInvItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(index, itemnum)

    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, itemnum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        If sendUpdate Then Call SendInventoryUpdate(index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

End Function

Function HasSpell(ByVal index As Long, ByVal Spellnum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, i) = Spellnum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal index As Long)
    Dim i As Long
    Dim n As Long
    Dim mapnum As Long
    Dim Msg As String

    If Not IsPlaying(index) Then Exit Sub
    mapnum = GetPlayerMap(index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(mapnum, i).Num > 0) And (MapItem(mapnum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(mapnum, i).x = GetPlayerX(index)) Then
                    If (MapItem(mapnum, i).y = GetPlayerY(index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(index, MapItem(mapnum, i).Num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(index, n, MapItem(mapnum, i).Num)
    
                            If isItemStackable(GetPlayerInvItemNum(index, n)) Then
                                Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(mapnum, i).Value)
                                Msg = MapItem(mapnum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            End If
    
                            ' Erase item from the map
                            ClearMapItem i, mapnum
                            
                            Call SendInventoryUpdate(index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), 0, 0)
                            SendActionMsg GetPlayerMap(index), Msg, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                            'ALATAR
                            Call CheckTasks(index, QUEST_TYPE_GOGATHER, GetItemNum(Trim$(Item(GetPlayerInvItemNum(index, n)).Name)))
                            '/ALATAR
                            Exit For
                        Else
                            Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal index As Long, ByVal mapItemNum As Long)
Dim mapnum As Long

    mapnum = GetPlayerMap(index)
    
    ' no lock or locked to player?
    If MapItem(mapnum, mapItemNum).playerName = vbNullString Or MapItem(mapnum, mapItemNum).playerName = Trim$(GetPlayerName(index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal invNum As Long, ByVal amount As Long, Optional ByVal SayMsg As Boolean = True)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Or TempPlayer(index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(index, invNum) > 0) Then
        If (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
            i = FindOpenMapItemSlot(GetPlayerMap(index))

            If i <> 0 Then
                MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, invNum)
                MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
                MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
                MapItem(GetPlayerMap(index), i).playerName = Trim$(GetPlayerName(index))
                MapItem(GetPlayerMap(index), i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(index), i).canDespawn = True
                MapItem(GetPlayerMap(index), i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME

                If isItemStackable(GetPlayerInvItemNum(index, invNum)) Then

                    ' Check if its more then they have and if so drop it all
                    If amount >= GetPlayerInvItemValue(index, invNum) Then
                        MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, invNum)
                        If SayMsg Then Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, invNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(index, invNum, 0)
                        Call SetPlayerInvItemValue(index, invNum, 0)
                    Else
                        MapItem(GetPlayerMap(index), i).Value = amount
                        If SayMsg Then Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & amount & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(index, invNum, GetPlayerInvItemValue(index, invNum) - amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(index), i).Value = 0
                    ' send message
                    If SayMsg Then Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, invNum, 0)
                    Call SetPlayerInvItemValue(index, invNum, 0)
                End If

                ' Send inventory update
                Call SendInventoryUpdate(index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, amount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), Trim$(GetPlayerName(index)), MapItem(GetPlayerMap(index), i).canDespawn)
            Else
                If SayMsg Then Call PlayerMsg(index, "Too many items on the floor.", BrightRed)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    Dim points As Byte
    
    level_count = 0
    
    Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
        expRollover = GetPlayerExp(index) - GetPlayerNextLevel(index)
        
        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + 1) Then
            Exit Sub
        End If
        
        points = 3
        'Check if triforce
        points = points + GetPlayerTriforcesNum(index)
        
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + points)
        Call SetPlayerExp(index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg "ÃÃÂ¿ÃÂ½" & GetPlayerName(index) & " has risen " & level_count & "level!", Brown
        Else
            'plural
            GlobalMsg "ÃÂ¯ÃÂ¿ÃÂ½" & GetPlayerName(index) & " has risen " & level_count & " levels!", Brown
        End If
        SendEXP index
        SendPlayerData index
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seLevelUp, 1
    End If
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal index As Long) As String
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerLogin = Trim$(Player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerPassword = Trim$(Player(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    Player(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerName = Trim$(Player(index).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Name = Name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerClass = Player(index).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerSprite = Player(index).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerLevel = Player(index).Level
End Function

Function SetPlayerLevel(ByVal index As Long, ByVal Level As Long) As Boolean
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then Exit Function
    Player(index).Level = Level
    SetPlayerLevel = True
End Function

Function GetPlayerNextLevel(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(index) + 1 + GetPlayerTriforcesNum(index)) ^ 3 - (6 * (GetPlayerLevel(index) + 1 + GetPlayerTriforcesNum(index)) ^ 2) + 17 * (GetPlayerLevel(index) + 1 + GetPlayerTriforcesNum(index)) - 12)
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(index).Exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long)
    Player(index).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerAccess = Player(index).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerPK = Player(index).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).PK = PK
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerVital = Player(index).vital(vital)
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal vital As Vitals, ByVal Value As Long)
    Player(index).vital(vital) = Value
    If GetPlayerVital(index, vital) > GetPlayerMaxVital(index, vital) Then
        Player(index).vital(vital) = GetPlayerMaxVital(index, vital)
    End If

    If GetPlayerVital(index, vital) < 0 Then
        Player(index).vital(vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal index As Long, ByVal stat As Stats) As Long
    Dim x As Long, i As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    
    x = Player(index).stat(stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(index).Equipment(i) > 0 Then
            If Item(Player(index).Equipment(i)).Add_Stat(stat) > 0 Then
                x = x + Item(Player(index).Equipment(i)).Add_Stat(stat)
            End If
        End If
    Next
    
    GetPlayerStat = x
End Function

Public Function GetPlayerRawStat(ByVal index As Long, ByVal stat As Stats) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerRawStat = Player(index).stat(stat)
End Function

Public Sub SetPlayerStat(ByVal index As Long, ByVal stat As Stats, ByVal Value As Long)
    Player(index).stat(stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerPOINTS = Player(index).points
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal points As Long)
    If points <= 0 Then points = 0
    Player(index).points = points
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerMap = Player(index).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal mapnum As Long)

    If mapnum > 0 And mapnum <= MAX_MAPS Then
        Player(index).Map = mapnum
    End If

End Sub

Function GetPlayerX(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerX = Player(index).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    Player(index).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerY = Player(index).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    Player(index).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerDir = Player(index).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    If invSlot = 0 Then Exit Function
    GetPlayerInvItemNum = Player(index).Inv(invSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long, ByVal itemnum As Long)
    Player(index).Inv(invSlot).Num = itemnum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerInvItemValue = Player(index).Inv(invSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    Player(index).Inv(invSlot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal spellslot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerSpell = Player(index).Spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal spellslot As Long, ByVal Spellnum As Long)
    Player(index).Spell(spellslot) = Spellnum
End Sub

Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Player(index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    Player(index).Equipment(EquipmentSlot) = invNum
End Sub
Function GetPlayerVisible(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
GetPlayerVisible = Player(index).Visible
End Function
Sub SetPlayerVisible(ByVal index As Long, ByVal Visible As Long)
Player(index).Visible = Visible
End Sub

' ToDo
Sub OnDeath(ByVal index As Long, Optional ByVal RespawnSite As Byte = 0)
    Dim i As Long
    
    'Respawn Site = 0 if normal fluctuation (warp if map boot is defined), = 1 if always warp to initial site
    ' Set HP to nothing
    Call SetPlayerVital(index, Vitals.HP, 0)

    ' Drop all worn items
    'For i = 1 To Equipment.Equipment_Count - 1
        'If GetPlayerEquipment(Index, i) > 0 Then
            'PlayerMapDropItem Index, GetPlayerEquipment(Index, i), 0
        'End If
    'Next

    ' Warp player away
    Call SetPlayerDir(index, DIR_DOWN)
    
    With Map(GetPlayerMap(index))
        ' to the bootmap if it is set
        If .BootMap > 0 And RespawnSite = 0 Then
            PlayerWarp index, .BootMap, .BootX, .BootY
        Else
        Call PlayerWarp(index, Class(Player(index).Class).StartMap, Class(Player(index).Class).StartMapX, Class(Player(index).Class).StartMapY)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(index).spellBuffer.Spell = 0
    TempPlayer(index).spellBuffer.Timer = 0
    TempPlayer(index).spellBuffer.target = 0
    TempPlayer(index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(index)
    
TempPlayer(index).InBank = False
TempPlayer(index).InShop = 0
If TempPlayer(index).InTrade > 0 Then
For i = 1 To MAX_INV
TempPlayer(index).TradeOffer(i).Num = 0
TempPlayer(index).TradeOffer(i).Value = 0
TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).Num = 0
TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).Value = 0
Next

TempPlayer(index).InTrade = 0
TempPlayer(TempPlayer(index).InTrade).InTrade = 0

SendCloseTrade index
SendCloseTrade TempPlayer(index).InTrade
End If
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.mp, GetPlayerMaxVital(index, Vitals.mp))
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.mp)
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

    ' If the player the attacker killed was a pk then take it away
    'If GetPlayerPK(index) = YES Then
        'Call SetPlayerPK(index, NO)
        'Call SendPlayerData(index)
    'End If

End Sub

Sub CheckResource(ByVal index As Long, ByVal x As Long, ByVal y As Long)
        Dim Resource_num As Long
        Dim Resource_index As Long
        Dim rX As Long, rY As Long
        Dim i As Long
        Dim Damage As Long
        Dim Reward_index As Byte
   
        If Map(GetPlayerMap(index)).Tile(x, y).Type <> TILE_TYPE_RESOURCE Then Exit Sub
   
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(index)).Tile(x, y).Data1
        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
                If ResourceCache(GetPlayerMap(index)).ResourceData(i).x = x Then
                        If ResourceCache(GetPlayerMap(index)).ResourceData(i).y = y Then
                                Resource_num = i
                        End If
                End If
        Next
        If Resource_num > 0 Then
   
                If Resource(Resource_index).ToolRequired > 0 Then
                        If GetPlayerEquipment(index, Weapon) > 0 Then
                                If Item(GetPlayerEquipment(index, Weapon)).Data3 <> Resource(Resource_index).ToolRequired Then
                                        PlayerMsg index, "You don't have the right tool equipped.", BrightRed
                                        Exit Sub
                                Else
                                        Damage = RAND(1, Item(GetPlayerEquipment(index, Weapon)).Data2)
                                End If
                        Else
                                PlayerMsg index, "You need a tool", BrightRed
                                Exit Sub
                        End If
                Else
                        If GetPlayerEquipment(index, Weapon) > 0 Then
                                Damage = RAND(1, Item(GetPlayerEquipment(index, Weapon)).Data2)
                        Else
                                Damage = RAND(1, (GetPlayerStat(index, Stats.Strength) / 5))
                        End If
                End If
                   
                ' inv space?
                ' n = calculateresourcereward(Resource_index)
                'If Resource(Resource_index).ItemReward > 0 Then
                        'If FindOpenInvSlot(index, Resource(Resource_index).ItemReward) = 0 Then
                                'PlayerMsg index, "You don't have space in the inventory.", BrightRed
                                'Exit Sub
                        'End If
                'End If
                 'inv space?
                Reward_index = CalculateResourceRewardindex(Resource_index)
                If Reward_index > 0 Then
                        If FindOpenInvSlot(index, Resource(Resource_index).Rewards(Reward_index).ItemReward) = 0 Then
                                PlayerMsg index, "You don't have space in the inventory.", BrightRed
                                Exit Sub
                        End If
                End If
                ' check if already cut down
                If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 0 Then
                                   
                        rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).x
                        rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).y
                                ' check if damage is more than health
                                If Damage > 0 Then
                                        ' cut it down!
                                        If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                                                SendActionMsg GetPlayerMap(index), "-" & ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                                                SendResourceCacheToMap GetPlayerMap(index), Resource_num
                                                If Reward_index <> 0 Then
                                                ' send message if it exists
                                                If Resource(Resource_index).ItemSuccessMessage = False And Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                                        SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                                ElseIf Resource(Resource_index).ItemSuccessMessage = True And Resource(Resource_index).Rewards(Reward_index).ItemReward > 0 Then
                                                        SendActionMsg GetPlayerMap(index), "ÃÂ¯ÃÂ¿ÃÂ½" & Trim$(Item(Resource(Resource_index).Rewards(Reward_index).ItemReward).Name) & "!", BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                                End If
                                                ' carry on
                                                'GiveInvItem index, Resource(Resource_index).ItemReward, 1
                                                
                                                GiveInvItem index, Resource(Resource_index).Rewards(Reward_index).ItemReward, 1
                                                SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                                                End If
                                        Else
                                                ' just do the damage
                                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage
                                                SendActionMsg GetPlayerMap(index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                                                SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                                        End If
                                        ' send the sound
                                        SendMapSound index, rX, rY, SoundEntity.seResource, Resource_index
                                        'ALATAR
                                    Call CheckTasks(index, QUEST_TYPE_GOTRAIN, Resource_index)
                                        '/ALATAR
                                Else
                                        ' too weak
                                        SendActionMsg GetPlayerMap(index), "Failure!", BrightRed, 1, (rX * 32), (rY * 32)
                                End If
                        Else
                                ' send message if it exists
                                If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                                        SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                End If
                        End If
                End If
           
End Sub
Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Bank(index).Item(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal itemnum As Long)
    Bank(index).Item(BankSlot).Num = itemnum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Bank(index).Item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank(index).Item(BankSlot).Value = ItemValue
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal invSlot As Long, ByVal amount As Long)
Dim BankSlot

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, invSlot))
        
    If BankSlot > 0 Then
        If isItemStackable(GetPlayerInvItemNum(index, invSlot)) Then
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Long, ByVal amount As Long)
Dim invSlot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerBankItemValue(index, BankSlot) Then
        Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    If invSlot > 0 Then
        If isItemStackable(GetPlayerBankItemNum(index, BankSlot)) Then
            Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), amount)
            Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - amount)
            If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
            Else
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub
Public Sub KillPlayer(ByVal index As Long, Optional ByVal LoseExp As Byte = 0)
Dim Exp As Long

    If LoseExp = 0 Then

    ' Calculate exp to give attacker
    Exp = GetPlayerExp(index) \ 3

    ' Make sure we dont get less then 0
    If Exp < 0 Then Exp = 0
        If Exp = 0 Then
            Call PlayerMsg(index, "You have not lost any experience.", BrightRed)
        Else
            Call SetPlayerExp(index, GetPlayerExp(index) - Exp)
            SendEXP index
            Call PlayerMsg(index, "You've lost" & Exp & "of experience!", BrightRed)
        End If
    
    End If
    
    
    Call OnDeath(index)
End Sub

Public Sub UseItem(ByVal index As Long, ByVal invNum As Long)
Dim n As Long, i As Long, tempItem As Long, x As Long, y As Long, itemnum As Long, b As Long, j As Long
    
    For j = 1 To MAX_INV
    Next
    
    b = FindOpenInvSlot(index, j)
    
    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(index, invNum) > 0) And (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(index, invNum)).Data2
        itemnum = GetPlayerInvItemNum(index, invNum)
        
        ' Find out what kind of item it is
        Select Case Item(itemnum).Type
            Case ITEM_TYPE_ARMOR
            
                If CanPlayerEquipItem(index, itemnum) = False Then Exit Sub

                If GetPlayerEquipment(index, Armor) > 0 Then
                    tempItem = GetPlayerEquipment(index, Armor)
                End If

                SetPlayerEquipment index, itemnum, Armor
                PlayerMsg index, "You equip yourself" & Item(itemnum).Name, BrightGreen
                TakeInvItem index, itemnum, 0

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.mp)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_WEAPON
            
                If CanPlayerEquipItem(index, itemnum) = False Then Exit Sub
                
                If Item(itemnum).istwohander = True Then
                    If GetPlayerEquipment(index, Shield) > 0 Then
                        If GetPlayerEquipment(index, Weapon) > 0 Then
                            If b < 1 Then
                                Call PlayerMsg(index, "You don't have space in your inventory.", BrightRed)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                If Not CanPlayerEquipItem(index, Weapon) Then
                    Call PlayerMsg(index, "You can not equip this item", BrightRed)
                    Exit Sub
                End If
                    
                If GetPlayerEquipment(index, Weapon) > 0 Then
                    tempItem = GetPlayerEquipment(index, Weapon)
                End If
                

                SetPlayerEquipment index, itemnum, Weapon
                PlayerMsg index, "You equip yourself" & Item(itemnum).Name, BrightGreen
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                If Item(itemnum).istwohander = True Then
                If GetPlayerEquipment(index, Shield) > 0 Then
                GiveInvItem index, GetPlayerEquipment(index, Shield), 0
                SetPlayerEquipment index, 0, Shield
                End If
                End If
                
                
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.mp)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_HELMET
            
                If CanPlayerEquipItem(index, itemnum) = False Then Exit Sub

                If GetPlayerEquipment(index, Helmet) > 0 Then
                    tempItem = GetPlayerEquipment(index, Helmet)
                End If

                SetPlayerEquipment index, itemnum, Helmet
                PlayerMsg index, "You equip yourself" & CheckGrammar(Item(itemnum).Name), BrightGreen
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.mp)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_SHIELD
            
                If CanPlayerEquipItem(index, itemnum) = False Then Exit Sub

                If GetPlayerEquipment(index, Shield) > 0 Then
                    tempItem = GetPlayerEquipment(index, Shield)
                End If

                SetPlayerEquipment index, itemnum, Shield
                PlayerMsg index, "You equip yourself" & CheckGrammar(Item(itemnum).Name), BrightGreen
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.mp)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                If GetPlayerEquipment(index, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).istwohander = True Then
                GiveInvItem index, GetPlayerEquipment(index, Weapon), 0
                SetPlayerEquipment index, 0, Weapon
                End If
                End If
                
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            ' consumable
            Case ITEM_TYPE_CONSUME
                If TempPlayer(index).FreeAction = False Then
                Exit Sub
                End If
                
                If CanPlayerEquipItem(index, itemnum) = False Then Exit Sub
                
                ' add hp
                If Item(itemnum).AddHP > 0 Then
                    Player(index).vital(Vitals.HP) = Player(index).vital(Vitals.HP) + Item(itemnum).AddHP
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, HP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add mp
                If Item(itemnum).AddMP > 0 Then
                    Player(index).vital(Vitals.mp) = Player(index).vital(Vitals.mp) + Item(itemnum).AddMP
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, mp
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add exp
                If Item(itemnum).AddEXP > 0 Then
                    SetPlayerExp index, GetPlayerExp(index) + Item(itemnum).AddEXP
                    CheckPlayerLevelUp index
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendEXP index
                End If
                Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                Call TakeInvItem(index, Player(index).Inv(invNum).Num, 1)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            
            If Item(itemnum).ConsumeItem <> 0 Then
            GiveInvItem index, Item(itemnum).ConsumeItem, 1
            End If
            
            TempPlayer(index).FreeAction = False
            
            Case ITEM_TYPE_KEY
            
                If CanPlayerEquipItem(index, itemnum) = False Then Exit Sub
                
                Select Case GetPlayerDir(index)
                    Case DIR_UP

                        If GetPlayerY(index) > 0 Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) - 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_DOWN

                        If GetPlayerY(index) < Map(GetPlayerMap(index)).MaxY Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) + 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_LEFT

                        If GetPlayerX(index) > 0 Then
                            x = GetPlayerX(index) - 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                    Case DIR_RIGHT

                        If GetPlayerX(index) < Map(GetPlayerMap(index)).MaxX Then
                            x = GetPlayerX(index) + 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY Then

                    ' Check if the key they are using matches the map key
                    If itemnum = Map(GetPlayerMap(index)).Tile(x, y).Data1 Then
                        TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                        TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                        'SendMapKey index, X, Y, 1
                        SendMapKeyToMap GetPlayerMap(index), x, y, 1
                        Call MapMsg(GetPlayerMap(index), "The door has been opened.", White)
                        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSwitchFloor, 1
                        
                        Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, x, y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(index)).Tile(x, y).Data2 = 1 Then
                            Call TakeInvItem(index, itemnum, 0)
                            Call PlayerMsg(index, "The key was destroyed.", Yellow)
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_SPELL
            
                If CanPlayerEquipItem(index, itemnum) = False Then Exit Sub
                
                ' Get the spell num
                n = Item(itemnum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq

                        If i <= GetPlayerLevel(index) Then
                            i = FindOpenSpellSlot(index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(index, n) Then
                                    Call SetPlayerSpell(index, i, n)
                                    Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                                    Call TakeInvItem(index, itemnum, 0)
                                    Call PlayerMsg(index, "You've learned a new skill. Now you can use " & Trim$(Spell(n).Name) & ".", BrightGreen)
                                    Call SendPlayerSpells(index)
                                Else
                                    Call PlayerMsg(index, "You already know this skill.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(index, "You can't learn any more skills.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(index, "You must be level" & i & " to learn this skill.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(index, "It can only be learned by" & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
                
        Case ITEM_TYPE_RESET_POINTS
        
            i = ResetPlayerPoints(index)
            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + i)
            Call TakeInvItem(index, itemnum, 0)
            Call PlayerMsg(index, "You have reset your points, now you have " & i & " points!", BrightGreen)
            
        Case ITEM_TYPE_TRIFORCE
        
            'Triforce Type
            If Not GetPlayerLevel(index) >= MIN_LEVEL_TO_RESET Then
                PlayerMsg index, "Need to be Lvl" & MIN_LEVEL_TO_RESET & " to acquire the triforce", BrightRed
                Exit Sub
            End If
            
            SendOpenTriforce index
            ' send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum

        Case ITEM_TYPE_REDEMPTION
            
            If GetPlayerPK(index) = PK_PLAYER Then
                PlayerMsg index, "You have been redeemed, your past has been set free", BrightGreen
                Call PlayerRedemption(index)
                Call TakeInvItem(index, itemnum, 0)
                Call SendPlayerData(index)
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            End If
            
        End Select
    End If
End Sub

Sub CheckDoor(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim Door_num As Long
    Dim i As Long
    Dim n As Long
    Dim key As Long
    Dim tmpIndex As Long
    
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_DOOR Then
        Door_num = Map(GetPlayerMap(index)).Tile(x, y).Data1



        If Door_num > 0 Then
            If Doors(Door_num).DoorType = 0 Then
                If Player(index).PlayerDoors(Door_num).state = 0 Then
                    If Doors(Door_num).UnlockType = 0 Then
                        For i = 1 To MAX_INV
                            key = GetPlayerInvItemNum(index, i)
                            If Doors(Door_num).key = key Then
                                TakeInvItem index, key, 1
                                ' Previous method: SWITCH 1.00
                                ' Starting new method: access to all map players to the switch
                                For n = 1 To Player_HighIndex
                                    If GetPlayerMap(index) = GetPlayerMap(n) Then
                                        Player(n).PlayerDoors(Door_num).state = 1
                                        Player(n).PlayerDoors(Doors(Door_num).Switch).state = 1
                                        SendPlayerData (n)
                                        Call SendPlayerDoor(n, Doors(Door_num).Switch)
                                        PlayerMsg n, "Something has been unlocked", BrightBlue
                                        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSwitch, 1
                                    End If
                                Next
                                ' End
                                Exit Sub
                            End If
                        Next
                        PlayerMsg index, "You don't have the right key to open the door.", BrightBlue
                    ElseIf Doors(Door_num).UnlockType = 1 Then
                        If Doors(Door_num).state = 0 Then
                            PlayerMsg index, "You haven't pushed the right switch to open the door.", BrightBlue
                        End If
                    ElseIf Doors(Door_num).UnlockType = 2 Then
                        PlayerMsg index, "This door is not closed.", BrightBlue
                    End If
                    
                Else
                    PlayerMsg index, "This door is already open.", BrightBlue
                End If
            ElseIf Doors(Door_num).DoorType = 1 Then
                If Player(index).PlayerDoors(Door_num).state = 0 Then
                'previous method: SWITCH 1.01
                'start new method: Access all players on map to the switch
                    For n = 1 To Player_HighIndex
                        If GetPlayerMap(index) = GetPlayerMap(n) Then
                            Player(n).PlayerDoors(Door_num).state = 1
                            Player(n).PlayerDoors(Doors(Door_num).Switch).state = 1
                            SendPlayerData (n)
                            Call SendPlayerDoor(n, Doors(Door_num).Switch)
                            PlayerMsg n, "The switch has been activated", BrightBlue
                            SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSwitch, 1
                        End If
                    Next
                'end
                Else
                    'Previous method: SWITCH 1.02
                    'Starting new method: same as previous sub methods
                        For n = 1 To Player_HighIndex
                            If GetPlayerMap(index) = GetPlayerMap(n) Then
                                    Player(n).PlayerDoors(Door_num).state = 0
                                    Player(n).PlayerDoors(Doors(Door_num).Switch).state = 0
                                    SendPlayerData (n)
                                    Call SendPlayerDoor(n, Doors(Door_num).Switch)
                                    PlayerMsg n, "The switch has been turned off", BrightBlue
                                    SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSwitch, 1
                            End If
                        Next
                    'end
                End If
            End If
        End If
    End If
End Sub


Public Sub PlayerNavigation(ByVal index As Long)
'Dim Class As Long, x As Long, y As Long, newSprite As Long

' Seleccionar caso Clase del jugador
    ' Si la clase es zora, avanzar un paso y poner la nueva clase y sprite
    ' Si la clase no es zora
        '  se tiene el item bote, avanzar un paso y poner la nueva clase y sprite
        '  no se tiene el item bote, retroceder un paso para evitar que entren al mar
    ' Si la clase es de navegacion, volver al estado original usando el sprite del jugador actual
'fin del caso

'otras tareas:
'Desaquipar jugador? esconder items equipados? <- necesita variable booleana para identificar si esconder o no
'
End Sub

Public Function ResetPlayerPoints(ByVal index As Long) As Long
Dim i As Byte, sum As Long
ResetPlayerPoints = 0
'PlayerUnequip (index)
sum = 0

For i = 1 To Stats.Stat_Count - 1
    Do While Player(index).stat(i) > Class(GetPlayerClass(index)).stat(i)
        Player(index).stat(i) = Player(index).stat(i) - 1
        sum = sum + 1
    Loop
Next

ResetPlayerPoints = sum
    
End Function

Public Sub PlayerPVPDrops(ByVal index As Long)
Dim i As Long
Dim itemnum As Long
Dim ItemValue As Long

For i = 1 To MAX_INV
    
If GetPlayerInvItemNum(index, i) > 0 Then
    If IsItemDroppable(GetPlayerInvItemNum(index, i), index) Then 'check if dropable
        If isItemStackable(GetPlayerInvItemNum(index, i)) Then
            ItemValue = GetPlayerInvItemValue(index, i)
            If ItemValue > 0 Then
                'Drop 1 at least
                ItemValue = ItemValue * (1 / 10)
                If ItemValue = 0 Then ItemValue = 1
            
                Call PlayerMapDropItem(index, i, ItemValue, False)
            End If
        Else
            Call PlayerMapDropItem(index, i, 1, False)
        End If
    End If
End If
Next

For i = 1 To Equipment.Equipment_Count - 1
    If GetPlayerEquipment(GetPlayerInvItemNum(index, i), index) > 0 Then
        If IsItemDroppable(index, GetPlayerInvItemNum(index, i)) Then
            Call PlayerUnequipItemAndDrop(index, i)
        End If
    End If
Next

Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & "he dies and his items fall to the ground!", Yellow)

End Sub


Public Sub ResetPlayer(ByVal index As Long)
    Dim i As Long
    
    'pk
    Player(index).PK = NO
    'lvl
    Player(index).Level = 1
    
    'points
    Player(index).points = 0
    
    'exp
    Player(index).Exp = 0
    SendEXP (index)
    
    
    'inventory
    For i = 1 To MAX_INV
        Player(index).Inv(i).Num = 0
        Player(index).Inv(i).Value = 0
    Next
    Call SendInventory(index)
    
    'Equipment
    For i = 1 To Equipment.Equipment_Count - 1
        Player(index).Equipment(i) = 0
    Next
    SendWornEquipment index
    SendMapEquipment index
    
    'Quests
    For i = 1 To MAX_QUESTS
        Player(index).PlayerQuest(i).Status = 0
        Player(index).PlayerQuest(i).ActualTask = 0
        Player(index).PlayerQuest(i).CurrentCount = 0
    Next
    Call SendPlayerQuests(index)
    
    'Spells
    For i = 1 To MAX_PLAYER_SPELLS
        Player(index).Spell(i) = 0
    Next
    Call SendPlayerSpells(index)
    
    'npc info
    For i = 1 To MAX_NPCS
        Player(index).NPC(i).Kills = 0
    Next
    
    'hotbars
    For i = 1 To MAX_HOTBAR
        Player(index).Hotbar(i).Slot = 0
        Player(index).Hotbar(i).sType = 0
    Next
    Call SendHotbar(index)
    'stats
    For i = 1 To Stats.Stat_Count - 1
        Player(index).stat(i) = Class(GetPlayerClass(index)).stat(i)
    Next
    
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
        
    Call SendStats(index)
    
    Call SendPlayerData(index)

End Sub

Public Sub ComputePlayerReset(ByVal index As Long, ByVal triforce As TriforceType)
    Dim Colour As Byte
    Dim message As String
    Dim i As Byte
    Dim found As Boolean
    
    If Not IsPlaying(index) Then Exit Sub
    
    If Not GetPlayerLevel(index) >= 80 Then
        PlayerMsg index, "You must be lvl 80 like m minimum", BrightRed
        Exit Sub
    End If
    If GetPlayerTriforcesNum(index) > 0 Then
        PlayerMsg index, "You have already been reborn", BrightRed
        Exit Sub
    End If
    If GetPlayerTriforce(index, triforce) = True Then
        PlayerMsg index, "You already have that acquired Triforce", BrightRed
        Exit Sub
    End If
    
    found = False
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) > 0 Then
            If Item(Player(index).Inv(i).Num).Type = ITEM_TYPE_TRIFORCE Then
                found = True
                Player(index).Inv(i).Num = 0
                Player(index).Inv(i).Value = 0
                Call SendInventoryUpdate(index, i)
                Exit For
            End If
        End If
    Next
    
    If Not found Then
        PlayerMsg index, "You don't have the triforce", BrightRed
        Exit Sub
    End If
    
    Call ResetPlayer(index)
    Player(index).triforce(triforce) = True
    
    Select Case triforce
    Case TRIFORCE_COURAGE
        message = "of Courage"
        Colour = BrightGreen
    Case TRIFORCE_WISDOM
        message = "of Wisdom"
        Colour = BrightBlue
    Case TRIFORCE_POWER
        message = "of Power"
        Colour = BrightRed
    End Select
    
    
    For i = 1 To TriforceType.TriforceType_Count - 1
        If GetPlayerTriforce(index, i) = True Then
        Select Case i
            Case TRIFORCE_COURAGE
                SetPlayerStat index, Agility, GetPlayerStat(index, Agility) + 15
                SetPlayerStat index, endurance, GetPlayerStat(index, endurance) + 15
            Case TRIFORCE_WISDOM
                SetPlayerStat index, Intelligence, GetPlayerStat(index, Intelligence) + 15
                SetPlayerStat index, willpower, GetPlayerStat(index, willpower) + 15
            Case TRIFORCE_POWER
                SetPlayerStat index, Intelligence, GetPlayerStat(index, Intelligence) + 15
                SetPlayerStat index, Strength, GetPlayerStat(index, Strength) + 15
       End Select
       End If
    Next
    
    
    
    Call SendPlayerData(index)
    
    GlobalMsg GetPlayerName(index) & "he has acquired the triforce" & message, Colour
        

End Sub
Public Function GetPlayerTriforcesNum(ByVal index As Long) As Byte
Dim i As Byte
GetPlayerTriforcesNum = 0

For i = 1 To TriforceType.TriforceType_Count - 1
    If GetPlayerTriforce(index, i) = True Then
        GetPlayerTriforcesNum = GetPlayerTriforcesNum + 1
    End If
Next

End Function

Public Function GetPlayerTriforce(ByVal index As Long, ByVal triforce As TriforceType) As Boolean
Dim i As Byte
If Not IsPlaying(index) Then Exit Function

GetPlayerTriforce = False

If triforce > 0 And triforce < TriforceType_Count Then
    GetPlayerTriforce = Player(index).triforce(triforce)
End If

End Function

Public Function HasPlayerAnyTriforce(ByVal index As Long) As Boolean
HasPlayerAnyTriforce = False
Dim i As Byte

For i = 1 To TriforceType.TriforceType_Count - 1
    If GetPlayerTriforce(index, i) = True Then
        HasPlayerAnyTriforce = True
        Exit Function
    End If
Next
End Function

Public Function CanPlayerEquipItem(ByVal index As Long, ByVal itemnum As Long) As Boolean
Dim i As Byte

CanPlayerEquipItem = False

If Not (itemnum > 0 And itemnum <= MAX_ITEMS) Then Exit Function
' stat requirements
For i = 1 To Stats.Stat_Count - 1
    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
        PlayerMsg index, "You don't have the statistics necessary to equip yourself with this.", BrightRed
        Exit Function
    End If
Next
                
' level requirement
If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
    PlayerMsg index, "You don't have the necessary level to equip yourself with this one.", BrightRed
    Exit Function
End If
                
' class requirement
If Item(itemnum).ClassReq > 0 Then
    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
        PlayerMsg index, "You don't belong to the necessary class to equip yourself with this one.", BrightRed
        Exit Function
    End If
End If
                
' access requirement
If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
    PlayerMsg index, "You don't have the necessary access to equip yourself here.", BrightRed
    Exit Function
End If

'Triforce Requeriment
If Item(itemnum).BindType > 1 And Item(itemnum).BindType < 5 Then
    If Player(index).triforce(Item(itemnum).BindType - 1) = False Then
        PlayerMsg index, "You don't possess the triforce to equip yourself with this self", BrightRed
        Exit Function
    End If
ElseIf Item(itemnum).BindType = 5 Then
    If HasPlayerAnyTriforce(index) = False Then
        PlayerMsg index, "You must possess a triforce to equip yourself with it", BrightRed
        Exit Function
    End If
End If

CanPlayerEquipItem = True


End Function

Public Function IsPlayerNeutral(ByVal index As Long) As Boolean
IsPlayerNeutral = True
If index > 0 And index <= Player_HighIndex Then
    If Player(index).PK = YES Then
        IsPlayerNeutral = False
    End If
End If
End Function

Public Sub SetPlayerJustice(ByVal Killer As Long, ByVal Killed As Long)
If Not (Killer > 0 And Killer <= Player_HighIndex And Killed > 0 And Killed <= Player_HighIndex) Then Exit Sub

Select Case IsPlayerNeutral(Killer)
    Case True
        If IsPlayerNeutral(Killed) Then 'Player Killed Hero or Normal
            Player(Killer).PK = PK_PLAYER
            Player(Killer).PKPoints = Player(Killer).PKPoints + 1
            Call GlobalMsg(GetPlayerName(Killer) & " has become a murderer!", BrightRed)
        Else 'Player Killed PK
            If GetPlayerPK(Killer) = HERO_PLAYER Then
                Call GlobalMsg(GetPlayerName(Killer) & " has done justice!", Yellow)
            Else
                Player(Killer).PK = HERO_PLAYER
                Call GlobalMsg(GetPlayerName(Killer) & " has become a hero!", Yellow)
            End If
            Player(Killer).HeroPoints = Player(Killer).HeroPoints + 1
            
        End If
    Case False 'Killer Player is PK
        If IsPlayerNeutral(Killed) Then 'Add points in case of killed is neutral
            Call GlobalMsg(GetPlayerName(Killer) & " has committed a crime!", BrightRed)
            Player(Killer).PKPoints = Player(Killer).PKPoints + 1
        End If
End Select
End Sub

Public Sub PlayerRedemption(ByVal index As Long)
    'Special Punishments to the player
    Call SetPlayerPK(index, NONE_PLAYER)
End Sub

Public Function CheckSafeMode(ByVal attacker As Long, ByVal victim As Long) As Boolean
    'True: Player can't attack cause his safemode
    'False: Player can attack, if safe mode then victim is PK
    If IsPlayerNeutral(victim) Then
        If GetPlayerSafeMode(attacker) = True Then
            CheckSafeMode = True
        Else
            CheckSafeMode = False
        End If
    Else
        CheckSafeMode = False
    End If

End Function

Public Function GetPlayerSafeMode(ByVal index As Long) As Boolean

If TempPlayer(index).SafeMode = NO Then
    GetPlayerSafeMode = False
ElseIf TempPlayer(index).SafeMode = YES Then
    GetPlayerSafeMode = True
End If

End Function

Public Function GetPlayerNameColorByTriforce(ByVal index As Long) As Long

Dim color As Byte
Dim i As Byte

i = GetPlayerTriforcesNum(index)

'Normal Color
If i = 0 Then
    GetPlayerNameColorByTriforce = BrightGreen
    Exit Function
Else
    If GetPlayerTriforce(index, TRIFORCE_WISDOM) Then
        color = BrightBlue
    End If
    If GetPlayerTriforce(index, TRIFORCE_COURAGE) Then
        color = Green
    End If
    If GetPlayerTriforce(index, TRIFORCE_POWER) Then
        color = Red
    End If
End If


GetPlayerNameColorByTriforce = color

End Function

Public Function GetPlayerTriforcesName(ByVal index As Long) As String
Dim Chain As String
Dim i As Byte
Dim j As Byte
i = GetPlayerTriforcesNum(index)
Chain = ""
If i = 0 Then
    Chain = ""
Else
    For j = 1 To TriforceType.TriforceType_Count - 1
        If GetPlayerTriforce(index, j) = True Then
        Select Case j
            Case TriforceType.TRIFORCE_COURAGE
                Chain = Chain + "<Valor>"
            Case TriforceType.TRIFORCE_WISDOM
                Chain = Chain + "<Wisdom>"
            Case TriforceType.TRIFORCE_POWER
                Chain = Chain + "<Power>"
        End Select
        End If
    
    Next
End If

GetPlayerTriforcesName = Chain

End Function
