Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, X As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long, tmr100 As Long
    Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems As Long, LastUpdatePlayerVitals As Long

    ServerOnline = True

    Do While ServerOnline
        Tick = GetRealTickCount
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick
        
        If Tick > tmr100 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                        If GetRealTickCount > TempPlayer(i).spellBuffer.Timer + (Spell(player(i).Spell(TempPlayer(i).spellBuffer.Spell)).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.Target, TempPlayer(i).spellBuffer.tType
                            TempPlayer(i).spellBuffer.Spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.Target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If

                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < GetRealTickCount Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For X = 1 To MAX_DOTS
                        HandleDoT_Player i, X
                        HandleHoT_Player i, X
                    Next
                    
                    Call CheckPetActions(i)
                    
                    If Not TempPlayer(i).MovementsStack Is Nothing Then
                        If Not TempPlayer(i).MovementsStack.IsEmpty Then
                            ForcePlayerMove i, 1, TempPlayer(i).MovementsStack.Front.GetDir
                            TempPlayer(i).MovementsStack.Pop
                        End If
                    End If
                        
                End If
            Next
            
            CheckNPCSMovement Tick
            
            frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
            tmr100 = GetRealTickCount + 100
        End If

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).state > sckConnected Then
                    Call CloseSocket(i)
                End If
                
                
            Next
            
            Dim tickmap0 As Long
            tickmap0 = GetRealTickCount
            UpdateMapLogic
            tickmap0 = GetRealTickCount - tickmap0
            frmServer.lblMapTime.Caption = "MapUpdate(ms): " & tickmap0
            
            
            
            tmr500 = GetRealTickCount + 500
        End If

        If Tick > tmr1000 Then
            
            For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If TempPlayer(i).FreeAction = False Then TempPlayer(i).FreeAction = True
                
                'Handles Guild Invites
                If TempPlayer(i).tmpGuildInviteSlot > 0 Then
                    If Tick > TempPlayer(i).tmpGuildInviteTimer Then
                        If GuildData(TempPlayer(i).tmpGuildInviteSlot).In_Use = True Then
                            PlayerMsg i, GetTranslation("El tiempo de aceptar invitación al clan se agotó") & " " & GuildData(TempPlayer(i).tmpGuildInviteSlot).Guild_Name & ".", BrightRed, , False
                            TempPlayer(i).tmpGuildInviteSlot = 0
                            TempPlayer(i).tmpGuildInviteTimer = 0
                        Else
                            'Just remove this guild has been unloaded
                            TempPlayer(i).tmpGuildInviteSlot = 0
                            TempPlayer(i).tmpGuildInviteTimer = 0
                        End If
                    End If
                End If
                
                TempPlayer(i).InactiveTime = TempPlayer(i).InactiveTime + 1
                If GetInactiveTime(i) >= 900 And GetPlayerAccess_Mode(i) < ADMIN_MONITOR Then
                    AlertMsg i, "has sido expulsado por inactividad"
                End If
                
                CheckPlayerStatsBuffer i, Tick
                
                CheckPlayerActionsProtections i
                CheckPlayerActions i, Tick
                CheckPlayerProtections i, Tick
                
            End If
            Next
                
            If isShuttingDown Then
                Call HandleShutdown
            End If
            
            ClearQuestions
            
            CheckMutedPlayers Tick
            
            CheckWaitingNPCS Tick
            
            
            tmr1000 = GetRealTickCount + 500
        End If
        
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                For X = 1 To MAX_PLAYER_PROJECTILES
                    If TempPlayer(i).ProjecTile(X).Pic > 0 Then
                        ' handle the projec tile
                        HandleProjecTile i, X
                    End If
                Next
            End If
        Next
        
        If Tick > FloodTimer Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    Call CheckFlood(i)
                End If
            Next
            FloodTimer = GetRealTickCount + FLOOD_LAPSE * 1000
        End If
        
        
        If Tick > KillPointsTimer Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    Call AddPlayerPointsByTimeLapse(i)
                End If
            Next
            KillPointsTimer = GetRealTickCount + KILL_POINTS_LAPSE * 1000
        End If
        
        If Tick > NeutralPlayerPointsTimer Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    Call AddNeutralPlayerPoints(i)
                End If
            Next
            NeutralPlayerPointsTimer = NeutralPlayerPointsTimer + NEUTRAL_POINTS_LAPSE * 1000
        End If

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetRealTickCount + 5000
            
            'Dim m As clsMap
            'Set m = New clsMap
            
            'm.ReadData
        End If

        ' Checks to save players every 5 minutes - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            ClearIPTries
            LastUpdateSavePlayers = GetRealTickCount + 150000
        End If
        
        
        'Checks if is necesary to change weather
        If Tick > LastWeatherUpdate And WeatherTime <> 0 Then
            'Sentinell
            If WeatherProbability <= 100 Then
                Select Case RAND(1, 100)
                Case Is <= CLng(WeatherProbability)
                    ActivateWeather
                Case Else
                    DisableWeather
                End Select
            ElseIf WeatherProbability = 101 Then
                UpdateWeather
            End If
            LastWeatherUpdate = GetRealTickCount + WeatherTime
        End If
        
        frmServer.lblLoopTime.Caption = "Loop(ms): " & GetRealTickCount - Tick
        IsServerBug = False
        

        If Not CPSUnlock Then Sleep SleepTime
                DoEvents
                
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS \ 4
            TickCPS = Tick + 4000
            CPS = 0
            UpdateTrafficStadistics
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Sub CheckPetActions(ByVal index As Long)
    'Attack targets
    Dim X As Long, mapnum As Long, Target As Long, TargetType As Byte
    X = PlayerHasPetInMap(index)
    
    If X > 0 Then
        mapnum = GetPlayerMap(index)
        If MapNpc(mapnum).NPC(X).Num > 0 Then
            Target = MapNpc(mapnum).NPC(X).Target
            TargetType = MapNpc(mapnum).NPC(X).TargetType
        
            ' Check if the npc can attack the targeted player player
            If Target > 0 Then
                If TargetType = TARGET_TYPE_PLAYER Then ' player
        
                    ' Is the target playing and on the same map?
                    If IsPlaying(Target) And GetPlayerMap(Target) = mapnum Then
                        TryNpcAttackPlayer X, Target
                    Else
                        ' Player left map or game, set target to 0
                        MapNpc(mapnum).NPC(X).Target = 0
                        MapNpc(mapnum).NPC(X).TargetType = 0 ' clear
                        PetFollowOwner MapNpc(mapnum).NPC(X).PetData.Owner
                    End If
                ElseIf TargetType = TARGET_TYPE_NPC Then 'npc
                        'Npc vs Npc
                        Call TryNpcAttackNpc(mapnum, X, MapNpc(mapnum).NPC(X).Target)
                End If
            End If
        End If
    End If
End Sub


Private Sub UpdateMapLogic()
    Dim i As Long, X As Long, mapnum As Long, N As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, npcnum As Long
    Dim Target As Long, TargetType As Byte, DidWalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim TargetX As Long, TargetY As Long, target_verify As Boolean
    Dim MAP_HIGH_NPC As Long, MAP_HIGH_ITEM As Long
    
    'For mapnum = 1 To MAX_MAPS
    For CurrentMapIndex = 1 To GetNumberOfMapsWithPlayers
    
    mapnum = GetMapNumByMapReference(CurrentMapIndex)
    If mapnum = 0 Then
        Exit Sub
    End If
    If TempMap(mapnum).Exists Then
    
            'improving speed
            MAP_HIGH_NPC = TempMap(mapnum).npc_highindex
            MAP_HIGH_ITEM = TempMap(mapnum).Item_highindex
            ' items appearing to everyone
            For i = 1 To MAP_HIGH_ITEM
                With MapItem(mapnum, i)
                If .Num > 0 Then
                    If .playerName <> vbNullString Then
                        ' make item public?
                        If .playerTimer < GetRealTickCount Then
                            ' make it public
                            .playerName = vbNullString
                            .playerTimer = 0
                            ' send updates to everyone
                            SendSpawnItemToMap mapnum, i
                        End If
                    Else
                        ' despawn item?
                        If .isDrop Then
                            If .Timer < GetRealTickCount Then
                                ' despawn it
                                ClearMapItem i, mapnum
                                ' send updates to everyone
                                SendSpawnItemToMap mapnum, i
                            End If
                        End If
                    End If
                End If
                End With
            Next
            
            'check if there are items to spawn
            CheckMapWaitingItem (mapnum)
        
            '  Close the doors
            'Extremly ineficient method, used door timer sentinell in order to save X And Y testing
            Dim MapNumDoors As Integer
            MapNumDoors = TempTile(mapnum).NumDoors
            
            While MapNumDoors > 0
                With TempTile(mapnum).Door(MapNumDoors)
                If .DoorTimer <> 0 And GetRealTickCount > .DoorTimer Then
                    .DoorTimer = 0
                    .state = Not (.state)
                    If CanRenderTempDoor(mapnum, MapNumDoors) Then
                        SendMapKeyToMap mapnum, .X, .Y, .state
                    End If
                Else 'check if weight switch
                    If GetDoorType(.DoorNum) = DOOR_TYPE_WEIGHTSWITCH Then
                        If IsDoorOpened(mapnum, MapNumDoors) Then
                            If Not IsSomebodyOnSwitch(mapnum, MapNumDoors) Then
                                CheckWeightSwitch mapnum, MapNumDoors
                            End If
                        End If
                    End If
                End If
                MapNumDoors = MapNumDoors - 1
                End With
            Wend
            
        
            ' check for DoTs + hots
            'For i = 1 To MAX_MAP_NPCS
            For i = 1 To MAP_HIGH_NPC
                If MapNpc(mapnum).NPC(i).Num > 0 Then
                    For X = 1 To MAX_DOTS
                        HandleDoT_Npc mapnum, i, X
                        HandleHoT_Npc mapnum, i, X
                    Next
                End If
            Next

        

            TickCount = GetRealTickCount
            
            'For X = 1 To MAX_MAP_NPCS
            For X = 1 To MAP_HIGH_NPC
                npcnum = MapNpc(mapnum).NPC(X).Num

                ' Check Attack on Sight
                ' Make sure theres a npc with the map
                If npcnum > 0 Then
                
                    'Test On Moment Actions
                    

                    ' If the npc is a attack on sight, search for a player on the map
                    If NPC(npcnum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(npcnum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(mapnum).NPC(X).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = mapnum And MapNpc(mapnum).NPC(X).Target = 0 And GetPlayerAccess_Mode(i) <= ADMIN_MONITOR Then
                                        N = NPC(npcnum).range
                                        DistanceX = MapNpc(mapnum).NPC(X).X - GetPlayerX(i)
                                        DistanceY = MapNpc(mapnum).NPC(X).Y - GetPlayerY(i)
    
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
    
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= N And DistanceY <= N Then
                                            If NPC(npcnum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                If Len(Trim$(NPC(npcnum).AttackSay)) > 0 Then
                                                    'Call PlayerMsg(i, Trim$(NPC(npcnum).Name) & ": " & Trim$(NPC(npcnum).AttackSay), SayColor)
                                                    
                                                    Call SendActionMsg(mapnum, GetTranslation(NPC(npcnum).AttackSay), SayColor, 1, MapNpc(mapnum).NPC(X).X * 32, MapNpc(mapnum).NPC(X).Y * 32)
                                                End If
                                                MapNpc(mapnum).NPC(X).TargetType = 1 ' player
                                                MapNpc(mapnum).NPC(X).Target = i
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                
                target_verify = False

                If TickCount > MapNpc(mapnum).NPC(X).StunTimer + (MapNpc(mapnum).NPC(X).StunDuration * 1000) Then
                    MapNpc(mapnum).NPC(X).StunDuration = 0
                    MapNpc(mapnum).NPC(X).StunTimer = 0
                End If
                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                'If mapnpc(mapnum).NPC(x).Num > 0 Then
                    'If mapnpc(mapnum).NPC(x).StunDuration > 0 Then
                        ' check if we can unstun them
                        'If GetRealTickCount > mapnpc(mapnum).NPC(x).StunTimer + (mapnpc(mapnum).NPC(x).StunDuration * 1000) Then
                            'mapnpc(mapnum).NPC(x).StunDuration = 0
                            'mapnpc(mapnum).NPC(x).StunTimer = 0
                        'End If
                    'Else
                        'If mapnpc(mapnum).NPC(x).Target > 0 Then
                            'If GetMapPetOwner(mapnum, x) = 0 Then
                                'Call ComputeNPCTargetMovement(mapnum, x)
                            'End If
                        'Else 'Non target movement
                            'Call ComputeNPCNonTargetMovement(mapnum, x)
                        'End If
                    'End If
                'End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapNpc(mapnum).NPC(X).Num > 0 Then
                    Target = MapNpc(mapnum).NPC(X).Target
                    TargetType = MapNpc(mapnum).NPC(X).TargetType

                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                    
                        If TargetType = 1 Then ' player

                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) And GetPlayerMap(Target) = mapnum Then
                                TryNpcAttackPlayer X, Target
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(mapnum).NPC(X).Target = 0
                                MapNpc(mapnum).NPC(X).TargetType = 0 ' clear
                                PetFollowOwner MapNpc(mapnum).NPC(X).PetData.Owner
                            End If
                        ElseIf TargetType = 2 Then 'npc
                                'Npc vs Npc
                                Call TryNpcAttackNpc(mapnum, X, MapNpc(mapnum).NPC(X).Target)
                        End If
                        
                    End If
                    
                    ' Spell Casting
                    For i = 1 To MAX_NPC_SPELLS
                        If NPC(npcnum).Spell(i) > 0 Then
                            If MapNpc(mapnum).NPC(X).SpellTimer(i) + (Spell(NPC(npcnum).Spell(i)).CastTime * 1000) < GetRealTickCount Then
                                
                                If Not ChoosePetSpellingMethod(GetMapPetOwner(mapnum, X), X, i, NPC(npcnum).Spell(i)) Then
                                    'NPC can not autoheal, so find out what kind of magic should it invoque
                                    Select Case TargetType
                                    Case 1 'Player
                                        NpcSpellPlayer X, Target, i
                                    Case 2 'Npc
                                        NpcSpellNpc mapnum, X, Target, i
                                    End Select
                                End If
                            End If
                        End If
                    Next
                End If
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's Vitals //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's Vitals
                If Not MapNpc(mapnum).NPC(X).stopRegen Then
                    If MapNpc(mapnum).NPC(X).Num > 0 And TickCount > GiveNPCHPTimer + 5000 Then 'every 5 seconds
                        For i = 1 To Vitals.Vital_Count - 1
                            If MapNpc(mapnum).NPC(X).vital(i) > 0 Or i = Vitals.MP Then
                                MapNpc(mapnum).NPC(X).vital(i) = MapNpc(mapnum).NPC(X).vital(i) + GetNpcVitalRegen(mapnum, X, i)
    
                                ' Check if they have more then they should and if so just set it to max
                                If MapNpc(mapnum).NPC(X).vital(i) > GetNpcMaxVital(mapnum, X, i) Then
                                    MapNpc(mapnum).NPC(X).vital(i) = GetNpcMaxVital(mapnum, X, i)
                                End If
                                                       
                            End If
                        Next
                        SendMapNpcVitals mapnum, X
                    End If
                Else 'Check if we need to set the Regen Variable ON
                    If TickCount > MapNpc(mapnum).NPC(X).stopRegenTimer + 5000 Then
                        'Turn on
                        MapNpc(mapnum).NPC(X).stopRegen = False
                    End If
                End If


                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                'If mapnpc(mapnum).NPC(X).Num = 0 And mapnpc(mapnum).NPC(X).MapNPCNum > 0 Then
                    'If TickCount > mapnpc(mapnum).NPC(X).SpawnWait + (NPC(map(mapnum).NPC(mapnpc(mapnum).NPC(X).MapNPCNum)).SpawnSecs * 1000) Then
                        'Call SpawnNpc(mapnpc(mapnum).NPC(X).MapNPCNum, mapnum)
                    'End If
                'End If
                ' //////////////////////////////////////
                ' // This is used for dispawning an NPC //
                ' //////////////////////////////////////
                If IsTempNPC(mapnum, X) And MapNpc(mapnum).NPC(X).Num <> NPC_SKULLTULA Then
                    If TickCount > MapNpc(mapnum).NPC(X).SpawnWait Then 'li'l cheat, won't affect skutullas
                        Call SendClearMapNpcToMap(mapnum, X)
                        Call ClearSingleMapNpc(X, mapnum)
                        'Call SendMapNpcToMap(mapnum, X)
                    End If
                End If
            Next
            
        End If
        
        ' Respawning Resources
        If ResourceCache(mapnum).Resource_Count > 0 Then
            For i = 1 To ResourceCache(mapnum).Resource_Count
                Resource_index = map(mapnum).Tile(ResourceCache(mapnum).ResourceData(i).X, ResourceCache(mapnum).ResourceData(i).Y).Data1

                If Resource_index > 0 And Resource_index <= MAX_RESOURCES Then
                    If ResourceCache(mapnum).ResourceData(i).ResourceState = 1 Then   ' dead or fucked up
                        If ResourceCache(mapnum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < GetRealTickCount Then
                            ResourceCache(mapnum).ResourceData(i).ResourceTimer = GetRealTickCount
                            ResourceCache(mapnum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(mapnum).ResourceData(i).cur_health = Resource(Resource_index).health
                            SendSingleResourceCacheToMap mapnum, i
                        End If
                    End If
                End If
            Next
        'End If
    End If
        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If TickCount > GiveNPCHPTimer + 5000 Then
        GiveNPCHPTimer = GetRealTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetRealTickCount > KeyTimer + 15000 Then
        KeyTimer = GetRealTickCount
    End If

End Sub

Private Sub UpdatePlayerVitals()
Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not TempPlayer(i).stopRegen Then
                If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                    Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                    Call SendVital(i, Vitals.HP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
    
                If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                    Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                    Call SendVital(i, Vitals.MP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
    Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd(GetTranslation("Guardando jugadores en línea..."))

        For i = 1 To Player_HighIndex

            If IsPlaying(i) Then
                Call SavePlayer(i)
                Call SaveBank(i)
            End If

            DoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 10 Then
        Call GlobalMsg("Automated Server Shutdown in " & Secs & " seconds.", Cyan, False)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Apagado.", BrightRed)
        Call DestroyServer
        'Call RestartServer
    End If

End Sub

Public Function GetRealTickCount() As Long
    If GetTickCount < 0 Then
        GetRealTickCount = GetTickCount + MAX_LONG
    Else
        GetRealTickCount = GetTickCount
    End If
End Function

Sub UpdateTrafficStadistics()
    frmServer.lblPacketsReceived.Caption = "Packets Received / Second: " & PacketsReceived
    frmServer.lblPacketsSent.Caption = "Packets Sent / Second: " & PacketsSent
    frmServer.lblBytesSent.Caption = "Bytes Sent / Second: " & BytesSent
    frmServer.lblBytesReceived.Caption = "Bytes Received / Second: " & BytesReceived
    
    PacketsReceived = 0
    PacketsSent = 0
    BytesSent = 0
    BytesReceived = 0

End Sub



