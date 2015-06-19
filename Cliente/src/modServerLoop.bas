Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, x As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long
    Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems As Long, LastUpdatePlayerVitals As Long

    ServerOnline = True

    Do While ServerOnline
        Tick = GetTickCount
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick
        
        If Tick > tmr25 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                        If GetTickCount > TempPlayer(i).spellBuffer.Timer + (Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell)).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.target, TempPlayer(i).spellBuffer.tType
                            TempPlayer(i).spellBuffer.Spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If GetTickCount > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < GetTickCount Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For x = 1 To MAX_DOTS
                        HandleDoT_Player i, x
                        HandleHoT_Player i, x
                    Next
                End If
            Next
            frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
            tmr25 = GetTickCount + 25
        End If

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).state > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            UpdateMapLogic
            tmr500 = GetTickCount + 500
        End If

            If Tick > tmr1000 Then
            
            For i = 1 To Player_HighIndex
        If TempPlayer(i).FreeAction = False Then TempPlayer(i).FreeAction = True
            Next
                
            If isShuttingDown Then
                Call HandleShutdown
            End If
            tmr1000 = GetTickCount + 1000
        End If
        
                For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                For x = 1 To MAX_PLAYER_PROJECTILES
                    If TempPlayer(i).ProjecTile(x).Pic > 0 Then
                        ' handle the projec tile
                        HandleProjecTile i, x
                    End If
                Next
            End If
        Next

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCount + 5000
        End If

        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If Tick > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = GetTickCount + 300000
        End If

        ' Checks to save players every 5 minutes - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCount + 300000
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
            LastWeatherUpdate = GetTickCount + WeatherTime
        End If
        
'Handles Guild Invites
For i = 1 To Player_HighIndex
If IsPlaying(i) Then
If TempPlayer(i).tmpGuildInviteSlot > 0 Then
If Tick > TempPlayer(i).tmpGuildInviteTimer Then
If GuildData(TempPlayer(i).tmpGuildInviteSlot).In_Use = True Then
PlayerMsg i, "El tiempo de aceptar invitación al clan se agotó" & GuildData(TempPlayer(i).tmpGuildInviteSlot).Guild_Name & ".", BrightRed
TempPlayer(i).tmpGuildInviteSlot = 0
TempPlayer(i).tmpGuildInviteTimer = 0
Else
'Just remove this guild has been unloaded
TempPlayer(i).tmpGuildInviteSlot = 0
TempPlayer(i).tmpGuildInviteTimer = 0
End If
End If
End If
End If

Next i
        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Private Sub UpdateMapSpawnItems()
    Dim x As Long
    Dim y As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For y = 1 To MAX_MAPS

        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(y) Then

            ' Clear out unnecessary junk
            For x = 1 To MAX_MAP_ITEMS
                Call ClearMapItem(x, y)
            Next

            ' Spawn the items
            Call SpawnMapItems(y)
            Call SendMapItemsToAll(y)
        End If

        DoEvents
    Next

End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, x As Long, Mapnum As Long, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long
    Dim target As Long, targetType As Byte, DidWalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim TargetX As Long, TargetY As Long, target_verify As Boolean

    For Mapnum = 1 To MAX_MAPS
        ' items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(Mapnum, i).Num > 0 Then
                If MapItem(Mapnum, i).playerName <> vbNullString Then
                    ' make item public?
                    If MapItem(Mapnum, i).playerTimer < GetTickCount Then
                        ' make it public
                        MapItem(Mapnum, i).playerName = vbNullString
                        MapItem(Mapnum, i).playerTimer = 0
                        ' send updates to everyone
                        SendMapItemsToAll Mapnum
                    End If
                    ' despawn item?
                    If MapItem(Mapnum, i).canDespawn Then
                        If MapItem(Mapnum, i).despawnTimer < GetTickCount Then
                            ' despawn it
                            ClearMapItem i, Mapnum
                            ' send updates to everyone
                            SendMapItemsToAll Mapnum
                        End If
                    End If
                End If
            End If
        Next
        
        
        
        If PlayersOnMap(Mapnum) = YES Then
        '  Close the doors
            'Extremly ineficient method, used door timer sentinell in order to save X And Y testing
            If TempTile(Mapnum).DoorTimer <> 0 And GetTickCount > TempTile(Mapnum).DoorTimer + 60000 Then
            For x1 = 0 To Map(Mapnum).MaxX
                For y1 = 0 To Map(Mapnum).MaxY
                    If Map(Mapnum).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(Mapnum).DoorOpen(x1, y1) = YES Then
                        TempTile(Mapnum).DoorOpen(x1, y1) = NO
                        SendMapKeyToMap Mapnum, x1, y1, 0
                        TempTile(Mapnum).DoorTimer = 0
                    End If
                Next
            Next
            End If
        End If
        
        ' check for DoTs + hots
        For i = 1 To MAX_MAP_NPCS
            If mapnpc(Mapnum).NPC(i).Num > 0 Then
                For x = 1 To MAX_DOTS
                    HandleDoT_Npc Mapnum, i, x
                    HandleHoT_Npc Mapnum, i, x
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(Mapnum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(Mapnum).Resource_Count
                Resource_index = Map(Mapnum).Tile(ResourceCache(Mapnum).ResourceData(i).x, ResourceCache(Mapnum).ResourceData(i).y).Data1

                If Resource_index > 0 And Resource_index <= MAX_RESOURCES Then
                    If ResourceCache(Mapnum).ResourceData(i).ResourceState = 1 Or ResourceCache(Mapnum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(Mapnum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < GetTickCount Then
                            ResourceCache(Mapnum).ResourceData(i).ResourceTimer = GetTickCount
                            ResourceCache(Mapnum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(Mapnum).ResourceData(i).cur_health = Resource(Resource_index).health
                            SendResourceCacheToMap Mapnum, i
                        End If
                    End If
                End If
            Next
        End If

        If PlayersOnMap(Mapnum) = YES Then
            TickCount = GetTickCount
            
            For x = 1 To MAX_MAP_NPCS
                NpcNum = mapnpc(Mapnum).NPC(x).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Mapnum).NPC(x) > 0 And mapnpc(Mapnum).NPC(x).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not mapnpc(Mapnum).NPC(x).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = Mapnum And mapnpc(Mapnum).NPC(x).target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        n = NPC(NpcNum).Range
                                        DistanceX = mapnpc(Mapnum).NPC(x).x - GetPlayerX(i)
                                        DistanceY = mapnpc(Mapnum).NPC(x).y - GetPlayerY(i)
    
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
    
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                If Len(Trim$(NPC(NpcNum).AttackSay)) > 0 Then
                                                    'Call PlayerMsg(i, Trim$(NPC(npcnum).Name) & ": " & Trim$(NPC(npcnum).AttackSay), SayColor)
                                                    
                                                    Call SendActionMsg(Mapnum, Trim$(NPC(NpcNum).AttackSay), SayColor, 1, mapnpc(Mapnum).NPC(x).x * 32, mapnpc(Mapnum).NPC(x).y * 32)
                                                End If
                                                mapnpc(Mapnum).NPC(x).targetType = 1 ' player
                                                mapnpc(Mapnum).NPC(x).target = i
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Mapnum).NPC(x) > 0 And mapnpc(Mapnum).NPC(x).Num > 0 Then
                    If mapnpc(Mapnum).NPC(x).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > mapnpc(Mapnum).NPC(x).StunTimer + (mapnpc(Mapnum).NPC(x).StunDuration * 1000) Then
                            mapnpc(Mapnum).NPC(x).StunDuration = 0
                            mapnpc(Mapnum).NPC(x).StunTimer = 0
                        End If
                    Else
                            
                        target = mapnpc(Mapnum).NPC(x).target
                        targetType = mapnpc(Mapnum).NPC(x).targetType
    
                        ' Check to see if its time for the npc to walk
                        If NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER And NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_SLIDE Then
                        
                            If targetType = 1 Then ' player
    
                                ' Check to see if we are following a player or not
                                If target > 0 Then
        
                                    ' Check if the player is even playing, if so follow'm
                                    If IsPlaying(target) And GetPlayerMap(target) = Mapnum Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = GetPlayerY(target)
                                        TargetX = GetPlayerX(target)
                                        'Check if Player Pet has to defend his owner
                                        If TempPlayer(target).TempPetSlot > 0 And TempPlayer(target).TempPetSlot <> x Then
                                            If TempPlayer(target).PetHasOwnTarget = NO Then
                                                'Pet has not a target, let's catch this npc
                                                TempPlayer(target).PetHasOwnTarget = x
                                                mapnpc(Mapnum).NPC(TempPlayer(target).TempPetSlot).targetType = TARGET_TYPE_NPC
                                                mapnpc(Mapnum).NPC(TempPlayer(target).TempPetSlot).target = x
                                            End If
                                        End If
                                    Else
                                        mapnpc(Mapnum).NPC(x).targetType = 0 ' clear
                                        mapnpc(Mapnum).NPC(x).target = 0
                                    End If
                                End If
                            
                            ElseIf targetType = 2 Then 'npc
                                
                                If target > 0 Then
                                    
                                    If mapnpc(Mapnum).NPC(target).Num > 0 Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = mapnpc(Mapnum).NPC(target).y
                                        TargetX = mapnpc(Mapnum).NPC(target).x
                                    Else
                                        mapnpc(Mapnum).NPC(x).targetType = 0 ' clear
                                        mapnpc(Mapnum).NPC(x).target = 0
                                    End If
                                End If
                            End If
                            
                            If target_verify Then
                                
                                i = Int(Rnd * 5)
    
                                ' Lets move the npc
                                Select Case i
                                    Case 0
    
                                        ' Up
                                        If mapnpc(Mapnum).NPC(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_UP) Then
                                                Call NpcMove(Mapnum, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If mapnpc(Mapnum).NPC(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_DOWN) Then
                                                Call NpcMove(Mapnum, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If mapnpc(Mapnum).NPC(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_LEFT) Then
                                                Call NpcMove(Mapnum, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If mapnpc(Mapnum).NPC(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_RIGHT) Then
                                                Call NpcMove(Mapnum, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 1
    
                                        ' Right
                                        If mapnpc(Mapnum).NPC(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_RIGHT) Then
                                                Call NpcMove(Mapnum, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If mapnpc(Mapnum).NPC(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_LEFT) Then
                                                Call NpcMove(Mapnum, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If mapnpc(Mapnum).NPC(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_DOWN) Then
                                                Call NpcMove(Mapnum, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If mapnpc(Mapnum).NPC(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_UP) Then
                                                Call NpcMove(Mapnum, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 2
    
                                        ' Down
                                        If mapnpc(Mapnum).NPC(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_DOWN) Then
                                                Call NpcMove(Mapnum, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If mapnpc(Mapnum).NPC(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_UP) Then
                                                Call NpcMove(Mapnum, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If mapnpc(Mapnum).NPC(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_RIGHT) Then
                                                Call NpcMove(Mapnum, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If mapnpc(Mapnum).NPC(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_LEFT) Then
                                                Call NpcMove(Mapnum, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 3
    
                                        ' Left
                                        If mapnpc(Mapnum).NPC(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_LEFT) Then
                                                Call NpcMove(Mapnum, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If mapnpc(Mapnum).NPC(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_RIGHT) Then
                                                Call NpcMove(Mapnum, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If mapnpc(Mapnum).NPC(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_UP) Then
                                                Call NpcMove(Mapnum, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If mapnpc(Mapnum).NPC(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(Mapnum, x, DIR_DOWN) Then
                                                Call NpcMove(Mapnum, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                End Select
    
                                ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If mapnpc(Mapnum).NPC(x).x - 1 = TargetX And mapnpc(Mapnum).NPC(x).y = TargetY Then
                                        If mapnpc(Mapnum).NPC(x).Dir <> DIR_LEFT Then
                                            Call NpcDir(Mapnum, x, DIR_LEFT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If mapnpc(Mapnum).NPC(x).x + 1 = TargetX And mapnpc(Mapnum).NPC(x).y = TargetY Then
                                        If mapnpc(Mapnum).NPC(x).Dir <> DIR_RIGHT Then
                                            Call NpcDir(Mapnum, x, DIR_RIGHT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If mapnpc(Mapnum).NPC(x).x = TargetX And mapnpc(Mapnum).NPC(x).y - 1 = TargetY Then
                                        If mapnpc(Mapnum).NPC(x).Dir <> DIR_UP Then
                                            Call NpcDir(Mapnum, x, DIR_UP)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If mapnpc(Mapnum).NPC(x).x = TargetX And mapnpc(Mapnum).NPC(x).y + 1 = TargetY Then
                                        If mapnpc(Mapnum).NPC(x).Dir <> DIR_DOWN Then
                                            Call NpcDir(Mapnum, x, DIR_DOWN)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    ' We could not move so Target must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)
    
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
    
                                            If CanNpcMove(Mapnum, x, i) Then
                                                Call NpcMove(Mapnum, x, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
    
                            Else 'Non target movement
                                
                                If Map(Mapnum).NPCSProperties(x).Movement > 0 Then
                                
                                    If Trim$(Movements(Map(Mapnum).NPCSProperties(x).Movement).Name) = vbNullString Then GoTo JumpLine 'This has been done for evit to repeat code
                                    
                                    i = ComputeActualMovement(Map(Mapnum).NPCSProperties(x), Mapnum, x)
                                    If i >= 0 And i <= 3 Then
                                        Call NpcMove(Mapnum, x, i, MOVING_WALKING)
                                    End If
                                    
                                Else
JumpLine:                           i = Int(Rnd * 4)
    
                                    If i = 1 Then
                                        i = Int(Rnd * 4)
                                        Select Case NPC(mapnpc(Mapnum).NPC(x).Num).Behaviour
                                        Case NPC_BEHAVIOUR_BLADE
                                            Dim CanBladeMove As Integer
                                            CanBladeMove = CanBladeNpcMove(Mapnum, x, i)
                                            If CanBladeMove > 0 Then
                                                Call NpcMove(Mapnum, x, i, MOVING_WALKING)
                                                Call ParseAction(CanBladeMove, Map(Mapnum).NPCSProperties(x).Action, 0) 'tile Match
                                            ElseIf CanBladeMove = 0 Then
                                                Call NpcMove(Mapnum, x, i, MOVING_WALKING)
                                            End If
                                                                                
                                        Case Else
                                            If CanNpcMove(Mapnum, x, i) Then
                                                Call NpcMove(Mapnum, x, i, MOVING_WALKING)
                                            End If
                                            
                                        End Select
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Mapnum).NPC(x) > 0 And mapnpc(Mapnum).NPC(x).Num > 0 Then
                    target = mapnpc(Mapnum).NPC(x).target
                    targetType = mapnpc(Mapnum).NPC(x).targetType

                    ' Check if the npc can attack the targeted player player
                    If target > 0 Then
                    
                        If targetType = 1 Then ' player

                            ' Is the target playing and on the same map?
                            If IsPlaying(target) And GetPlayerMap(target) = Mapnum Then
                                TryNpcAttackPlayer x, target
                            Else
                                ' Player left map or game, set target to 0
                                mapnpc(Mapnum).NPC(x).target = 0
                                mapnpc(Mapnum).NPC(x).targetType = 0 ' clear
                            End If
                        ElseIf targetType = 2 Then 'npc
                                'Npc vs Npc
                                Call TryNpcAttackNpc(Mapnum, x, mapnpc(Mapnum).NPC(x).target)
                        End If
                        
                    End If
' Spell Casting
                For i = 1 To MAX_NPC_SPELLS
                    If NPC(NpcNum).Spell(i) > 0 Then
                        If mapnpc(Mapnum).NPC(x).SpellTimer(i) + (Spell(NPC(NpcNum).Spell(i)).CastTime * 1000) < GetTickCount Then
                            
                            Select Case targetType
                            Case 0 'None
                                If ChoosePetSpellingMethod(GetMapPetOwner(Mapnum, x), x, i, NPC(NpcNum).Spell(i)) = False Then
                                End If
                            Case 1 'Player
                                NpcSpellPlayer x, target, i
                            Case 2 'Npc
                                NpcSpellNpc Mapnum, x, target, i
                            End Select
                       
                        End If
                     End If
                Next
                End If
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's Vitals //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's Vitals
                If Not mapnpc(Mapnum).NPC(x).stopRegen Then
                    If mapnpc(Mapnum).NPC(x).Num > 0 And TickCount > GiveNPCHPTimer + 5000 Then 'every 5 seconds
                        For i = 1 To Vitals.Vital_Count - 1
                            If mapnpc(Mapnum).NPC(x).Vital(i) > 0 Or i = Vitals.mp Then
                                mapnpc(Mapnum).NPC(x).Vital(i) = mapnpc(Mapnum).NPC(x).Vital(i) + GetNpcVitalRegen(NpcNum, i, mapnpc(Mapnum).NPC(x).PetData.owner)
    
                                ' Check if they have more then they should and if so just set it to max
                                If mapnpc(Mapnum).NPC(x).Vital(i) > GetNpcMaxVital(NpcNum, i, mapnpc(Mapnum).NPC(x).PetData.owner) Then
                                    mapnpc(Mapnum).NPC(x).Vital(i) = GetNpcMaxVital(NpcNum, i, mapnpc(Mapnum).NPC(x).PetData.owner)
                                End If
                                                       
                            End If
                        Next
                        SendMapNpcVitals Mapnum, x
                    End If
                Else 'Check if we need to set the Regen Variable ON
                    If TickCount > mapnpc(Mapnum).NPC(x).stopRegenTimer + 5000 Then
                        'Turn on
                        mapnpc(Mapnum).NPC(x).stopRegen = False
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If mapnpc(Mapnum).NPC(x).Num = 0 And Map(Mapnum).NPC(x) > 0 Then
                    If TickCount > mapnpc(Mapnum).NPC(x).SpawnWait + (NPC(Map(Mapnum).NPC(x)).SpawnSecs * 1000) Then
                        Call SpawnNpc(x, Mapnum)
                    End If
                End If

            Next

        End If

        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If TickCount > GiveNPCHPTimer + 5000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
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
    
                If GetPlayerVital(i, Vitals.mp) <> GetPlayerMaxVital(i, Vitals.mp) Then
                    Call SetPlayerVital(i, Vitals.mp, GetPlayerVital(i, Vitals.mp) + GetPlayerVitalRegen(i, Vitals.mp))
                    Call SendVital(i, Vitals.mp)
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
        Call TextAdd("Saving all online players...")

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
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Apagado del Servidor en: " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Apagado.", BrightRed)
        Call DestroyServer
    End If

End Sub
