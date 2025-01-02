Attribute VB_Name = "modNPC"
Option Explicit

Public Function HasNPCMaxVital(ByVal vital As Vitals, ByVal mapnum As Long, ByVal mapnpcnum As Long) As Boolean
HasNPCMaxVital = False

If Not (mapnum > 0 And mapnum <= MAX_MAPS And mapnpcnum > 0 And mapnpcnum < MAX_MAP_NPCS) Then Exit Function

If Not (MapNpc(mapnum).NPC(mapnpcnum).Num > 0) Then Exit Function

If MapNpc(mapnum).NPC(mapnpcnum).vital(vital) >= GetNpcMaxVital(mapnum, mapnpcnum, vital) Then
    HasNPCMaxVital = True
End If

End Function

Sub RefreshMapNPCS(ByVal index As Long, Optional ByVal OmitNPCS As Boolean = False)

Dim Buffer As clsBuffer

If Not (OmitNPCS) Then
    Call SendMapNpcsTo(index, GetPlayerMap(index))
End If

Set Buffer = New clsBuffer
Buffer.WriteLong SMapDone
SendDataTo index, Buffer.ToArray()
Set Buffer = Nothing

End Sub

Sub CheckNPCSMovement(ByVal Tick As Long)
    Dim CurrentMapIndex As Long
    For CurrentMapIndex = 1 To GetNumberOfMapsWithPlayers
    
        Dim mapnum As Long
        mapnum = GetMapNumByMapReference(CurrentMapIndex)
        If mapnum = 0 Then Exit Sub
        If Not TempMap(mapnum).Exists Then Exit Sub
        
        Dim X As Long
        For X = 1 To GetMapNpcHighIndex(mapnum)
            Call ComputeNPCMovement(mapnum, X, Tick)
        Next
    Next
End Sub

Sub ResetMapNPCSTarget(ByVal mapnum As Long, ByVal mapnpcnum As Long)
Dim i As Long
'Sets all map npc's target to 0 if it's target was the parameter mapnpcnum

'Erase NPC info
For i = 1 To MAX_MAP_NPCS
    If i <> mapnpcnum Then
            If MapNpc(mapnum).NPC(i).Target = mapnpcnum And MapNpc(mapnum).NPC(i).TargetType = TARGET_TYPE_NPC Then
                    MapNpc(mapnum).NPC(i).Target = 0
                    MapNpc(mapnum).NPC(i).Target = TARGET_TYPE_NONE
                    If IsMapNPCaPet(mapnum, i) Then
                        ResetPetTarget GetMapPetOwner(mapnum, i)
                    End If
            End If
    End If
Next

End Sub

Sub SendNpcAttackAnimation(ByVal mapnum As Long, ByVal mapnpcnum As Long)
Dim Buffer As clsBuffer

' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong mapnpcnum
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub KillNpc(ByVal mapnum As Long, ByVal mapnpcnum As Long)
Dim i As Long
        If Not (mapnum > 0 And mapnum <= MAX_MAPS) Or Not (mapnpcnum > 0 And mapnpcnum <= MAX_MAP_NPCS) Then Exit Sub
        
        Call SendAnimation(mapnum, 145, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y)
        
        If IsTempNPC(mapnum, mapnpcnum) Then
            If GetNPCNum(mapnum, mapnpcnum) = NPC_SKULLTULA Then
                Call RespawnRandomNPC(GetNPCNum(mapnum, mapnpcnum), mapnum)
            End If
            Call SendClearMapNpcToMap(mapnum, mapnpcnum)
            Call ClearSingleMapNpc(mapnpcnum, mapnum)
            'Call SendMapNpcToMap(mapnum, mapnpcnum)
        ElseIf IsMapNPCaPet(mapnum, mapnpcnum) Then
            PlayerMsg GetMapPetOwner(mapnum, mapnpcnum), "Your pet has died!", White
            Call PetDisband(GetMapPetOwner(mapnum, mapnpcnum), mapnum, True)
        Else
            Call SendClearMapNpcToMap(mapnum, mapnpcnum)
            Call ClearSingleMapNpc(mapnpcnum, mapnum)
            If map(mapnum).NPC(mapnpcnum) > 0 Then
                Call AddWaitingNPC(mapnum, mapnpcnum, NPC(map(mapnum).NPC(mapnpcnum)).SpawnSecs)
            End If
            'Call SendMapNpcToMap(mapnum, mapnpcnum)

            'mapnpc(mapnum).NPC(mapnpcnum).Num = 0
            'mapnpc(mapnum).NPC(MapNPCNum).SpawnWait = GetRealTickCount
            'mapnpc(mapnum).NPC(mapnpcnum).vital(Vitals.HP) = 0
            
            'Checks if NPC was a pet
            'If IsMapNPCaPet(mapnum, mapnpcnum) Then
                'Call PetDisband(GetMapPetOwner(mapnum, mapnpcnum), mapnum) 'The pet was killed
            'End If
            
            'Restart NPC's Target if it was mapnpcnum
            
            
            ' clear DoTs and HoTs
            'For i = 1 To MAX_DOTS
                'With mapnpc(mapnum).NPC(mapnpcnum).DoT(i)
                    '.Spell = 0
                    '.Timer = 0
                    '.caster = 0
                    '.StartTime = 0
                    '.Used = False
                'End With
                
                'With mapnpc(mapnum).NPC(mapnpcnum).HoT(i)
                    '.Spell = 0
                    '.Timer = 0
                    '.caster = 0
                    '.StartTime = 0
                    '.Used = False
                'End With
            'Next
            
            'Dim buffer As clsBuffer
            ' send death to the map
            'Set buffer = New clsBuffer
            'buffer.WriteLong SNpcDead
            'buffer.WriteLong GetMapNpcNumForClient(mapnum, mapnpcnum)
            'SendDataToMap mapnum, buffer.ToArray()
            'Set buffer = Nothing
        End If
            
        ResetMapNPCSTarget mapnum, mapnpcnum
        
        'Loop through entire map and purge NPC from targets
        ResetMapPlayerNPCTarget mapnum, mapnpcnum
        
        
        
        
        
End Sub

Sub ResetMapPlayerNPCTarget(ByVal mapnum As Long, ByVal mapnpcnum As Long)
    Dim a As Variant
    For Each a In GetMapPlayerCollection(mapnum)
        If TempPlayer(a).TargetType = TARGET_TYPE_NPC Then
            If TempPlayer(a).Target = mapnpcnum Then
                TempPlayer(a).Target = 0
                TempPlayer(a).TargetType = TARGET_TYPE_NONE
                SendTarget a
            End If
        End If
    Next
End Sub

Public Function GetNpcLevel(ByVal npcnum As Long, Optional ByVal PetOwner As Long = 0) As Long
Dim Value As Long


If npcnum < 1 Or npcnum > MAX_NPCS Then Exit Function


If PlayerHasPetInMap(PetOwner) > 0 Then
    GetNpcLevel = player(PetOwner).Pet(TempPlayer(PetOwner).TempPet.ActualPet).level
Else
    GetNpcLevel = NPC(npcnum).level
End If


End Function

Public Function GetNpcStat(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal stat As Stats, Optional ByVal ConsiderPetBase As Boolean = True) As Long
Dim Value As Long
Dim PetOwner As Long

Dim npcnum As Long
npcnum = GetNPCNum(mapnum, mapnpcnum)
If npcnum = 0 Then Exit Function

Value = 0
If ConsiderPetBase Then
    Value = NPC(npcnum).stat(stat)
End If

PetOwner = GetMapPetOwner(mapnum, mapnpcnum)
If PetOwner > 0 Then
    Value = Value + GetPlayerPetStat(PetOwner, stat)
End If

GetNpcStat = Value

End Function

Sub CheckNPCSlide(ByVal index As Long, ByVal mapnpcnum As Long, ByVal X As Long, ByVal Y As Long, ByVal dir As Byte)

Dim mapnum As Long
Dim npcnum As Long
Dim i As Long

mapnum = GetPlayerMap(index)
If Not (mapnum > 0 And mapnum <= MAX_MAPS) Then Exit Sub

If mapnpcnum = 0 Or mapnpcnum > MAX_MAP_NPCS Then Exit Sub

npcnum = MapNpc(mapnum).NPC(mapnpcnum).Num

If Not (npcnum > 0 And npcnum <= MAX_NPCS) Then Exit Sub

If NPC(npcnum).Behaviour = NPC_BEHAVIOUR_SLIDE Then 'found
    If MapNpc(mapnum).NPC(mapnpcnum).X = X And MapNpc(mapnum).NPC(mapnpcnum).Y = Y Then
        If CanNpcMove(mapnum, mapnpcnum, dir, True) Then
            Call NpcMove(mapnum, mapnpcnum, dir, 1)
            Call SendSoundToMap(mapnum, X, Y, SoundEntity.seSlide, 1)
            
            
        End If
    End If
End If

End Sub


Public Sub RespawnMapSlideNPC(ByVal mapnum As Long, ByVal mapnpcnum As Long, Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0)
Dim npcnum As Long

If Not (mapnum > 0 And mapnum <= MAX_MAPS) Then Exit Sub

If Not (mapnpcnum > 0 And mapnpcnum <= MAX_MAP_NPCS) Then Exit Sub

npcnum = MapNpc(mapnum).NPC(mapnpcnum).Num

If Not (npcnum > 0 And npcnum < MAX_NPCS) Then Exit Sub

If NPC(npcnum).Behaviour <> NPC_BEHAVIOUR_SLIDE Then Exit Sub

SpawnNpc mapnpcnum, mapnum, X, Y

End Sub

Public Sub RespawnMapSlideNPCS(ByVal mapnum As Long)
Dim X As Long
Dim Y As Long
Dim mapnpcnum As Long
Dim npcnum As Long

If Not (mapnum > 0 And mapnum < MAX_MAPS) Then Exit Sub

For X = 1 To map(mapnum).MaxX
    For Y = 1 To map(mapnum).MaxY
        If map(mapnum).Tile(X, Y).Type = TILE_TYPE_NPCSPAWN Then
            mapnpcnum = map(mapnum).Tile(X, Y).Data1
            If mapnpcnum > 0 And mapnpcnum < MAX_MAP_NPCS Then
                npcnum = MapNpc(mapnum).NPC(mapnpcnum).Num
                If npcnum > 0 Then
                    If NPC(npcnum).Behaviour = NPC_BEHAVIOUR_SLIDE Then
                        RespawnMapSlideNPC mapnum, mapnpcnum, X, Y
                    End If
                End If
            End If
        End If
    Next
Next

End Sub




Public Function CanNPCBehaviourAttack(ByVal npcnum As Long) As Boolean
CanNPCBehaviourAttack = True
Dim Beh As Byte
Beh = NPC(npcnum).Behaviour
If (Beh = NPC_BEHAVIOUR_FRIENDLY Or Beh = NPC_BEHAVIOUR_SHOPKEEPER Or Beh = NPC_BEHAVIOUR_BLADE Or Beh = NPC_BEHAVIOUR_SLIDE) Then
    CanNPCBehaviourAttack = False
End If

End Function

Public Function IsTileWalkableByBladeNPC(ByVal TileType As Byte) As Boolean

IsTileWalkableByBladeNPC = False

If TileType = TILE_TYPE_WALKABLE Or TileType = TILE_TYPE_ITEM Or TileType = TILE_TYPE_NPCSPAWN Or TileType = TILE_TYPE_RESOURCE Or TileType = TILE_TYPE_ICE Then
    IsTileWalkableByBladeNPC = True
End If

End Function

Public Function IsTileWalkableByPlayer(ByVal TileType As Byte) As Boolean
    If TileType = TILE_TYPE_WALKABLE Or TileType = TILE_TYPE_ITEM Then
        IsTileWalkableByPlayer = True
    End If
End Function


Public Function SpawnTempNPC(ByVal npcnum As Long, ByVal mapnum As Long, Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1) As Integer

    If mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Function
    
    Dim i As Integer
    SpawnTempNPC = 0
    Dim slot As Integer
    
    For i = 1 To MAX_MAP_NPCS
        'If Map(PlayerMap).Npc(i) = 0 Then
         If map(mapnum).NPC(i) = 0 And MapNpc(mapnum).NPC(i).Num = 0 Then
            slot = i
            Exit For
         End If
    Next
    
    If slot = 0 Then
        Exit Function
    End If

    'map(mapnum).NPC(slot) = NPCNum
    MapNpc(mapnum).NPC(slot).Num = npcnum
    MapNpc(mapnum).NPC(slot).IsTempNPC = True
    
    If X = -1 Then X = RAND(0, map(mapnum).MaxX)
    If Y = -1 Then Y = RAND(0, map(mapnum).MaxY)
   
    Call SpawnNpc(slot, mapnum, X, Y, npcnum)
    
    
    SpawnTempNPC = slot
    
End Function




Public Function IsTempNPC(ByVal mapnum As Long, ByVal mapnpcnum As Long) As Boolean
If mapnum < 1 Or mapnum > MAX_MAPS Or mapnpcnum < 1 Or mapnpcnum > MAX_MAP_NPCS Then Exit Function
IsTempNPC = MapNpc(mapnum).NPC(mapnpcnum).IsTempNPC
End Function


Public Function GetNPCAttackTimer(ByVal mapnum As Long, ByVal mapnpcnum As Long) As Long
    Dim i As Long
    i = GetMapPetOwner(mapnum, mapnpcnum)
    If i > 0 Then
        Dim MinFactor As Single
        Dim MaxFactor As Single
        MinFactor = 1
        MaxFactor = 2 / 5
        Dim factor As Double
        factor = (MaxFactor - MinFactor) / MAX_PET_STAT * GetNpcStat(mapnum, mapnpcnum, Agility) + MinFactor
        If factor < MaxFactor Then
            factor = MaxFactor
        End If
        GetNPCAttackTimer = 1000 * factor
    Else
        GetNPCAttackTimer = 1000
    End If
    
End Function

Function GetNPCTarget(ByVal mapnum As Long, ByVal mapnpcnum As Long) As Long
    If mapnum = 0 Or mapnpcnum = 0 Then Exit Function
    GetNPCTarget = MapNpc(mapnum).NPC(mapnpcnum).Target
End Function

Public Function GetNPCMoveTimer(ByVal mapnum As Long, ByVal mapnpcnum As Long) As Long
    Dim i As Long
    i = GetMapPetOwner(mapnum, mapnpcnum)
    If i > 0 And GetNPCTarget(mapnum, mapnpcnum) <> 0 Then
        Dim MinFactor As Single
        Dim MaxFactor As Single
        MinFactor = 2 / 5
        MaxFactor = 2 / 20
        Dim factor As Double
        factor = (MaxFactor - MinFactor) / MAX_PET_STAT * (GetNpcStat(mapnum, mapnpcnum, Agility) * 1.3) + MinFactor
        If factor < MaxFactor Then
            factor = MaxFactor
        End If
        GetNPCMoveTimer = 1000 * factor
    Else
        GetNPCMoveTimer = GetNPCSpeed(GetNPCNum(mapnum, mapnpcnum))
    End If
End Function

Function GetNPCSpeed(ByVal npcnum As Long) As Long
    If npcnum = 0 Then Exit Function
    GetNPCSpeed = NPC(npcnum).Speed
End Function

Sub ComputeNPCMovement(ByVal mapnum As Long, ByVal X As Long, ByVal Tick As Long)
    If mapnum = 0 Or X = 0 Then Exit Sub
    If GetNPCNum(mapnum, X) = 0 Then Exit Sub
    If MapNpc(mapnum).NPC(X).MoveTimer > Tick Then Exit Sub
    If MapNpc(mapnum).NPC(X).StunDuration > 0 Then Exit Sub
    
    If GetNPCTarget(mapnum, X) > 0 Then
        Call ComputeNPCTargetMovement(mapnum, X)
    Else
        Call ComputeNPCNonTargetMovement(mapnum, X)
    End If
    
    CheckNPCVision mapnum, X
End Sub


Sub ComputeNPCTargetMovement(ByVal mapnum As Long, ByVal X As Long)
' X = mapnpcnum
Dim Target As Long, TargetType As Byte, DidWalk As Boolean, TargetX As Long, TargetY As Long, target_verify As Boolean
Dim npcnum As Long
Target = MapNpc(mapnum).NPC(X).Target
TargetType = MapNpc(mapnum).NPC(X).TargetType
npcnum = MapNpc(mapnum).NPC(X).Num

If npcnum < 1 Or npcnum > MAX_NPCS Then Exit Sub

' Check to see if its time for the npc to walk
If NPC(npcnum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Or NPC(npcnum).Behaviour = NPC_BEHAVIOUR_SLIDE Then Exit Sub


If TargetType = 1 Then ' player

    ' Check to see if we are following a player or not
    If Target > 0 Then

        ' Check if the player is even playing, if so follow'm
        If IsPlaying(Target) And GetPlayerMap(Target) = mapnum Then
            DidWalk = False
            target_verify = True
            TargetY = GetPlayerY(Target)
            TargetX = GetPlayerX(Target)
            'Check if Player Pet has to defend his owner
            If TempPlayer(Target).TempPet.TempPetSlot > 0 And TempPlayer(Target).TempPet.TempPetSlot <> X Then
                If TempPlayer(Target).TempPet.PetState = Passive Then
                    ResetPetTarget (Target)
                ElseIf TempPlayer(Target).TempPet.PetHasOwnTarget = NO Then
                    If IsinRange(5, TargetX, TargetY, MapNpc(mapnum).NPC(X).X, MapNpc(mapnum).NPC(X).Y) Then
                        'Pet has not a target, let's catch this npc
                        If MapNpc(mapnum).NPC(X).PetData.Owner > 0 Then
                        'NPC is a pet!
                            If RAND(1, 2) <> 1 Then
                                TempPlayer(Target).TempPet.PetHasOwnTarget = X
                                MapNpc(mapnum).NPC(TempPlayer(Target).TempPet.TempPetSlot).TargetType = TARGET_TYPE_NPC
                                MapNpc(mapnum).NPC(TempPlayer(Target).TempPet.TempPetSlot).Target = X
                            Else
                                TempPlayer(Target).TempPet.PetHasOwnTarget = MapNpc(mapnum).NPC(X).PetData.Owner
                                MapNpc(mapnum).NPC(TempPlayer(Target).TempPet.TempPetSlot).TargetType = TARGET_TYPE_PLAYER
                                MapNpc(mapnum).NPC(TempPlayer(Target).TempPet.TempPetSlot).Target = MapNpc(mapnum).NPC(X).PetData.Owner
                            End If
                        Else
                        'NPC is not a pet
                            TempPlayer(Target).TempPet.PetHasOwnTarget = X
                            MapNpc(mapnum).NPC(TempPlayer(Target).TempPet.TempPetSlot).TargetType = TARGET_TYPE_NPC
                            MapNpc(mapnum).NPC(TempPlayer(Target).TempPet.TempPetSlot).Target = X
                        End If
                    End If
                End If
            End If
        Else
            MapNpc(mapnum).NPC(X).TargetType = 0 ' clear
            MapNpc(mapnum).NPC(X).Target = 0
            'PetFollowOwner MapNpc(mapnum).NPC(X).PetData.Owner
            'TempPlayer(Target).TempPet.PetHasOwnTarget = 0
        End If
    End If

ElseIf TargetType = 2 Then 'npc
    
    If Target > 0 Then
        
        If MapNpc(mapnum).NPC(Target).Num > 0 Then
            DidWalk = False
            target_verify = True
            TargetY = MapNpc(mapnum).NPC(Target).Y
            TargetX = MapNpc(mapnum).NPC(Target).X
        Else
            MapNpc(mapnum).NPC(X).TargetType = 0 ' clear
            MapNpc(mapnum).NPC(X).Target = 0
            PetFollowOwner MapNpc(mapnum).NPC(X).PetData.Owner
        End If
    End If
End If

If target_verify Then

    Dim NPCX As Long, NPCY As Long
    Dim NextX As Long, NextY As Long
    
    NPCX = MapNpc(mapnum).NPC(X).X
    NPCY = MapNpc(mapnum).NPC(X).Y
    
    Dim dir As Byte
    If Abs(NPCX - TargetX) + Abs(NPCY - TargetY) <= 1 Then
        dir = GetDirByCollindantPos(NPCX, NPCY, TargetX, TargetY)
        If MapNpc(mapnum).NPC(X).dir <> dir Then
            Call NpcDir(mapnum, X, dir)
        End If
        Exit Sub
    End If
    
    Dim PriorityDirs() As Byte
    PriorityDirs = GetDirPriority(NPCX, NPCY, TargetX, TargetY)
    
    
    For dir = 0 To 2
        If PriorityDirs(dir) <> GetOppositeDir(MapNpc(mapnum).NPC(X).LastDir) Then
            If CanNpcMove(mapnum, X, PriorityDirs(dir)) Then
                Call NpcMove(mapnum, X, PriorityDirs(dir), MOVING_WALKING)
                MapNpc(mapnum).NPC(X).LastDir = PriorityDirs(dir)
                DidWalk = True
                Exit For
            End If
        End If
    Next
    
    
    
    
    
   

   

    

        ' We could not move so Target must be behind something, walk randomly.
    If Not DidWalk Then
        Dim i As Long
        i = Int(Rnd * 2)

        If i = 1 Then
            i = Int(Rnd * 4)

            If CanNpcMove(mapnum, X, i) Then
                Call NpcMove(mapnum, X, i, MOVING_WALKING)
                DidWalk = True
                MapNpc(mapnum).NPC(X).LastDir = i
            End If
        End If
    End If
End If

If DidWalk Then
    MapNpc(mapnum).NPC(X).MoveTimer = GetRealTickCount + GetNPCMoveTimer(mapnum, X)
End If
End Sub


Sub ComputeNPCNonTargetMovement(ByVal mapnum As Long, ByVal X As Long)

Dim i As Byte
Dim npcnum As Long
npcnum = MapNpc(mapnum).NPC(X).Num

If npcnum < 1 Or npcnum > MAX_NPCS Then Exit Sub
' Check to see if its time for the npc to walk
If NPC(npcnum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Or NPC(npcnum).Behaviour = NPC_BEHAVIOUR_SLIDE Then Exit Sub

If map(mapnum).NPCSProperties(X).Movement > 0 Then
    If Trim$(Movements(map(mapnum).NPCSProperties(X).Movement).Name) <> vbNullString Then
        i = GetNextMovementDir(map(mapnum).NPCSProperties(X), mapnum, X)
        Call NpcMove(mapnum, X, i, MOVING_WALKING)
    End If
Else
    'i = Int(Rnd * 5)

    If Rnd <= 0.2 Then
        i = Int(Rnd * 4)
        Select Case NPC(MapNpc(mapnum).NPC(X).Num).Behaviour
        Case NPC_BEHAVIOUR_BLADE
            Dim CanBladeMove As Integer
            CanBladeMove = CanBladeNpcMove(mapnum, X, i)
            
            If CanBladeMove > 0 Then
                Call NpcMove(mapnum, X, i, MOVING_WALKING)
                Dim ActionNum As Byte
                ActionNum = map(mapnum).NPCSProperties(X).Action
                If ActionNum > 0 Then
                    Dim AMoment As Byte
                    AMoment = Actions(ActionNum).Moment
                    If AMoment = TileMatch Or AMoment = InFrontRange Then
                        Call ParseAction(CanBladeMove, map(mapnum).NPCSProperties(X).Action)
                    End If
                End If
                    
                    
            ElseIf CanBladeMove = 0 Then
                Call NpcMove(mapnum, X, i, MOVING_WALKING)
            End If
                                                       
        Case Else
            If CanNpcMove(mapnum, X, i) Then
                Call NpcMove(mapnum, X, i, MOVING_WALKING)
            End If
            
        End Select
    End If
End If

MapNpc(mapnum).NPC(X).MoveTimer = GetRealTickCount + GetNPCMoveTimer(mapnum, X)

End Sub

Function GetNPCNum(ByVal mapnum As Long, ByVal mapnpcnum As Long) As Long
    If mapnum = 0 Or mapnpcnum = 0 Then Exit Function
    GetNPCNum = MapNpc(mapnum).NPC(mapnpcnum).Num
End Function


Function GetNPCBaseStat(ByVal npcnum As Long, ByVal stat As Stats) As Long
    If npcnum < 1 Or npcnum > MAX_NPCS Then Exit Function
    GetNPCBaseStat = NPC(npcnum).stat(stat)
End Function

Function GetNPCVital(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal vital As Vitals) As Long
    If mapnum = 0 Or mapnpcnum = 0 Then Exit Function
    GetNPCVital = MapNpc(mapnum).NPC(mapnpcnum).vital(vital)
End Function

Function GetNPCSFightChance(ByVal mapnum As Long, ByVal mapnpcnum1 As Long, ByVal mapnpcnum2 As Long) As Single
'1: equal chance, > 1: low chance, < 1: good chance
Dim damage1to2 As Long
Dim damage2to1 As Long

damage1to2 = GetNpcDamage(mapnum, mapnpcnum1)
damage1to2 = GetNpcDefense(mapnum, mapnpcnum2, damage1to2)

damage2to1 = GetNpcDamage(mapnum, mapnpcnum2)
damage2to1 = GetNpcDefense(mapnum, mapnpcnum1, damage2to1)

Dim SecToWin1 As Single, SecToWin2 As Single

SecToWin1 = GetNPCVital(mapnum, mapnpcnum2, HP) / damage1to2 * GetNPCAttackTimer(mapnum, mapnpcnum1)
SecToWin2 = GetNPCVital(mapnum, mapnpcnum1, HP) / damage2to1 * GetNPCAttackTimer(mapnum, mapnpcnum2)

GetNPCSFightChance = SecToWin1 / SecToWin2

End Function

Function GetNPCSSpellChance(ByVal mapnum As Long, ByVal mapnpcnum1 As Long, ByVal mapnpcnum2 As Long) As Single


End Function


Function NPCExists(ByVal npcnum As Long) As Boolean
If LenB(Trim$(NPC(npcnum).Name)) > 0 And Asc(NPC(npcnum).Name) <> 0 Then
    NPCExists = True
End If
End Function

Function GetMapNPCAction(ByVal mapnum As Long, ByVal mapnpcnum As Long) As Long
    GetMapNPCAction = map(mapnum).NPCSProperties(mapnpcnum).Action
End Function


Function GetNPCName(ByVal npcnum As Long) As String
    GetNPCName = Trim$(NPC(npcnum).Name)
End Function

Function CanSlideThroughTile(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    If OutOfBoundries(X, Y, mapnum) Then Exit Function
    
    Select Case GetTileType(mapnum, X, Y)
    Case TILE_TYPE_WALKABLE, TILE_TYPE_ITEM, TILE_TYPE_NPCSPAWN, TILE_TYPE_ICE, TILE_TYPE_NPCAVOID, TILE_TYPE_SCRIPT
        CanSlideThroughTile = True
        
    Case TILE_TYPE_DOOR, TILE_TYPE_KEY
        Dim TempDoorNum As Long
        TempDoorNum = GetTempDoorNumberByTile(mapnum, X, Y)
        CanSlideThroughTile = True
        If TempDoorNum > 0 Then
            If Not IsTempDoorWalkable(mapnum, TempDoorNum) Then
                CanSlideThroughTile = False
                Exit Function
            End If
            
            If IsDoorOpened(mapnum, TempDoorNum) Then
                CanSlideThroughTile = True
            End If
            
            If GetDoorType(TempTile(mapnum).Door(TempDoorNum).DoorNum) = DOOR_TYPE_WEIGHTSWITCH Then
                Call CheckWeightSwitch(mapnum, TempDoorNum)
            End If
        End If
        
    Case TILE_TYPE_RESOURCE
        If isWalkableResource(mapnum, X, Y) Then
            CanSlideThroughTile = True
        Else
            CanSlideThroughTile = False
        End If

    Case Else
        CanSlideThroughTile = False
    End Select

End Function

Function CanNpcMoveThroughTile(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    If OutOfBoundries(X, Y, mapnum) Then Exit Function
    
    Select Case GetTileType(mapnum, X, Y)
    Case TILE_TYPE_WALKABLE, TILE_TYPE_ITEM, TILE_TYPE_NPCSPAWN, TILE_TYPE_ICE
        CanNpcMoveThroughTile = True
    Case TILE_TYPE_DOOR, TILE_TYPE_KEY
        Dim TempDoorNum As Long
        TempDoorNum = GetTempDoorNumberByTile(mapnum, X, Y)
        CanNpcMoveThroughTile = True
        If TempDoorNum > 0 Then
            If Not IsTempDoorWalkable(mapnum, TempDoorNum) Then
                CanNpcMoveThroughTile = False
                Exit Function
            End If
            
            If IsDoorOpened(mapnum, TempDoorNum) Then
                CanNpcMoveThroughTile = True
            End If
            
            If GetDoorType(TempTile(mapnum).Door(TempDoorNum).DoorNum) = DOOR_TYPE_WEIGHTSWITCH Then
                Call CheckWeightSwitch(mapnum, TempDoorNum)
            End If
        End If
    Case TILE_TYPE_RESOURCE
        If isWalkableResource(mapnum, X, Y) Then
            CanNpcMoveThroughTile = True
        Else
            CanNpcMoveThroughTile = False
        End If

    Case Else
        CanNpcMoveThroughTile = False
    End Select
End Function

Function CanNpcMoveToPos(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    CanNpcMoveToPos = True

    If FindPlayerByPos(mapnum, X, Y) > 0 Then
        CanNpcMoveToPos = False
        Exit Function
    End If
    
    ' Check to make sure that there is not another npc in the way
    
    If GetMapRefNPCNumByTile(GetMapRef(mapnum), X, Y) > 0 Then
        CanNpcMoveToPos = False
        Exit Function
    End If

    ' Directional blocking
    Dim prevx As Long, prevy As Long
    prevx = MapNpc(mapnum).NPC(mapnpcnum).X
    prevy = MapNpc(mapnum).NPC(mapnpcnum).Y
    
    If isDirBlocked(map(mapnum).Tile(prevx, prevy).DirBlock, GetDirByPos(prevx, prevy, X, Y) + 1) Then
        CanNpcMoveToPos = False
        Exit Function
    End If
    
    If Not CanNpcMoveThroughTile(mapnum, X, Y) Then
        CanNpcMoveToPos = False
        Exit Function
    End If

End Function

Function CanNpcMove(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal dir As Byte, Optional blIsSliding As Boolean = False) As Boolean
    Dim i As Long
    Dim X As Long
    Dim Y As Long

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
        CanNpcMove = False
        Exit Function
    End If
    
    CanNpcMove = True

    If FindPlayerByPos(mapnum, X, Y) > 0 Then
        CanNpcMove = False
        Exit Function
    End If
    'For i = 1 To Player_HighIndex
        'If IsPlaying(i) Then
            'If (GetPlayerMap(i) = mapnum) Then
                'If (GetPlayerX(i) = x) And (GetPlayerY(i) = y) Then
                    'CanNpcMove = False
                    'Exit Function
                'End If
            'End If
        'End If
    'Next
    
    ' Check to make sure that there is not another npc in the way
    
    If GetMapRefNPCNumByTile(GetMapRef(mapnum), X, Y) > 0 Then
        CanNpcMove = False
        Exit Function
    End If
    
    'For i = 1 To TempMap(mapnum).npc_highindex
        'If (i <> mapnpcnum) And (mapnpc(mapnum).NPC(i).Num > 0) And (mapnpc(mapnum).NPC(i).x = x) And (mapnpc(mapnum).NPC(i).y = y) Then
            'CanNpcMove = False
            'Exit Function
        'End If
    'Next
       
    ' Directional blocking
    If isDirBlocked(map(mapnum).Tile(prevx, prevy).DirBlock, dir + 1) Then
        CanNpcMove = False
        Exit Function
    End If
    
    If Not CanNpcMoveThroughTile(mapnum, X, Y) Then
    'sliding?
        If blIsSliding = True Then
            If CanSlideThroughTile(mapnum, X, Y) Then
                CanNpcMove = True
                Exit Function
            End If
        End If
        CanNpcMove = False
        Exit Function
    End If

End Function

Sub NpcMove(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal dir As Long, ByVal Movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Then
        Exit Sub
    End If
    
    If dir < 0 Or dir > 3 Then Exit Sub

    If Not ComputeNPCSingleMovement(mapnum, mapnpcnum, dir) Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcMove
    Buffer.WriteLong mapnpcnum
    Buffer.WriteByte dir
    Buffer.WriteLong Movement
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Sub NpcDir(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(mapnum).NPC(mapnpcnum).dir = dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong mapnpcnum
    Buffer.WriteLong dir
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SetAllWorldNpcs(ByVal npcnum As Long)
    Dim X As Long, i As Long
    For X = 1 To MAX_MAPS
        For i = 1 To GetMapNpcHighIndex(X)
            If GetNPCNum(X, i) > 0 Then
                MapNpc(X).NPC(i).Num = npcnum
            End If
        Next
    Next
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            SendMapNpcsTo i, GetPlayerMap(i)
        End If
    
    Next
End Sub
