Attribute VB_Name = "modNPC"
Option Explicit

Public Function HasNPCMaxVital(ByVal vital As Vitals, ByVal mapnum As Long, ByVal MapNPCNum As Long, Optional ByVal PetOwner As Long = 0) As Boolean
HasNPCMaxVital = False

If Not (mapnum > 0 And mapnum <= MAX_MAPS And MapNPCNum > 0 And MapNPCNum < MAX_MAP_NPCS) Then Exit Function

If Not (mapnpc(mapnum).NPC(MapNPCNum).Num > 0) Then Exit Function

If mapnpc(mapnum).NPC(MapNPCNum).vital(vital) >= GetNpcMaxVital(mapnpc(mapnum).NPC(MapNPCNum).Num, vital, PetOwner) Then
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

Sub ResetMapNPCSTarget(ByVal mapnum As Long, ByVal MapNPCNum As Long)
Dim i As Long
'Sets all map npc's target to 0 if it's target was the parameter mapnpcnum

'Erase NPC info
For i = 1 To MAX_MAP_NPCS
    If i <> MapNPCNum Then
            If mapnpc(mapnum).NPC(i).target = MapNPCNum And mapnpc(mapnum).NPC(i).targetType = TARGET_TYPE_NPC Then
                    mapnpc(mapnum).NPC(i).target = 0
                    mapnpc(mapnum).NPC(i).target = TARGET_TYPE_NONE
                    If mapnpc(mapnum).NPC(i).IsPet = YES Then
                        TempPlayer(mapnpc(mapnum).NPC(i).PetData.owner).PetHasOwnTarget = 0
                    End If
            End If
    End If
Next

End Sub

Sub SendNpcAttackAnimation(ByVal mapnum As Long, ByVal MapNPCNum As Long)
Dim Buffer As clsBuffer

' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong MapNPCNum
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub


Public Sub KillNpc(ByVal mapnum As Long, ByVal MapNPCNum As Long)
Dim i As Long
        If Not (mapnum > 0 And mapnum <= MAX_MAPS) Or Not (MapNPCNum > 0 And MapNPCNum <= MAX_MAP_NPCS) Then Exit Sub
        
        mapnpc(mapnum).NPC(MapNPCNum).Num = 0
        mapnpc(mapnum).NPC(MapNPCNum).SpawnWait = GetTickCount
        mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With mapnpc(mapnum).NPC(MapNPCNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With mapnpc(mapnum).NPC(MapNPCNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        Dim Buffer As clsBuffer
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong MapNPCNum
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapnum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = MapNPCNum Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
End Sub

Public Function GetNpcLevel(ByVal NPCNum As Long, Optional ByVal PetOwner As Long = 0) As Long
Dim Value As Long


If NPCNum < 1 Or NPCNum > MAX_NPCS Then Exit Function


If PlayerHasPetInMap(PetOwner) Then
    GetNpcLevel = Player(PetOwner).Pet(TempPlayer(PetOwner).ActualPet).Level
Else
    GetNpcLevel = NPC(NPCNum).Level
End If


End Function

Public Function GetNpcStat(ByVal NPCNum As Long, ByVal stat As Stats, Optional ByVal PetOwner As Long = 0, Optional ByVal ConsiderPetBase As Boolean = True) As Long
Dim Value As Long


If NPCNum < 1 Or NPCNum > MAX_NPCS Then Exit Function

Value = NPC(NPCNum).stat(stat)
 
If PlayerHasPetInMap(PetOwner) Then
    Select Case ConsiderPetBase
    Case True
        Value = Value + Player(PetOwner).Pet(TempPlayer(PetOwner).ActualPet).StatsAdd(stat)
    Case False
        Value = Player(PetOwner).Pet(TempPlayer(PetOwner).ActualPet).StatsAdd(stat)
    End Select
End If

GetNpcStat = Value

End Function

Sub CheckNPCSlide(ByVal index As Long, ByVal MapNPCNum As Long, ByVal x As Long, ByVal y As Long, ByVal Dir As Byte)

Dim mapnum As Long
Dim NPCNum As Long
Dim i As Long

mapnum = GetPlayerMap(index)
If Not (mapnum > 0 And mapnum <= MAX_MAPS) Then Exit Sub

If MapNPCNum = 0 Or MapNPCNum > MAX_MAP_NPCS Then Exit Sub

If Not (Dir >= 0 And Dir <= 3) Then Exit Sub

NPCNum = mapnpc(mapnum).NPC(MapNPCNum).Num

If Not (NPCNum > 0 And NPCNum <= MAX_NPCS) Then Exit Sub

If NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_SLIDE Then 'found
    If mapnpc(mapnum).NPC(MapNPCNum).x = x And mapnpc(mapnum).NPC(MapNPCNum).y = y Then
        If CanNpcMove(mapnum, MapNPCNum, Dir) Then
            Call NpcMove(mapnum, MapNPCNum, Dir, 1)
        End If
    End If
End If

End Sub


Public Sub RespawnMapSlideNPC(ByVal mapnum As Long, ByVal MapNPCNum As Long, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0)
Dim NPCNum As Long

If Not (mapnum > 0 And mapnum <= MAX_MAPS) Then Exit Sub

If Not (MapNPCNum > 0 And MapNPCNum <= MAX_MAP_NPCS) Then Exit Sub

NPCNum = mapnpc(mapnum).NPC(MapNPCNum).Num

If Not (NPCNum > 0 And NPCNum < MAX_NPCS) Then Exit Sub

If NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_SLIDE Then Exit Sub

SpawnNpc MapNPCNum, mapnum, x, y

End Sub

Public Sub RespawnMapSlideNPCS(ByVal mapnum As Long)
Dim x As Long
Dim y As Long
Dim MapNPCNum As Long
Dim NPCNum As Long

If Not (mapnum > 0 And mapnum < MAX_MAPS) Then Exit Sub

For x = 1 To Map(mapnum).MaxX
    For y = 1 To Map(mapnum).MaxY
        If Map(mapnum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
            MapNPCNum = Map(mapnum).Tile(x, y).Data1
            If MapNPCNum > 0 And MapNPCNum < MAX_MAP_NPCS Then
                NPCNum = mapnpc(mapnum).NPC(MapNPCNum).Num
                If NPCNum > 0 Then
                    If NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_SLIDE Then
                        RespawnMapSlideNPC mapnum, MapNPCNum, x, y
                    End If
                End If
            End If
        End If
    Next
Next

End Sub

Public Function GetNPCSpellDamage(ByVal Spellnum As Long, ByVal mapnum As Long, ByVal MapNPCNum As Long, Optional ByVal PetOwner As Long = 0) As Long

Dim vital As Long
Dim stat As Long
Dim NPCNum As Long

If Spellnum <= 0 Or Spellnum > MAX_SPELLS Then Exit Function
vital = Spell(Spellnum).vital

NPCNum = mapnpc(mapnum).NPC(MapNPCNum).Num
If NPCNum <= 0 Or NPCNum > MAX_NPCS Then Exit Function

stat = GetNpcStat(NPCNum, Intelligence, PetOwner, False)

Select Case PetOwner

Case Is > 0
    Dim Lvl As Long
    Dim Desv As Double
    Lvl = GetNpcLevel(NPCNum, PetOwner)
    Desv = GetStatDesviation(Lvl, stat)
    If Desv > 2 Then
        Desv = 2
    ElseIf Desv < -2 Then
        Desv = -2
    End If
    GetNPCSpellDamage = vital + ((vital / 4) * Desv)
Case 0
    GetNPCSpellDamage = vital + stat
End Select

End Function
