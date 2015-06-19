Attribute VB_Name = "modWarps"
Sub PlayerWarpBySpell(ByVal index As Long, ByVal spellnum As Long)
    CheckPlayerStateAtWarp index, Spell(spellnum).map
    PlayerWarp index, Spell(spellnum).map, Spell(spellnum).X, Spell(spellnum).Y, False
End Sub

Sub PlayerSpawn(ByVal index As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    CheckPlayerStateAtWarp index, mapnum
    PlayerWarp index, mapnum, X, Y, True
End Sub

Sub PlayerWarpByMapLimits(ByVal index As Long, ByVal dir As Byte)
    CheckPlayerStateAtWarp index, GetPlayerMap(index)
    Dim newmap As Long
    newmap = HasMapWarpByDir(dir, GetPlayerMap(index))
    Dim X As Long, Y As Long
    Call GetMapWarpPosition(dir, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), X, Y)
    PlayerWarp index, newmap, X, Y
End Sub

Sub PlayerWarpByEvent(ByVal index As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    If GetPlayerMap(index) = mapnum Then
        SetPlayerX index, X
        SetPlayerY index, Y
        SendPlayerXYToMap index
    Else
        CheckPlayerStateAtWarp index, mapnum
        PlayerWarp index, mapnum, X, Y
    End If
End Sub



Sub PlayerWarp(ByVal index As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal SendUpdateIfSameMap As Boolean = True)
    Dim shopnum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If
    
    'If IsActionBlocked(index, aTeleport) Then
        'SendMapUpdate index
        'Exit Sub
    'End If
        
    ' Save old map to send erase player data to
    
    
    OldMap = GetPlayerMap(index)
    

    ' Check if you are out of bounds
    If X > map(mapnum).MaxX Then X = map(mapnum).MaxX
    If Y > map(mapnum).MaxY Then Y = map(mapnum).MaxY
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    
    ' if same map then just send their co-ordinates
    If mapnum = OldMap Then
        SetPlayerX index, X
        SetPlayerY index, Y
        SendPlayerXYToMap index
        
        If SendUpdateIfSameMap Then
            TempPlayer(index).GettingMap = YES
            SendMapUpdate index
        End If
        Exit Sub
    End If
    
    Call SendLeaveMap(index, OldMap)
    
    If TempPlayer(index).TempPet.TempPetSlot > 0 Then
        PetDisband index, OldMap, False
    End If
    
    If ArePlayersOnMap(OldMap) = 0 Then
        For i = 1 To MAX_MAP_NPCS
    
            If MapNpc(OldMap).NPC(i).Num > 0 Then
                MapNpc(OldMap).NPC(i).vital(Vitals.HP) = GetNpcMaxVital(OldMap, i, Vitals.HP)
            End If
    
        Next
    End If
    
    DeleteMapPlayer index, OldMap
    AddMapPlayer index, mapnum
    
    ' clear target
    TempPlayer(index).Target = 0
    TempPlayer(index).TargetType = TARGET_TYPE_NONE
    SendTarget index

    

    'If OldMap <> mapnum Then
    
    'End If
    
    Call SendLoad(index)

    Call SetPlayerMap(index, mapnum)
    Call SetPlayerX(index, X)
    Call SetPlayerY(index, Y)
    
    

    ' send player's equipment to new map
    SendMapEquipment index
    
    ' send equipment of all people on new map
    If ArePlayersOnMap(mapnum) > 0 Then
        Dim a As Variant
        For Each a In GetMapPlayerCollection(mapnum)
            If a <> index Then
                SendMapEquipmentTo a, index
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    
    
        ' Regenerate all NPCs' health
        
    'this has to be placed in delete procedeture
    

    'End If

    ' Sets it so we know to process npcs on the map
    TempPlayer(index).GettingMap = YES
    
    'ALATAR
    Call CheckTasks(index, QUEST_TYPE_GOREACH, mapnum)
    '/ALATAR
    
    'Call SendMapNpcsTo(index, mapnum)
    Call SendMapUpdate(index)
End Sub
