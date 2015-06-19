Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ## Basic Calculations ##
' ################################

Function GetPlayerMaxVital(ByVal index As Long, ByVal vital As Vitals) As Long
If index > MAX_PLAYERS Then Exit Function
Select Case vital
Case HP
GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, endurance) / 2)) * 15 + 150
Case mp
GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 30 + 85
End Select
End Function

Function GetPlayerVitalRegen(ByVal index As Long, ByVal vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case vital
        Case HP
            'i = (GetPlayerStat(index, Stats.willpower) * 0.8) + 60
            i = GetPlayerMaxVital(index, HP) * (0.05 + (GetVitalRegenPercent(GetStatDesviation(GetPlayerLevel(index), GetPlayerStat(index, Stats.willpower)))))
        Case mp
            'i = (GetPlayerStat(index, Stats.willpower) / 4) + 60
            i = GetPlayerMaxVital(index, HP) * (0.03 + (GetVitalRegenPercent(GetStatDesviation(GetPlayerLevel(index), GetPlayerStat(index, Stats.willpower)))))
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

Function GetPlayerDamage(ByVal index As Long) As Long
    Dim weaponNum As Long
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(index, Weapon)
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(index, Strength) * Item(weaponNum).Data2 + (GetPlayerLevel(index) / 5)
    Else
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(index, Strength) + (GetPlayerLevel(index) / 5)
    End If

End Function
Function GetPlayerDef(ByVal index As Long) As Long
    Dim DefNum As Long
    Dim Def As Long
    
    GetPlayerDef = 0
    Def = 0
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(index, Armor)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Helmet) > 0 Then
        DefNum = GetPlayerEquipment(index, Helmet)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        DefNum = GetPlayerEquipment(index, Shield)
        Def = Def + Item(DefNum).Data2
    End If
    
   If Not GetPlayerEquipment(index, Armor) > 0 And Not GetPlayerEquipment(index, Helmet) > 0 And Not GetPlayerEquipment(index, Shield) > 0 Then
        GetPlayerDef = GetPlayerStat(index, endurance) + (GetPlayerLevel(index))
    Else
        GetPlayerDef = GetPlayerStat(index, endurance) * (Def / 40) + (GetPlayerLevel(index))
    End If
    

End Function

Function GetNpcMaxVital(ByVal NPCNum As Long, ByVal vital As Vitals, Optional ByVal PetOwner As Long = 0) As Long
    Dim x As Long

    ' Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If
    
    'Prevent Pet System
    If Not (PlayerHasPetInMap(PetOwner)) Then
        PetOwner = 0
    End If
        
    Select Case PetOwner
    
    Case Is > 0
        Select Case vital
            Case HP
                GetNpcMaxVital = NPC(NPCNum).HP + ((Player(PetOwner).Pet(TempPlayer(PetOwner).ActualPet).Level / 2) + (GetNpcStat(NPCNum, endurance, PetOwner) / 2) * 10)
            Case mp
                GetNpcMaxVital = 30 + ((Player(PetOwner).Pet(TempPlayer(PetOwner).ActualPet).Level / 2) + (GetNpcStat(NPCNum, Intelligence, PetOwner) / 2)) * 10
            End Select
    Case Else
            Select Case vital
            Case HP
                GetNpcMaxVital = NPC(NPCNum).HP
            Case mp
                GetNpcMaxVital = 30 + (NPC(NPCNum).stat(Intelligence) * 10) + 2
            End Select
    End Select

End Function

Function GetNpcVitalRegen(ByVal NPCNum As Long, ByVal vital As Vitals, Optional ByVal PetOwner As Long = 0) As Long
    Dim i As Long

    'Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If
    
    'Prevent Pet System
    If Not (PlayerHasPetInMap(PetOwner)) Then
        PetOwner = 0
    End If
    
    Select Case PetOwner
    
    Case Is > 0

        Select Case vital
            Case HP
                ' i = (GetNpcStat(NPCNum, willpower, PetOwner) * 0.8) + 6
                i = GetNpcMaxVital(NPCNum, HP, PetOwner) * (0.05 + (GetVitalRegenPercent(GetStatDesviation(GetNpcLevel(NPCNum, PetOwner), GetNpcStat(NPCNum, willpower, PetOwner, False)))))
            Case mp
                i = GetNpcMaxVital(NPCNum, mp, PetOwner) * (0.03 + (GetVitalRegenPercent(GetStatDesviation(GetNpcLevel(NPCNum, PetOwner), GetNpcStat(NPCNum, willpower, PetOwner, False)))))
        End Select
    
    Case 0
    
        Select Case vital
            Case HP
                i = (GetNpcStat(NPCNum, willpower) * 0.8) + 6
            Case mp
                i = (GetNpcStat(NPCNum, willpower) / 4) + 12.5
        End Select
    
    End Select
    
    GetNpcVitalRegen = i

End Function

Function GetNpcDamage(ByVal NPCNum As Long, Optional ByVal PetOwner As Long = 0) As Long
    GetNpcDamage = 0.085 * 5 * GetNpcStat(NPCNum, Strength, PetOwner) * NPC(NPCNum).Damage + (GetNpcLevel(NPCNum, PetOwner) / 5)
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCrit = False

    rate = GetPlayerStat(index, Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal victim As Long, Optional ByVal attacker As Long = 0) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerDodge = False
    If attacker > 0 Then
        rate = GetPlayerStat(victim, Agility) - GetPlayerStat(attacker, Agility)
    Else
        rate = GetPlayerStat(victim, Agility)
    End If
    
    If rate <= 0 Then
        rate = 0
    Else
        rate = rate * 0.35
    End If

    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal victim As Long, Optional ByVal attacker As Long = 0) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerParry = False
    
    If attacker > 0 Then
        rate = GetPlayerStat(victim, Strength) - GetPlayerStat(attacker, Strength)
    Else
        rate = GetPlayerStat(victim, Strength)
    End If
    
    
    If rate <= 0 Then
        rate = 0
    Else
        rate = rate * 0.07
    End If
    
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerParry = True
    End If
End Function

Public Function CanNpcBlock(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcBlock = False

    rate = 0
    ' TODO : make it based on shield lol
End Function

Public Function CanNpcCrit(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = GetNpcStat(NPCNum, Strength) * 0.09
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal index As Long, ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    If (GetNpcStat(NPCNum, Agility, index) < Player(index).stat(Stats.Agility)) Then
        rate = 0
    Else
        rate = (GetNpcStat(NPCNum, Agility, index) - Player(index).stat(Stats.Agility)) * 0.07
    End If
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal index As Long, ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    If (GetNpcStat(NPCNum, Strength, index) < Player(index).stat(Stats.Strength)) Then
        rate = 0
    Else
        rate = (GetNpcStat(NPCNum, Strength, index) - Player(index).stat(Stats.Strength)) * 0.07
    End If
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal index As Long, ByVal MapNPCNum As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim mapnum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(index, MapNPCNum) Then
    
        mapnum = GetPlayerMap(index)
        NPCNum = mapnpc(mapnum).NPC(MapNPCNum).Num
        
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(index, NPCNum) Then
            SendActionMsg mapnum, "Esquivado!", Pink, 1, (mapnpc(mapnum).NPC(MapNPCNum).x * 32), (mapnpc(mapnum).NPC(MapNPCNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(index, NPCNum) Then
            SendActionMsg mapnum, "Bloqueado!", Pink, 1, (mapnpc(mapnum).NPC(MapNPCNum).x * 32), (mapnpc(mapnum).NPC(MapNPCNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(MapNPCNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (GetNpcStat(NPCNum, Agility, mapnpc(mapnum).NPC(MapNPCNum).PetData.owner)))
        ' randomise from half to max hit
        Damage = RAND(Damage / 2, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCriticalHit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(index, MapNPCNum, Damage)
        Else
            Call PlayerMsg(index, "Tu ataque no hace nada.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal MapNPCNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapnum As Long
    Dim NPCNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If mapnpc(GetPlayerMap(attacker)).NPC(MapNPCNum).Num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(attacker)
    NPCNum = mapnpc(mapnum).NPC(MapNPCNum).Num
    
    'Pet check
    If mapnpc(mapnum).NPC(MapNPCNum).IsPet = YES Then
        If Not (CanPlayerAttackPet(mapnum, MapNPCNum, attacker)) Then
            Exit Function
        End If
    End If
    
    ' Make sure the npc isn't already dead
    If mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    'Can't attack own pet
    If TempPlayer(attacker).TempPetSlot = MapNPCNum Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
    
        ' exit out early
        If IsSpell Then
             If NPCNum > 0 Then
                If CanNPCBeAttacked(NPCNum) Then
                    TempPlayer(attacker).targetType = TARGET_TYPE_NPC
                    TempPlayer(attacker).target = MapNPCNum
                    SendTarget attacker
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If NPCNum > 0 And GetTickCount > TempPlayer(attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(attacker)
                Case DIR_UP
                    NpcX = mapnpc(mapnum).NPC(MapNPCNum).x
                    NpcY = mapnpc(mapnum).NPC(MapNPCNum).y + 1
                Case DIR_DOWN
                    NpcX = mapnpc(mapnum).NPC(MapNPCNum).x
                    NpcY = mapnpc(mapnum).NPC(MapNPCNum).y - 1
                Case DIR_LEFT
                    NpcX = mapnpc(mapnum).NPC(MapNPCNum).x + 1
                    NpcY = mapnpc(mapnum).NPC(MapNPCNum).y
                Case DIR_RIGHT
                    NpcX = mapnpc(mapnum).NPC(MapNPCNum).x - 1
                    NpcY = mapnpc(mapnum).NPC(MapNPCNum).y
            End Select

            If NpcX = GetPlayerX(attacker) Then
                If NpcY = GetPlayerY(attacker) Then
                    If CanNPCBeAttacked(NPCNum) Then
                        CanPlayerAttackNpc = True
                    Else
                        'ALATAR
                        If NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                            Call CheckTasks(attacker, QUEST_TYPE_GOTALK, NPCNum)
                            Call CheckTasks(attacker, QUEST_TYPE_GOGIVE, NPCNum)
                            Call CheckTasks(attacker, QUEST_TYPE_GOGET, NPCNum)
                            
                            If NPC(NPCNum).Quest = YES Then 'Alatar v1.2
                                If Player(attacker).PlayerQuest(NPC(NPCNum).Quest).Status = QUEST_COMPLETED Then
                                    If Quest(NPC(NPCNum).Quest).Repeat = YES Then
                                        Player(attacker).PlayerQuest(NPC(NPCNum).Quest).Status = QUEST_COMPLETED_BUT
                                        Exit Function
                                    End If
                                End If
                                If CanStartQuest(attacker, NPC(NPCNum).QuestNum) Then
                                    'if can start show the request message (speech1)
                                    QuestMessage attacker, NPC(NPCNum).QuestNum, Trim$(Quest(NPC(NPCNum).QuestNum).Speech(1)), NPC(NPCNum).QuestNum
                                    Exit Function
                                End If
                                If QuestInProgress(attacker, NPC(NPCNum).QuestNum) Then
                                    'if the quest is in progress show the meanwhile message (speech2)
                                    QuestMessage attacker, NPC(NPCNum).QuestNum, Trim$(Quest(NPC(NPCNum).QuestNum).Speech(2)), 0
                                    Exit Function
                                End If
                            End If
                        End If
                        '/ALATAR
                        If Len(Trim$(NPC(NPCNum).AttackSay)) > 0 Then
                            PlayerMsg attacker, Trim$(NPC(NPCNum).Name) & ": " & Trim$(NPC(NPCNum).AttackSay), White
                            'Call SendActionMsg(mapnum, Trim$(NPC(npcnum).Name) & ": " & Trim$(NPC(npcnum).AttackSay), SayColor, 1, mapnpc(mapnum).NPC(mapNpcNum).x * 32, mapnpc(mapnum).NPC(mapNpcNum).y * 32)
                            Call SpeechWindow(attacker, Trim$(NPC(NPCNum).AttackSay), NPCNum)
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal attacker As Long, ByVal MapNPCNum As Long, ByVal Damage As Long, Optional ByVal Spellnum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim DropNum As Integer
    Dim STR As Long
    Dim Def As Long
    Dim mapnum As Long
    Dim NPCNum As Long
    Dim Buffer As clsBuffer
    Dim npcexperience As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(attacker)
    NPCNum = mapnpc(mapnum).NPC(MapNPCNum).Num
    If NPCNum < 1 Then Exit Sub
    Name = Trim$(NPC(NPCNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount

    If Damage >= mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(attacker), "-" & mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.HP), BrightRed, 1, (mapnpc(mapnum).NPC(MapNPCNum).x * 32), (mapnpc(mapnum).NPC(MapNPCNum).y * 32)
        SendBlood GetPlayerMap(attacker), mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound attacker, mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y, SoundEntity.seSpell, Spellnum
        
        ' send animation
            If n > 0 Then
                If Not overTime Then
                    If Spellnum = 0 Then
                        Call SendAnimation(mapnum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y)
                    Else
                        Call SendAnimation(mapnum, Spell(Spellnum).SpellAnim, mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y)
                    End If
                End If
            End If

        ' Calculate exp to give attacker
        Exp = NPC(NPCNum).Exp * Options.npcexperience

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If
        
        'Pet ?
        If TempPlayer(attacker).TempPetSlot > 0 Then
            Call SharePetExp(attacker, TempPlayer(attacker).ActualPet, Exp, TempPlayer(attacker).PetExpPercent, True)
            
            'Auto Targetting
            If TempPlayer(attacker).PetHasOwnTarget = MapNPCNum Then
                'Objective Finished
                TempPlayer(attacker).PetHasOwnTarget = 0
            End If
                
        Else
            ' in party?
            If TempPlayer(attacker).inParty > 0 Then
                ' pass through party sharing function
                Party_ShareExp TempPlayer(attacker).inParty, Exp, attacker
            Else
                ' no party - keep exp for self
                GivePlayerEXP attacker, Exp
            End If
        End If
        
        'begin of the new system
        'Call the calculator for deciding which item has to be chanced
        DropNum = CalculateDropChances(NPCNum)
        'Drop the goods if they get it
        If DropNum > 0 Then
            n = Int(Rnd * NPC(NPCNum).Drops(DropNum).DropChance) + 1

            'Drop the selected item
            If n = 1 Then
                Call SpawnItem(NPC(NPCNum).Drops(DropNum).DropItem, NPC(NPCNum).Drops(DropNum).DropItemValue, mapnum, mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y)
            End If
        End If
        
        'Begin of the old system
        'n = Int(Rnd * NPC(npcnum).DropChance) + 1

        'If n = 1 Then
            'Call SpawnItem(NPC(npcnum).DropItem, NPC(npcnum).DropItemValue, mapNum, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y)
        'End If
        
        'end of the old system

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        mapnpc(mapnum).NPC(MapNPCNum).Num = 0
        mapnpc(mapnum).NPC(MapNPCNum).SpawnWait = GetTickCount
        mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.HP) = 0
        
        'Checks if NPC was a pet
        If mapnpc(mapnum).NPC(MapNPCNum).IsPet = YES Then
            Call PetDisband(mapnpc(mapnum).NPC(MapNPCNum).PetData.owner, mapnum) 'The pet was killed
        End If
        
        'Restart NPC's Target if it was mapnpcnum
        Call ResetMapNPCSTarget(mapnum, MapNPCNum)
        
        
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
        
        'ALATAR
        Call CheckTasks(attacker, QUEST_TYPE_GOSLAY, NPCNum)
        '/ALATAR
        
        'Player NPC info
        Player(attacker).NPC(NPCNum).Kills = Player(attacker).NPC(NPCNum).Kills + 1
        
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
        
       
    
    Else
        ' NPC not dead, just do the damage
        mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.HP) = mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.HP) - Damage
        
        'target the NPC
        TempPlayer(attacker).targetType = TARGET_TYPE_NPC
        TempPlayer(attacker).target = MapNPCNum
        SendTarget attacker
        
        ' Check for a weapon and say damage
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (mapnpc(mapnum).NPC(MapNPCNum).x * 32), (mapnpc(mapnum).NPC(MapNPCNum).y * 32)
        SendBlood GetPlayerMap(attacker), mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound attacker, mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y, SoundEntity.seSpell, Spellnum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If Spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, MapNPCNum)
            End If
        End If

        ' Set the NPC target to the player
        mapnpc(mapnum).NPC(MapNPCNum).targetType = 1 ' player
        mapnpc(mapnum).NPC(MapNPCNum).target = attacker
        
        'Set the NPC target to player's pet target
        If TempPlayer(attacker).TempPetSlot > 0 And TempPlayer(attacker).TempPetSlot < MAX_MAP_NPCS And TempPlayer(attacker).PetHasOwnTarget = 0 Then
            mapnpc(mapnum).NPC(TempPlayer(attacker).TempPetSlot).targetType = TARGET_TYPE_NPC
            mapnpc(mapnum).NPC(TempPlayer(attacker).TempPetSlot).target = MapNPCNum
            'Auto Targetting
            TempPlayer(attacker).PetHasOwnTarget = MapNPCNum
        End If
            

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(mapnpc(mapnum).NPC(MapNPCNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If mapnpc(mapnum).NPC(i).Num = mapnpc(mapnum).NPC(MapNPCNum).Num Then
                    mapnpc(mapnum).NPC(i).target = attacker
                    mapnpc(mapnum).NPC(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        mapnpc(mapnum).NPC(MapNPCNum).stopRegen = True
        mapnpc(mapnum).NPC(MapNPCNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If Spellnum > 0 Then
            If Spell(Spellnum).StunDuration > 0 Then StunNPC MapNPCNum, mapnum, Spellnum
            ' DoT
            If Spell(Spellnum).Duration > 0 Then
                AddDoT_Npc mapnum, MapNPCNum, Spellnum, attacker
            End If
        End If
        
        SendMapNpcVitals mapnum, MapNPCNum
    End If

    If Spellnum = 0 Then
        ' Reset attack timer
        TempPlayer(attacker).AttackTimer = GetTickCount
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal MapNPCNum As Long, ByVal index As Long)
Dim mapnum As Long, NPCNum As Long, blockAmount As Long, Damage As Long
Dim Buffer As clsBuffer
    ' Can the npc attack the player?
    If CanNpcAttackPlayer(MapNPCNum, index) Then
        mapnum = GetPlayerMap(index)
        NPCNum = mapnpc(mapnum).NPC(MapNPCNum).Num
        
        
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seNpc, mapnpc(mapnum).NPC(MapNPCNum).Num
        
        ' Send this packet so they can see the npc attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong ServerPackets.SNpcAttack
        Buffer.WriteLong MapNPCNum
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
        mapnpc(mapnum).NPC(MapNPCNum).AttackTimer = GetTickCount
        
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg mapnum, "¡Esquivado!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(index) Then
            SendActionMsg mapnum, "¡Bloqueado!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (GetPlayerStat(index, Agility) / 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "¡Crítico!", BrightCyan, 1, (mapnpc(mapnum).NPC(MapNPCNum).x * 32), (mapnpc(mapnum).NPC(MapNPCNum).y * 32)
        End If

        Damage = Damage / (GetPlayerDef(index) / 20)
        
        Damage = Damage * 0.8

        If Damage > 0 Then
            Call NpcAttackPlayer(MapNPCNum, index, Damage)
        Else
            SendActionMsg mapnum, "¡Evitado!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal MapNPCNum As Long, ByVal index As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapnum As Long
    Dim NPCNum As Long
    Dim Buffer As clsBuffer
    
    ' Check if player is loading
    If TempPlayer(index).IsLoading = True Then Exit Function
    
    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If mapnpc(GetPlayerMap(index)).NPC(MapNPCNum).Num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(index)
    NPCNum = mapnpc(mapnum).NPC(MapNPCNum).Num
    
    'Pet check
    If mapnpc(mapnum).NPC(MapNPCNum).IsPet = YES Then
        If Not (CanPetAttackPlayer(mapnum, MapNPCNum, index)) Then
            Exit Function
        End If
    End If

    ' Make sure the npc isn't already dead
    If mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    

    

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If
    
    'check if the NPC attacking us is actually our pet.
    'We don't want a rebellion on our hands now do we?
    If mapnpc(mapnum).NPC(MapNPCNum).PetData.owner = index Then Exit Function
    
    'Spell Check
    If IsSpell Then
        CanNpcAttackPlayer = True
    Else
        ' Make sure npcs dont attack more then once a second
        If GetTickCount < mapnpc(mapnum).NPC(MapNPCNum).AttackTimer + 1000 Then
            Exit Function
        End If
        
        ' Make sure they are on the same map
        If IsPlaying(index) Then
            If NPCNum > 0 Then

                ' Check if at same coordinates
                If (GetPlayerY(index) + 1 = mapnpc(mapnum).NPC(MapNPCNum).y) And (GetPlayerX(index) = mapnpc(mapnum).NPC(MapNPCNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) - 1 = mapnpc(mapnum).NPC(MapNPCNum).y) And (GetPlayerX(index) = mapnpc(mapnum).NPC(MapNPCNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = mapnpc(mapnum).NPC(MapNPCNum).y) And (GetPlayerX(index) + 1 = mapnpc(mapnum).NPC(MapNPCNum).x) Then
                            CanNpcAttackPlayer = True
                        Else
                            If (GetPlayerY(index) = mapnpc(mapnum).NPC(MapNPCNum).y) And (GetPlayerX(index) - 1 = mapnpc(mapnum).NPC(MapNPCNum).x) Then
                                CanNpcAttackPlayer = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
    End If

    
End Function

Sub NpcAttackPlayer(ByVal MapNPCNum As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim mapnum As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If mapnpc(GetPlayerMap(victim)).NPC(MapNPCNum).Num <= 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(victim)
    Name = Trim$(NPC(mapnpc(mapnum).NPC(MapNPCNum).Num).Name)
    
    '' Send this packet so they can see the npc attacking
    ''Set Buffer = New clsBuffer
    'Buffer.WriteLong SNpcAttack
    'Buffer.WriteLong mapNpcNum
    'SendDataToMap mapNum, Buffer.ToArray()
    'Set Buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    mapnpc(mapnum).NPC(MapNPCNum).stopRegen = True
    mapnpc(mapnum).NPC(MapNPCNum).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, mapnpc(mapnum).NPC(MapNPCNum).Num
        
        'Drop Items If npc was a  pet
        If GetMapPetOwner(mapnum, MapNPCNum) > 0 Then
            If Not (GetLevelDifference(GetMapPetOwner(mapnum, MapNPCNum), victim) > 20) Then
                Call PlayerPVPDrops(victim)
            End If
        End If
        
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " ha sido asesinado por " & Name, BrightRed)

        ' Set NPC target to 0
        mapnpc(mapnum).NPC(MapNPCNum).target = 0
        mapnpc(mapnum).NPC(MapNPCNum).targetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        Call SendAnimation(mapnum, NPC(mapnpc(GetPlayerMap(victim)).NPC(MapNPCNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, victim)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, mapnpc(mapnum).NPC(MapNPCNum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
    End If

End Sub

Sub PetSpellItself(ByVal mapnum As Long, ByVal MapNPCNum As Long, ByVal SpellSlotNum As Long)
    Dim Spellnum As Long
    Dim InitDamage As Long
    Dim DidCast As Boolean
    Dim MPCost As Long
    
    ' Check for subscript out of range
        If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Then
                Exit Sub
        End If

        ' Check for subscript out of range
        If mapnpc(mapnum).NPC(MapNPCNum).Num <= 0 Then
                Exit Sub
        End If
   
        If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub
        
        ' The Variables
        Spellnum = NPC(mapnpc(mapnum).NPC(MapNPCNum).Num).Spell(SpellSlotNum)
        
        MPCost = Spell(Spellnum).MPCost

        ' Check if they have enough MP
        If mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) < MPCost Then
            Exit Sub
        End If
        
        If Spellnum = 0 Then
        Exit Sub
        End If
        
        DidCast = False
   
        ' CoolDown Time
        If mapnpc(mapnum).NPC(MapNPCNum).SpellTimer(SpellSlotNum) > GetTickCount Then Exit Sub
   
        ' Spell Types
                Select Case Spell(Spellnum).Type
                    Case SPELL_TYPE_HEALHP
                        If Not Spell(Spellnum).IsAoE Then
                                InitDamage = GetNPCSpellDamage(Spellnum, mapnum, MapNPCNum, GetMapPetOwner(mapnum, MapNPCNum))
                                SpellNpc_Effect Vitals.HP, True, MapNPCNum, InitDamage, Spellnum, mapnum
                                DidCast = True
                        End If
                    Case SPELL_TYPE_HEALMP
                        If Not Spell(Spellnum).IsAoE Then
                                InitDamage = GetNPCSpellDamage(Spellnum, mapnum, MapNPCNum, GetMapPetOwner(mapnum, MapNPCNum))
                                SpellNpc_Effect Vitals.mp, True, MapNPCNum, InitDamage, Spellnum, mapnum
                                DidCast = True
                        End If
                    End Select
                    
                    If DidCast Then
                        mapnpc(mapnum).NPC(MapNPCNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(Spellnum).CDTime * 1000
                        SendNpcAttackAnimation mapnum, MapNPCNum
                        If mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) - MPCost < 0 Then
                            mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) = 0
                        Else
                            mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) = mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) - MPCost
                        End If
                    Else
                        mapnpc(mapnum).NPC(MapNPCNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(Spellnum).CDTime * 1000
                    End If
End Sub

Sub PetSpellOwner(ByVal MapNPCNum As Long, ByVal index As Long, SpellSlotNum As Long)
    Dim mapnum As Long
    Dim Spellnum As Long
    Dim InitDamage As Long
    Dim MPCost As Long
    Dim DidCast As Boolean
    
    ' Check for subscript out of range
        If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or IsPlaying(index) = False Then
                Exit Sub
        End If

        ' Check for subscript out of range
        If mapnpc(GetPlayerMap(index)).NPC(MapNPCNum).Num <= 0 Then
                Exit Sub
        End If
   
        If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub
        
        ' The Variables
        mapnum = GetPlayerMap(index)
        Spellnum = NPC(mapnpc(mapnum).NPC(MapNPCNum).Num).Spell(SpellSlotNum)
        
        MPCost = Spell(Spellnum).MPCost

        ' Check if they have enough MP
        If mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) < MPCost Then
            Exit Sub
        End If
        
        If Spellnum = 0 Then
        Exit Sub
        End If
        
        DidCast = False
   
        ' CoolDown Time
        If mapnpc(mapnum).NPC(MapNPCNum).SpellTimer(SpellSlotNum) > GetTickCount Then Exit Sub
   
        ' Spell Types
                Select Case Spell(Spellnum).Type
                    Case SPELL_TYPE_HEALHP
                        If Not Spell(Spellnum).IsAoE Then
                            If isInRange(Spell(Spellnum).Range + 3, mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y, GetPlayerX(index), GetPlayerY(index)) Then
                                InitDamage = GetNPCSpellDamage(Spellnum, mapnum, MapNPCNum, GetMapPetOwner(mapnum, MapNPCNum))
                                SpellPlayer_Effect Vitals.HP, True, index, InitDamage, Spellnum
                                DidCast = True
                            End If
                        End If
                    Case SPELL_TYPE_HEALMP
                        If Not Spell(Spellnum).IsAoE Then
                            If isInRange(Spell(Spellnum).Range + 3, mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y, GetPlayerX(index), GetPlayerY(index)) Then
                                InitDamage = GetNPCSpellDamage(Spellnum, mapnum, MapNPCNum, GetMapPetOwner(mapnum, MapNPCNum))
                                SpellPlayer_Effect Vitals.mp, True, index, InitDamage, Spellnum
                                DidCast = True
                            End If
                        End If
                    End Select
                    
                    If DidCast Then
                        mapnpc(mapnum).NPC(MapNPCNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(Spellnum).CDTime * 1000
                        SendNpcAttackAnimation mapnum, MapNPCNum
                        If mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) - MPCost < 0 Then
                            mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) = 0
                        Else
                            mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) = mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) - MPCost
                        End If
                    Else
                        mapnpc(mapnum).NPC(MapNPCNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(Spellnum).CDTime * 1000
                    End If
End Sub
Sub NpcSpellPlayer(ByVal MapNPCNum As Long, ByVal victim As Long, SpellSlotNum As Long)
        Dim mapnum As Long
        Dim i As Long
        Dim n As Long
        Dim Spellnum As Long
        Dim Buffer As clsBuffer
        Dim InitDamage As Long
        Dim Damage As Long
        Dim index As Long
        Dim PetOwner As Long
        Dim MPCost As Long
        Dim DidCast As Boolean

        ' Check for subscript out of range
        If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
                Exit Sub
        End If

        ' Check for subscript out of range
        If mapnpc(GetPlayerMap(victim)).NPC(MapNPCNum).Num <= 0 Then
                Exit Sub
        End If
   
        If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub
        
        ' The Variables
        mapnum = GetPlayerMap(victim)
        Spellnum = NPC(mapnpc(mapnum).NPC(MapNPCNum).Num).Spell(SpellSlotNum)
        
        If Spellnum = 0 Then
        Exit Sub
        End If
        
        MPCost = Spell(Spellnum).MPCost

        ' Check if they have enough MP
        If mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) < MPCost Then
            Exit Sub
        End If
        
   
        ' CoolDown Time
        If mapnpc(mapnum).NPC(MapNPCNum).SpellTimer(SpellSlotNum) > GetTickCount Then Exit Sub
   
        ' Spell Types
                Select Case Spell(Spellnum).Type
                        ' AOE Healing Spells
                        Case SPELL_TYPE_HEALHP
                        
                        'Auto-Heal?
                        If ChoosePetSpellingMethod(victim, MapNPCNum, SpellSlotNum, Spellnum) = True Then Exit Sub
                        ' Make sure an npc waits for the spell to cooldown
                            If Spell(Spellnum).IsAoE Then
                                'For i = 1 To MAX_PLAYERS
                                                   
                                                
                                'Next
                            Else
                                If Not HasNPCMaxVital(HP, mapnum, MapNPCNum, GetMapPetOwner(mapnum, MapNPCNum)) Then
                                    ' Non AOE Healing Spells
                                    InitDamage = GetNPCSpellDamage(Spellnum, mapnum, MapNPCNum, GetMapPetOwner(mapnum, MapNPCNum))
                                    SpellNpc_Effect Vitals.HP, True, MapNPCNum, InitDamage, Spellnum, mapnum
                                    DidCast = True
                                End If
                            End If
                                
                           
                        ' AOE Damaging Spells
                        Case SPELL_TYPE_DAMAGEHP
                        ' Make sure an npc waits for the spell to cooldown
                                If Spell(Spellnum).IsAoE Then
                                        For i = 1 To Player_HighIndex
                                                If CanNpcAttackPlayer(MapNPCNum, i, True) Then
                                                        If isInRange(Spell(Spellnum).AoE, mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y, GetPlayerX(i), GetPlayerY(i)) Then
                                                                InitDamage = GetNPCSpellDamage(Spellnum, mapnum, MapNPCNum, GetMapPetOwner(mapnum, MapNPCNum))
                                                                Damage = InitDamage - Player(i).stat(Stats.willpower)
                                                                If Damage <= 0 Then
                                                                    SendActionMsg GetPlayerMap(i), "RESISTISTE!", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
                                                                Else
                                                                    If Spell(Spellnum).StunDuration > 0 Then
                                                                        StunPlayer victim, Spellnum
                                                                    End If
                                                                    SendAnimation mapnum, Spell(Spellnum).SpellAnim, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
                                                                    NpcAttackPlayer MapNPCNum, i, Damage
                                                                    DidCast = True
                                                                End If
                                                        End If
                                                End If
                                        Next
                                ' Non AoE Damaging Spells
                                Else
                                    If CanNpcAttackPlayer(MapNPCNum, victim, True) Then
                                        If isInRange(Spell(Spellnum).Range, mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y, GetPlayerX(victim), GetPlayerY(victim)) Then
                                        InitDamage = GetNPCSpellDamage(Spellnum, mapnum, MapNPCNum, GetMapPetOwner(mapnum, MapNPCNum))
                                        Damage = InitDamage - Player(victim).stat(Stats.willpower)
                                                If Damage <= 0 Then
                                                        SendActionMsg GetPlayerMap(victim), "RESISTISTE!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
                                                Else
                                                    If Spell(Spellnum).StunDuration > 0 Then
                                                        StunPlayer victim, Spellnum
                                                    End If
                                                    NpcAttackPlayer MapNPCNum, victim, Damage
                                                    SendAnimation mapnum, Spell(Spellnum).SpellAnim, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
                                                    DidCast = True
                                                End If
                                        End If
                                    End If
                                End If

                                Case SPELL_TYPE_DAMAGEMP
                                    ' Make sure an npc waits for the spell to cooldown
                                    If Spell(Spellnum).IsAoE Then
                                        For i = 1 To Player_HighIndex
                                            If CanNpcAttackPlayer(MapNPCNum, victim, True) Then
                                                If isInRange(Spell(Spellnum).AoE, mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y, GetPlayerX(i), GetPlayerY(i)) Then
                                                    InitDamage = GetNPCSpellDamage(Spellnum, mapnum, MapNPCNum, GetMapPetOwner(mapnum, MapNPCNum))
                                                    Damage = InitDamage - Player(i).stat(Stats.willpower)
                                                    If Damage <= 0 Then
                                                        SendActionMsg GetPlayerMap(i), "¡RESISTISTE!", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
                                                    Else
                                                        SpellPlayer_Effect Vitals.mp, False, victim, Damage, Spellnum
                                                        DidCast = True
                                                    End If
                                                End If
                                            End If
                                        Next
                                    ' Non AoE DamagingMP Spells
                                    Else
                                        If isInRange(Spell(Spellnum).Range, mapnpc(mapnum).NPC(MapNPCNum).x, mapnpc(mapnum).NPC(MapNPCNum).y, GetPlayerX(victim), GetPlayerY(victim)) Then
                                            If CanNpcAttackPlayer(MapNPCNum, victim, True) Then
                                                InitDamage = GetNPCSpellDamage(Spellnum, mapnum, MapNPCNum, GetMapPetOwner(mapnum, MapNPCNum))
                                                Damage = InitDamage - Player(victim).stat(Stats.willpower)
                                                If Damage <= 0 Then
                                                    SendActionMsg GetPlayerMap(victim), "¡RESISTISTE!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
                                                Else
                                                    SpellPlayer_Effect Vitals.mp, False, victim, Damage, Spellnum
                                                    DidCast = True
                                                End If
                                            End If
                                        End If
                                    End If
                                    
                                    End Select
                        
                                    If DidCast Then
                                        mapnpc(mapnum).NPC(MapNPCNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(Spellnum).CDTime * 1000
                                        SendNpcAttackAnimation mapnum, MapNPCNum
                                        If mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) - MPCost < 0 Then
                                            mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) = 0
                                        Else
                                            mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) = mapnpc(mapnum).NPC(MapNPCNum).vital(Vitals.mp) - MPCost
                                        End If
                                        SendMapNpcVitals mapnum, MapNPCNum
                                    Else
                                        mapnpc(mapnum).NPC(MapNPCNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(Spellnum).CDTime * 1000
                                    End If

End Sub

Sub NpcSpellNpc(ByVal mapnum As Long, ByVal aMapNPCNum As Long, ByVal vMapNPCNum As Long, SpellSlotNum As Long)
        Dim i As Long
        Dim n As Long
        Dim Spellnum As Long
        Dim Buffer As clsBuffer
        Dim InitDamage As Long
        Dim Damage As Long
        Dim MaxHeals As Long
        Dim index As Long
        Dim PetOwner As Long
        Dim MPCost As Long
        Dim DidCast As Boolean
        
        ' Check for subscript out of range
        If aMapNPCNum <= 0 Or aMapNPCNum > MAX_MAP_NPCS Or vMapNPCNum <= 0 Or vMapNPCNum > MAX_MAP_NPCS Then
                Exit Sub
        End If

        ' Check for subscript out of range
        If mapnpc(mapnum).NPC(aMapNPCNum).Num <= 0 Or mapnpc(mapnum).NPC(vMapNPCNum).Num <= 0 Then
                Exit Sub
        End If
   
        If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub

        ' The Variables
        Spellnum = NPC(mapnpc(mapnum).NPC(aMapNPCNum).Num).Spell(SpellSlotNum)
        
        If Spellnum = 0 Then
        Exit Sub
        End If
        
        MPCost = Spell(Spellnum).MPCost

        ' Check if they have enough MP
        If mapnpc(mapnum).NPC(aMapNPCNum).vital(Vitals.mp) < MPCost Then
            Exit Sub
        End If
        
        'set cast to false
        DidCast = False
   
        ' CoolDown Time
        If mapnpc(mapnum).NPC(aMapNPCNum).SpellTimer(SpellSlotNum) > GetTickCount Then Exit Sub
   
        ' Spell Types
                Select Case Spell(Spellnum).Type
                        ' AOE Healing Spells
                        Case SPELL_TYPE_HEALHP
                        
                            'Testing if Pet Heals Player
                            If ChoosePetSpellingMethod(GetMapPetOwner(mapnum, aMapNPCNum), aMapNPCNum, SpellSlotNum, Spellnum) = True Then Exit Sub
                            
                        ' Make sure an npc waits for the spell to cooldown
                            If Spell(Spellnum).IsAoE Then
                                For i = 1 To MAX_MAP_NPCS
                                    If mapnpc(mapnum).NPC(i).Num > 0 Then
                                        If mapnpc(mapnum).NPC(i).vital(Vitals.HP) > 0 Then
                                            If isInRange(Spell(Spellnum).AoE, mapnpc(mapnum).NPC(aMapNPCNum).x, mapnpc(mapnum).NPC(aMapNPCNum).y, mapnpc(mapnum).NPC(i).x, mapnpc(mapnum).NPC(i).y) Then
                                                InitDamage = GetNPCSpellDamage(Spellnum, mapnum, aMapNPCNum, GetMapPetOwner(mapnum, aMapNPCNum))
                                                Select Case GetMapPetOwner(mapnum, aMapNPCNum)
                                                Case 0
                                                    SpellNpc_Effect Vitals.HP, True, i, InitDamage, Spellnum, mapnum
                                                Case Is > 0
                                                    If i = aMapNPCNum Then
                                                        If Not HasNPCMaxVital(HP, mapnum, aMapNPCNum, i) Then
                                                            SpellNpc_Effect Vitals.HP, True, i, InitDamage, Spellnum, mapnum
                                                        End If
                                                    End If
                                                End Select
                                                DidCast = True
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                ' Non AOE Healing Spells
                                If Not HasNPCMaxVital(HP, mapnum, aMapNPCNum, GetMapPetOwner(mapnum, aMapNPCNum)) Then
                                    InitDamage = GetNPCSpellDamage(Spellnum, mapnum, aMapNPCNum, GetMapPetOwner(mapnum, aMapNPCNum))
                                    SpellNpc_Effect Vitals.HP, True, aMapNPCNum, InitDamage, Spellnum, mapnum
                                    DidCast = True
                                End If
                            End If
                                
                           
                        ' AOE Damaging Spells
                        Case SPELL_TYPE_DAMAGEHP
                        ' Make sure an npc waits for the spell to cooldown
                            If Spell(Spellnum).IsAoE Then
                                For i = 1 To MAX_MAP_NPCS
                                    If CanNpcAttackNpc(mapnum, aMapNPCNum, vMapNPCNum, True) Then
                                        If isInRange(Spell(Spellnum).AoE, mapnpc(mapnum).NPC(aMapNPCNum).x, mapnpc(mapnum).NPC(aMapNPCNum).y, mapnpc(mapnum).NPC(vMapNPCNum).x, mapnpc(mapnum).NPC(vMapNPCNum).y) Then
                                            InitDamage = GetNPCSpellDamage(Spellnum, mapnum, aMapNPCNum, GetMapPetOwner(mapnum, aMapNPCNum))
                                            Damage = InitDamage - GetNpcStat(mapnpc(mapnum).NPC(vMapNPCNum).Num, Intelligence, mapnpc(mapnum).NPC(vMapNPCNum).PetData.owner)
                                            If Damage <= 0 Then
                                                SendActionMsg mapnum, "RESIST", Pink, 1, (mapnpc(mapnum).NPC(vMapNPCNum).x) * 32, (mapnpc(mapnum).NPC(vMapNPCNum).y * 32)
                                            Else
                                                If Spell(Spellnum).StunDuration > 0 Then
                                                    StunNPC vMapNPCNum, mapnum, Spellnum
                                                End If
                                                SendAnimation mapnum, Spell(Spellnum).SpellAnim, mapnpc(mapnum).NPC(vMapNPCNum).x, mapnpc(mapnum).NPC(vMapNPCNum).y, TARGET_TYPE_NPC, vMapNPCNum
                                                NpcAttackNpc mapnum, aMapNPCNum, vMapNPCNum, Damage
                                                DidCast = True
                                            End If
                                        End If
                                    End If
                                Next
                                
                            ' Non AoE Damaging Spells
                            Else
                                If CanNpcAttackNpc(mapnum, aMapNPCNum, vMapNPCNum, True) Then
                                    If isInRange(Spell(Spellnum).Range, mapnpc(mapnum).NPC(aMapNPCNum).x, mapnpc(mapnum).NPC(aMapNPCNum).y, mapnpc(mapnum).NPC(vMapNPCNum).x, mapnpc(mapnum).NPC(vMapNPCNum).y) Then
                                        InitDamage = GetNPCSpellDamage(Spellnum, mapnum, aMapNPCNum, GetMapPetOwner(mapnum, aMapNPCNum))
                                        Damage = InitDamage - GetNpcStat(mapnpc(mapnum).NPC(vMapNPCNum).Num, Intelligence, mapnpc(mapnum).NPC(vMapNPCNum).PetData.owner)
                                        If Damage <= 0 Then
                                            SendActionMsg mapnum, "RESIST", Pink, 1, (mapnpc(mapnum).NPC(vMapNPCNum).x) * 32, (mapnpc(mapnum).NPC(vMapNPCNum).y * 32)
                                        Else
                                            If Spell(Spellnum).StunDuration > 0 Then
                                                StunNPC vMapNPCNum, mapnum, Spellnum
                                            End If
                                            SendAnimation mapnum, Spell(Spellnum).SpellAnim, mapnpc(mapnum).NPC(vMapNPCNum).x, mapnpc(mapnum).NPC(vMapNPCNum).y, TARGET_TYPE_NPC, vMapNPCNum
                                            NpcAttackNpc mapnum, aMapNPCNum, vMapNPCNum, Damage
                                            DidCast = True
                                        End If
                                    End If
                                End If
                            End If
                                

                        Case SPELL_TYPE_DAMAGEMP
                        ' Make sure an npc waits for the spell to cooldown
                            If Spell(Spellnum).IsAoE Then
                                For i = 1 To MAX_MAP_NPCS
                                    If CanNpcAttackNpc(mapnum, aMapNPCNum, vMapNPCNum, True) Then
                                        If isInRange(Spell(Spellnum).AoE, mapnpc(mapnum).NPC(aMapNPCNum).x, mapnpc(mapnum).NPC(aMapNPCNum).y, mapnpc(mapnum).NPC(vMapNPCNum).x, mapnpc(mapnum).NPC(vMapNPCNum).y) Then
                                            InitDamage = GetNPCSpellDamage(Spellnum, mapnum, aMapNPCNum, GetMapPetOwner(mapnum, aMapNPCNum))
                                            Damage = InitDamage - GetNpcStat(mapnpc(mapnum).NPC(vMapNPCNum).Num, Intelligence, mapnpc(mapnum).NPC(vMapNPCNum).PetData.owner)
                                            If Damage <= 0 Then
                                                SendActionMsg mapnum, "RESIST", Pink, 1, (mapnpc(mapnum).NPC(vMapNPCNum).x) * 32, (mapnpc(mapnum).NPC(vMapNPCNum).y * 32)
                                            Else
                                                If Spell(Spellnum).StunDuration > 0 Then
                                                    StunNPC vMapNPCNum, mapnum, Spellnum
                                                End If
                                                If mapnpc(mapnum).NPC(vMapNPCNum).vital(Vitals.mp) > 0 Then
                                                    SpellNpc_Effect Vitals.mp, False, vMapNPCNum, Damage, Spellnum, mapnum
                                                    DidCast = True
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            ' Non AoE Damaging Spells
                            Else
                                If CanNpcAttackNpc(mapnum, aMapNPCNum, vMapNPCNum, True) Then
                                    If isInRange(Spell(Spellnum).Range, mapnpc(mapnum).NPC(aMapNPCNum).x, mapnpc(mapnum).NPC(aMapNPCNum).y, mapnpc(mapnum).NPC(vMapNPCNum).x, mapnpc(mapnum).NPC(vMapNPCNum).y) Then
                                        InitDamage = GetNPCSpellDamage(Spellnum, mapnum, aMapNPCNum, GetMapPetOwner(mapnum, aMapNPCNum))
                                        Damage = InitDamage - GetNpcStat(mapnpc(mapnum).NPC(vMapNPCNum).Num, Intelligence, mapnpc(mapnum).NPC(vMapNPCNum).PetData.owner)
                                        If Damage <= 0 Then
                                            SendActionMsg mapnum, "RESIST", Pink, 1, (mapnpc(mapnum).NPC(vMapNPCNum).x) * 32, (mapnpc(mapnum).NPC(vMapNPCNum).y * 32)
                                        Else
                                            If Spell(Spellnum).StunDuration > 0 Then
                                                StunNPC vMapNPCNum, mapnum, Spellnum
                                            End If
                                            If mapnpc(mapnum).NPC(vMapNPCNum).vital(Vitals.mp) > 0 Then
                                                SpellNpc_Effect Vitals.mp, False, vMapNPCNum, Damage, Spellnum, mapnum
                                                DidCast = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End Select
                        
                        If DidCast Then
                            mapnpc(mapnum).NPC(aMapNPCNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(Spellnum).CDTime * 1000
                            SendNpcAttackAnimation mapnum, aMapNPCNum
                            If mapnpc(mapnum).NPC(aMapNPCNum).vital(Vitals.mp) - MPCost < 0 Then
                                mapnpc(mapnum).NPC(aMapNPCNum).vital(Vitals.mp) = 0
                            Else
                                mapnpc(mapnum).NPC(aMapNPCNum).vital(Vitals.mp) = mapnpc(mapnum).NPC(aMapNPCNum).vital(Vitals.mp) - MPCost
                            End If
                            SendMapNpcVitals mapnum, aMapNPCNum
                        Else
                            mapnpc(mapnum).NPC(aMapNPCNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(Spellnum).CDTime * 1000
                        End If
                        
End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim Buffer As clsBuffer
Dim mapnum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, victim) Then
    
        mapnum = GetPlayerMap(attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim, attacker) Then
            SendActionMsg mapnum, "¡Esquivado!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(victim, attacker) Then
            SendActionMsg mapnum, "¡Bloqueado!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (GetPlayerStat(victim, Agility) / 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if can crit
        If CanPlayerCriticalHit(attacker) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
        End If

        Damage = Damage / (GetPlayerDef(victim) / 20)
        
        If Damage > 0 Then
            Call PlayerAttackPlayer(attacker, victim, Damage)
        Else
            Call PlayerMsg(attacker, "Tu ataque no hizo nada.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal IsProjectile As Boolean = False) As Boolean

    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If
    ' Check is victim is loading
    If TempPlayer(victim).IsLoading = True Then Exit Function
    
    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP
    
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If IsPlayerNeutral(victim) Then
            Call PlayerMsg(attacker, "¡Ésta es una zona segura!", BrightRed)
            Exit Function
        End If
    End If
    
    'Safe Mode
    If CheckSafeMode(attacker, victim) = True Then
        Exit Function
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "Los administradores no pueden atacar a otros usuarios.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "No puedes atacar a" & GetPlayerName(victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(attacker) < 10 Then
        Call PlayerMsg(attacker, "¡Estás bajo nivel 10, aún no puedes atacar a nadie!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(attacker, "¡" & GetPlayerName(victim) & " está bajo nivel 10, no puedes atacarle aún", BrightRed)
        Exit Function
    End If
    
    'Make sure the attacker's level isn't too high
    'If GetPlayerLevel(victim) + 10 < GetPlayerLevel(attacker) Then
        'Call PlayerMsg(attacker, "¡Vuestros niveles son demasiados diferentes!", BrightRed)
        'Exit Function
    'End If
    'Make sure the attacker's level isn't too low
    'If GetPlayerLevel(victim) - 10 > GetPlayerLevel(attacker) Then
        'Call PlayerMsg(attacker, "¡Vuestros niveles son demasiados diferentes!", BrightRed)
        'Exit Function
    'End If
    
    'make sure victim is not a guild partner
    If Player(attacker).GuildFileId > 1 Then
        If Player(attacker).GuildFileId = Player(victim).GuildFileId Then
            Call PlayerMsg(attacker, "¡" & GetPlayerName(victim) & " es miembro de tu clan!", BrightRed)
        Exit Function
        End If
    End If
    TempPlayer(attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(attacker).target = victim
    SendTarget attacker
    
    CanPlayerAttackPlayer = True
    
End Function
Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal Spellnum As Long = 0)
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, Spellnum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " ha sido asesinado por " & GetPlayerName(attacker), BrightRed)
        ' Calculate exp to give attacker
        Exp = CalculateLosenExp(attacker, victim, 8, 4)

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 0
        End If

        If Exp = 0 Then
            Call PlayerMsg(victim, "No perdistes experiencia.", BrightRed)
            Call PlayerMsg(attacker, "No recibistes experiencia.", BrightBlue)
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - Exp)
            SendEXP victim
            Call PlayerMsg(victim, "¡Has perdido " & Exp & " de experiencia!", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(attacker).inParty, Exp, attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, Exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(attacker) Then
                    If TempPlayer(i).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        'If IsPlayerNeutral(victim) Then
            'If IsPlayerNeutral(attacker) Then
                'Call SetPlayerPK(attacker, YES)
                'Call SendPlayerData(attacker)
                'Call GlobalMsg("¡" & GetPlayerName(attacker) & " se ha vuelto un asesino!", BrightRed)
            'End If

        'Else
            'Call GlobalMsg("¡" & GetPlayerName(victim) & " ha pagado el precio por ser un asesino!", BrightRed)
        'End If
        
        'ALATAR
        Call CheckTasks(attacker, QUEST_TYPE_GOKILL, victim)
        '/ALATAR
        
        'Only If victim level + 20 >= attacker level
        If Not (GetLevelDifference(attacker, victim) > 20) Then
            'Drop System
            Call PlayerPVPDrops(victim)
        End If
        
        Call OnDeath(victim, 0)
        
        Call SetPlayerJustice(attacker, victim)
        Call SendPlayerData(attacker)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, Spellnum
        
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If Spellnum > 0 Then
            If Spell(Spellnum).StunDuration > 0 Then StunPlayer victim, Spellnum
            ' DoT
            If Spell(Spellnum).Duration > 0 Then
                AddDoT_Player victim, Spellnum, attacker
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal index As Long, ByVal spellslot As Long)
    Dim Spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim targetType As Byte
    Dim target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    Spellnum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)
    
    If Spellnum <= 0 Or Spellnum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(index, Spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg index, "Habilidad recargándose", BrightRed
        Exit Sub
    End If

    MPCost = Spell(Spellnum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.mp) < MPCost Then
        Call PlayerMsg(index, "No tienes suficiente MP!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(Spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "Necesitas el lvl " & LevelReq & " para usar esta habilidad.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(Spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "Necesitas ser admin.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(Spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Solo la clase " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " puede utilizar esta magia.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(Spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(Spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(Spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(index).targetType
    target = TempPlayer(index).target
    Range = Spell(Spellnum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                PlayerMsg index, "No tienes un objetivo.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(target), GetPlayerY(target)) Then
                    PlayerMsg index, "Objetivo fuera de rango.", BrightRed
                Else
                    ' go through spell types
                    If Spell(Spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(Spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), mapnpc(mapnum).NPC(target).x, mapnpc(mapnum).NPC(target).y) Then
                    PlayerMsg index, "Objetivo fuera de rango.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(Spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(Spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation mapnum, Spell(Spellnum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendActionMsg mapnum, "Casting " & Trim$(Spell(Spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        TempPlayer(index).spellBuffer.Spell = spellslot
        TempPlayer(index).spellBuffer.Timer = GetTickCount
        TempPlayer(index).spellBuffer.target = TempPlayer(index).target
        TempPlayer(index).spellBuffer.tType = TempPlayer(index).targetType
        Exit Sub
    Else
        SendClearSpellBuffer index
    End If
End Sub

Public Sub CastSpell(ByVal index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
    Dim Spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
   
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
   
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    Spellnum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, Spellnum) Then Exit Sub

    MPCost = Spell(Spellnum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.mp) < MPCost Then
        Call PlayerMsg(index, "No tienes suficiente MP!", BrightRed)
        Exit Sub
    End If
   
    LevelReq = Spell(Spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "Debes tener nivel " & LevelReq & " para usar ésta habilidad.", BrightRed)
        Exit Sub
    End If
   
    AccessReq = Spell(Spellnum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "Debes ser administrador para usar esta habilidad.", BrightRed)
        Exit Sub
    End If
   
    ClassReq = Spell(Spellnum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Solo " & (Class(ClassReq).Name) & " puede usar esta habilidad.", BrightRed)
            Exit Sub
        End If
    End If
   
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(Spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(Spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(Spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
   
    ' set the vital
    vital = Spell(Spellnum).vital + (GetPlayerLevel(index) * 2) + (GetPlayerStat(index, Intelligence) * 2)
    AoE = Spell(Spellnum).AoE
    Range = Spell(Spellnum).Range
   
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(Spellnum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, index, vital, Spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.mp, True, index, vital, Spellnum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SetPlayerDir index, Spell(Spellnum).Dir
                    PlayerWarp index, Spell(Spellnum).Map, Spell(Spellnum).x, Spell(Spellnum).y
                    SendAnimation GetPlayerMap(index), Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
               
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                Else
                    x = mapnpc(mapnum).NPC(target).x
                    y = mapnpc(mapnum).NPC(target).y
                End If
               
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "El objetivo no está al alcance.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            Select Case Spell(Spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, True) Then
                                            SendAnimation mapnum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                             PlayerAttackPlayer index, i, vital - (Player(i).stat(Stats.willpower) * 5), Spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If mapnpc(mapnum).NPC(i).Num > 0 Then
                            If mapnpc(mapnum).NPC(i).vital(HP) > 0 Then
                                If isInRange(AoE, x, y, mapnpc(mapnum).NPC(i).x, mapnpc(mapnum).NPC(i).y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        SendAnimation mapnum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc index, i, vital - GetNpcStat(mapnpc(mapnum).NPC(i).Num, willpower, mapnpc(mapnum).NPC(i).PetData.owner), Spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(Spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(Spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.mp
                        increment = True
                    ElseIf Spell(Spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.mp
                        increment = False
                    End If
                   
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect VitalType, increment, i, vital, Spellnum
                                    DidCast = True
                                End If
                            End If
                        End If
                    End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If mapnpc(mapnum).NPC(i).Num > 0 Then
                            If mapnpc(mapnum).NPC(i).vital(HP) > 0 Then
                                If isInRange(AoE, x, y, mapnpc(mapnum).NPC(i).x, mapnpc(mapnum).NPC(i).y) Then
                                    SpellNpc_Effect VitalType, increment, i, vital, Spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
            End Select
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
           
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(target)
                y = GetPlayerY(target)
            Else
                x = mapnpc(mapnum).NPC(target).x
                y = mapnpc(mapnum).NPC(target).y
            End If
               
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "El objetivo no está al alcance.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
            
            If Spell(Spellnum).Type = SPELL_TYPE_DAMAGEHP Then
                If targetType = TARGET_TYPE_PLAYER And target = index Then
                    PlayerMsg index, "No puedes atacarte a ti mismo.", BrightRed
                    Exit Sub
                End If
            End If
            
            Select Case Spell(Spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, target, True) Then
                            If vital > 0 Then
                                SendAnimation mapnum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                PlayerAttackPlayer index, target, vital - (Player(target).stat(Stats.willpower) * 5), Spellnum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, target, True) Then
                            If vital > 0 Then
                                SendAnimation mapnum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                                PlayerAttackNpc index, target, vital + (Player(target).stat(Stats.Intelligence) * 20) - GetNpcStat(mapnpc(mapnum).NPC(target).Num, willpower, mapnpc(mapnum).NPC(target).PetData.owner), Spellnum
                                DidCast = True
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(Spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.mp
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(Spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.mp
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(Spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    End If
                   
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(Spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, target, True) Then
                                SpellPlayer_Effect VitalType, increment, target, vital, Spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, target, vital, Spellnum
                        End If
                    Else
                        If Spell(Spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, target, True) Then
                                SpellNpc_Effect VitalType, increment, target, vital, Spellnum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, target, vital, Spellnum, mapnum
                        End If
                    End If
            End Select
    End Select
   
    If DidCast Then
        Call SetPlayerVital(index, Vitals.mp, GetPlayerVital(index, Vitals.mp) - MPCost)
        Call SendVital(index, Vitals.mp)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
       
        TempPlayer(index).SpellCD(spellslot) = GetTickCount + (Spell(Spellnum).CDTime * 1000)
        Call SendCooldown(index, spellslot)
        SendActionMsg mapnum, Trim$(Spell(Spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
    End If
End Sub
Public Sub SpellPlayer_Effect(ByVal vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal Spellnum As Long)
Dim sSymbol As String * 1
Dim Colour As Long
Dim MaxVital As Long
Dim VitalComp As Long
            If Damage > 0 Then
        
                MaxVital = GetPlayerMaxVital(index, vital)
                If increment Then
                
                    If Spell(Spellnum).Duration > 0 Then
                        AddHoT_Player index, Spellnum
                    End If
                    'add vital
                    VitalComp = MaxVital
                    sSymbol = "+"
                    
                    If vital = Vitals.HP Then Colour = BrightGreen
                    If vital = Vitals.mp Then Colour = BrightBlue
                Else
                    'substract vital
                    VitalComp = 0
                    Damage = -Damage
                    
                    sSymbol = "-"
                    Colour = Blue
                End If
                
                If GetPlayerVital(index, vital) = VitalComp Then
                    'Time saver
                    Exit Sub
                ElseIf (GetPlayerVital(index, vital) + Damage >= VitalComp And increment) Or (GetPlayerVital(index, vital) + Damage <= VitalComp And Not increment) Then
                    SetPlayerVital index, vital, VitalComp
                Else
                    SetPlayerVital index, vital, GetPlayerVital(index, vital) + Damage
                End If
                
                
                If Spellnum > 0 Then
                    SendAnimation GetPlayerMap(index), Spell(Spellnum).SpellAnim, GetPlayerX(index), GetPlayerY(index), TARGET_TYPE_PLAYER, index
                End If
                SendActionMsg GetPlayerMap(index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                
                ' send the sound
                SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, Spellnum
                
                Call SendVital(index, vital)
                
        End If
End Sub

Public Sub SpellNpc_Effect(ByVal vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal Spellnum As Long, ByVal mapnum As Long)
Dim sSymbol As String * 1
Dim Colour As Long
Dim MaxVital As Long
Dim VitalComp As Long

'Do not us this procedeture on hp substracting

        If Damage > 0 Then
        
                MaxVital = GetNpcMaxVital(mapnpc(mapnum).NPC(index).Num, vital, mapnpc(mapnum).NPC(index).PetData.owner)
                If increment Then
                    'add vital
                    VitalComp = MaxVital
                    sSymbol = "+"
                    
                    If vital = Vitals.HP Then Colour = BrightGreen
                    If vital = Vitals.mp Then Colour = BrightBlue
                Else
                    'substract vital
                    VitalComp = 0
                    Damage = -Damage
                    
                    sSymbol = "-"
                    Colour = Blue
                End If
                
                If mapnpc(mapnum).NPC(index).vital(vital) = VitalComp Then
                    'Time saver
                    Exit Sub
                ElseIf (mapnpc(mapnum).NPC(index).vital(vital) + Damage >= MaxVital And increment) Or (mapnpc(mapnum).NPC(index).vital(vital) + Damage <= VitalComp And Not increment) Then
                    mapnpc(mapnum).NPC(index).vital(vital) = VitalComp
                Else
                    mapnpc(mapnum).NPC(index).vital(vital) = mapnpc(mapnum).NPC(index).vital(vital) + Damage
                End If
                
                
                SendAnimation mapnum, Spell(Spellnum).SpellAnim, mapnpc(mapnum).NPC(index).x, mapnpc(mapnum).NPC(index).y, TARGET_TYPE_NPC, index
                SendActionMsg mapnum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, mapnpc(mapnum).NPC(index).x * 32, mapnpc(mapnum).NPC(index).y * 32
                
                ' send the sound
                SendMapSound index, mapnpc(mapnum).NPC(index).x, mapnpc(mapnum).NPC(index).y, SoundEntity.seSpell, Spellnum
                
                Call SendMapNpcVitals(mapnum, index)
                
        End If
End Sub

Public Sub AddDoT_Player(ByVal index As Long, ByVal Spellnum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            If .Spell = Spellnum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = Spellnum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal index As Long, ByVal Spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).HoT(i)
            If .Spell = Spellnum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = Spellnum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal Spellnum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With mapnpc(mapnum).NPC(index).DoT(i)
            If .Spell = Spellnum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = Spellnum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal Spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With mapnpc(mapnum).NPC(index).HoT(i)
            If .Spell = Spellnum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = Spellnum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal index As Long, ByVal dotNum As Long)
    With TempPlayer(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, index, True) Then
                    PlayerAttackPlayer .Caster, index, Spell(.Spell).vital
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal index As Long, ByVal hotNum As Long)
    With TempPlayer(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                'SendActionMsg Player(Index).Map, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, Player(Index).x * 32, Player(Index).y * 32
                SetPlayerVital index, Vitals.HP, GetPlayerVital(index, Vitals.HP) + Spell(.Spell).vital
                SendVital index, Vitals.HP
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal dotNum As Long)
    With mapnpc(mapnum).NPC(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, index, True) Then
                    PlayerAttackNpc .Caster, index, Spell(.Spell).vital, , True
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal hotNum As Long)
        With mapnpc(mapnum).NPC(index).HoT(hotNum)
                If .Used And .Spell > 0 Then
                        ' time to tick?
                        If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                                        SendActionMsg mapnum, "+" & Spell(.Spell).vital, BrightGreen, ACTIONMSG_SCROLL, mapnpc(mapnum).NPC(index).x * 32, mapnpc(mapnum).NPC(index).y * 32
                                        mapnpc(mapnum).NPC(index).vital(Vitals.HP) = mapnpc(mapnum).NPC(index).vital(Vitals.HP) + Spell(.Spell).vital
                                Else
                                        SendActionMsg mapnum, "+" & Spell(.Spell).vital, BrightBlue, ACTIONMSG_SCROLL, mapnpc(mapnum).NPC(index).x * 32, mapnpc(mapnum).NPC(index).y * 32
                                        mapnpc(mapnum).NPC(index).vital(Vitals.mp) = mapnpc(mapnum).NPC(index).vital(Vitals.mp) + Spell(.Spell).vital
                                        
                                        If mapnpc(mapnum).NPC(index).vital(Vitals.mp) > GetNpcMaxVital(index, Vitals.mp, mapnpc(mapnum).NPC(index).PetData.owner) Then
                                                mapnpc(mapnum).NPC(index).vital(Vitals.mp) = GetNpcMaxVital(index, Vitals.mp, mapnpc(mapnum).NPC(index).PetData.owner)
                                        End If
                                End If
                                
                                .Timer = GetTickCount
                                ' check if DoT is still active - if NPC died it'll have been purged
                                If .Used And .Spell > 0 Then
                                        ' destroy hoT if finished
                                        If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                                                .Used = False
                                                .Spell = 0
                                                .Timer = 0
                                                .Caster = 0
                                                .StartTime = 0
                                        End If
                                End If
                        End If
                End If
        End With
End Sub

Public Sub StunPlayer(ByVal index As Long, ByVal Spellnum As Long)
    ' check if it's a stunning spell
    If Spell(Spellnum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(index).StunDuration = Spell(Spellnum).StunDuration
        TempPlayer(index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned index
        ' tell him he's stunned
        PlayerMsg index, "¡Estás paralizado!", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal index As Long, ByVal mapnum As Long, ByVal Spellnum As Long)
    ' check if it's a stunning spell
    If Spell(Spellnum).StunDuration > 0 Then
        ' set the values on index
        mapnpc(mapnum).NPC(index).StunDuration = Spell(Spellnum).StunDuration
        mapnpc(mapnum).NPC(index).StunTimer = GetTickCount
    End If
End Sub

Public Function CalculateDropChances(ByVal NPCNum As Long) As Integer

'equal probability distribution
Dim i As Byte, n As Byte, j As Byte
Dim BoolVect(1 To MAX_NPC_DROPS) As Boolean

For i = 1 To CByte(MAX_NPC_DROPS)
    BoolVect(i) = False
Next

n = 0
For i = 1 To CByte(MAX_NPC_DROPS)
    If NPC(NPCNum).Drops(i).DropItem > 0 Then
        n = n + 1
        BoolVect(i) = True
    End If
Next
If n = 0 Then
    CalculateDropChances = 0
Else
    i = RAND(1, n)
    j = 0
    For n = 1 To CByte(MAX_NPC_DROPS)
        If BoolVect(n) = True Then
            j = j + 1
        End If
        If j = i Then
            CalculateDropChances = n
            Exit Function
        End If
    Next
End If
            
End Function

' ###################################
' ##      NPC Attacking NPC     ##
' ###################################

Public Sub TryNpcAttackNpc(ByVal mapnum As Long, ByVal mapNPCnumAttacker As Long, ByVal mapNpcNumVictim As Long)
Dim NPCNumAttacker As Long, NPCNumVictim As Long, blockAmount As Long, Damage As Long

    ' Can the npc attack the player?
    If CanNpcAttackNpc(mapnum, mapNPCnumAttacker, mapNpcNumVictim) Then
    
        NPCNumAttacker = mapnpc(mapnum).NPC(mapNPCnumAttacker).Num
        NPCNumVictim = mapnpc(mapnum).NPC(mapNpcNumVictim).Num
        
        
        ' send the sound
        'SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seNpc, mapnpc(MapNum).NPC(MapNPCNum).Num
        
        ' send the sound
        SendSoundToMap mapnum, mapnpc(mapnum).NPC(mapNpcNumVictim).x, mapnpc(mapnum).NPC(mapNpcNumVictim).y, SoundEntity.seNpc, mapnpc(mapnum).NPC(mapNPCnumAttacker).Num
        
        ' check if NPC can avoid the attack
        If CanNpcDodgeNpc(mapnum, mapNPCnumAttacker, mapNpcNumVictim) Then
            SendActionMsg mapnum, "¡Esquivado!", Pink, 1, (mapnpc(mapnum).NPC(mapNpcNumVictim).x * 32), (mapnpc(mapnum).NPC(mapNpcNumVictim).y * 32)
            Exit Sub
        End If
        If CanNpcParryNpc(mapnum, mapNPCnumAttacker, mapNpcNumVictim) Then
            SendActionMsg mapnum, "¡Bloqueado!", Pink, 1, (mapnpc(mapnum).NPC(mapNpcNumVictim).x * 32), (mapnpc(mapnum).NPC(mapNpcNumVictim).y * 32)
            Exit Sub
        End If

        Damage = GetNpcDamage(NPCNumAttacker, mapnpc(mapnum).NPC(mapNPCnumAttacker).PetData.owner)

        'blockAmount = CanNpcBlock(mapNpcNumVictim)
        'Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, GetNpcStat(NPCNumVictim, endurance, mapnpc(mapnum).NPC(mapNpcNumVictim).PetData.owner))
        ' randomise from half to max hit
        Damage = RAND(Damage / 2, Damage)
        ' * 1.5 if crit hit
        
        If CanNpcCrit(NPCNumAttacker) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "¡Crítico!", BrightCyan, 1, (mapnpc(mapnum).NPC(mapNPCnumAttacker).x * 32), (mapnpc(mapnum).NPC(mapNPCnumAttacker).y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackNpc(mapnum, mapNPCnumAttacker, mapNpcNumVictim, Damage)
        Else
            SendActionMsg mapnum, "¡Evitado!", Cyan, 1, mapnpc(mapnum).NPC(mapNpcNumVictim).x * 32, mapnpc(mapnum).NPC(mapNpcNumVictim).y * 32
        End If
    End If
End Sub



Function CanNpcAttackNpc(ByVal mapnum As Long, ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim aNPCNum As Long
    Dim vNPCNum As Long
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long
    
    CanNpcAttackNpc = False

    ' Check for subscript out of range
    If attacker <= 0 Or attacker > MAX_MAP_NPCS Then
        Exit Function
    End If
    
    If victim <= 0 Or victim > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If mapnpc(mapnum).NPC(attacker).Num <= 0 Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If mapnpc(mapnum).NPC(victim).Num <= 0 Then
        Exit Function
    End If

    aNPCNum = mapnpc(mapnum).NPC(attacker).Num
    vNPCNum = mapnpc(mapnum).NPC(victim).Num
    
    
    If aNPCNum <= 0 Then Exit Function
    If vNPCNum <= 0 Then Exit Function
    
    'Pet check
    If mapnpc(mapnum).NPC(attacker).IsPet = YES And mapnpc(mapnum).NPC(victim).IsPet = YES Then
        If Not (CanPetAttackPet(mapnum, attacker, victim)) Then
            Exit Function
        End If
    End If
    
    'Check npc type
    If CanNPCBeAttacked(vNPCNum) = False Then Exit Function

    ' Make sure the npcs arent already dead
    If mapnpc(mapnum).NPC(attacker).vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure the npc isn't already dead
    If mapnpc(mapnum).NPC(victim).vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    If IsSpell Then
        CanNpcAttackNpc = True
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < mapnpc(mapnum).NPC(attacker).AttackTimer + 1000 Then
        Exit Function
    End If
    
    mapnpc(mapnum).NPC(attacker).AttackTimer = GetTickCount
    
    AttackerX = mapnpc(mapnum).NPC(attacker).x
    AttackerY = mapnpc(mapnum).NPC(attacker).y
    VictimX = mapnpc(mapnum).NPC(victim).x
    VictimY = mapnpc(mapnum).NPC(victim).y

    ' Check if at same coordinates
    If (VictimY + 1 = AttackerY) And (VictimX = AttackerX) Then
        CanNpcAttackNpc = True
    Else

        If (VictimY - 1 = AttackerY) And (VictimX = AttackerX) Then
            CanNpcAttackNpc = True
        Else

            If (VictimY = AttackerY) And (VictimX + 1 = AttackerX) Then
                CanNpcAttackNpc = True
            Else

                If (VictimY = AttackerY) And (VictimX - 1 = AttackerX) Then
                    CanNpcAttackNpc = True
                End If
            End If
        End If
    End If

End Function

Sub NpcAttackNpc(ByVal mapnum As Long, ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim aNPCNum As Long
    Dim vNPCNum As Long
    Dim n As Long
    Dim PetOwner As Long
    Dim DropNum As Integer
    
    If attacker <= 0 Or attacker > MAX_MAP_NPCS Then Exit Sub
    If victim <= 0 Or victim > MAX_MAP_NPCS Then Exit Sub
    
    If Damage <= 0 Then Exit Sub
    
    aNPCNum = mapnpc(mapnum).NPC(attacker).Num
    vNPCNum = mapnpc(mapnum).NPC(victim).Num
    
    If aNPCNum <= 0 Then Exit Sub
    If vNPCNum <= 0 Then Exit Sub
    
    'set the victim's target to the pet attacking it
    mapnpc(mapnum).NPC(victim).targetType = 2 'Npc
    mapnpc(mapnum).NPC(victim).target = attacker
    
    ' set the regen timer
    mapnpc(mapnum).NPC(attacker).stopRegen = True
    mapnpc(mapnum).NPC(victim).stopRegenTimer = GetTickCount
    
    ' Send this packet so they can see the person attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong attacker
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing

    If Damage >= mapnpc(mapnum).NPC(victim).vital(Vitals.HP) Then
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (mapnpc(mapnum).NPC(victim).x * 32), (mapnpc(mapnum).NPC(victim).y * 32)
        SendBlood mapnum, mapnpc(mapnum).NPC(victim).x, mapnpc(mapnum).NPC(victim).y
        
        ' Set NPC target to 0
        mapnpc(mapnum).NPC(attacker).target = 0
        mapnpc(mapnum).NPC(attacker).targetType = 0
        'reset the targetter for the player
        
        'Call the calculator for deciding which item has to be chanced
        DropNum = CalculateDropChances(vNPCNum)
        
        'Drop the goods if they get it
        If DropNum > 0 Then
            n = Int(Rnd * NPC(vNPCNum).Drops(DropNum).DropChance) + 1

            'Drop the selected item
            If n = 1 Then
                Call SpawnItem(NPC(vNPCNum).Drops(DropNum).DropItem, NPC(vNPCNum).Drops(DropNum).DropItemValue, mapnum, mapnpc(mapnum).NPC(victim).x, mapnpc(mapnum).NPC(victim).y)
            End If
        End If
        
        Call ResetMapNPCSTarget(mapnum, victim)
        
        If mapnpc(mapnum).NPC(attacker).IsPet = YES Then
            TempPlayer(mapnpc(mapnum).NPC(attacker).PetData.owner).target = 0
            TempPlayer(mapnpc(mapnum).NPC(attacker).PetData.owner).targetType = TARGET_TYPE_NONE
            
            PetOwner = mapnpc(mapnum).NPC(attacker).PetData.owner
            
            SendTarget PetOwner
            
            Call SharePetExp(mapnpc(mapnum).NPC(attacker).PetData.owner, TempPlayer(mapnpc(mapnum).NPC(attacker).PetData.owner).ActualPet, NPC(mapnpc(mapnum).NPC(victim).Num).Exp, TempPlayer(mapnpc(mapnum).NPC(attacker).PetData.owner).PetExpPercent, False)
            'objective finished
            TempPlayer(PetOwner).PetHasOwnTarget = 0
            
        End If
                      
        If mapnpc(mapnum).NPC(victim).IsPet = YES Then
            'Get the pet owners' index
            PetOwner = mapnpc(mapnum).NPC(victim).PetData.owner
            'Set the NPC's target on the owner now
            mapnpc(mapnum).NPC(attacker).targetType = 1 'player
            mapnpc(mapnum).NPC(attacker).target = PetOwner
            
            'objective finished
            TempPlayer(PetOwner).PetHasOwnTarget = 0
            'Set Spawn time
            TempPlayer(PetOwner).PetSpawnWait = GetTickCount
            'Pet Died
             mapnpc(mapnum).NPC(victim).vital(Vitals.HP) = 0
            'Disband the pet
            PetDisband PetOwner, GetPlayerMap(PetOwner)
                   
        End If
        
        ' Reset victim's stuff so it dies in loop
        mapnpc(mapnum).NPC(victim).Num = 0
        mapnpc(mapnum).NPC(victim).SpawnWait = GetTickCount
        mapnpc(mapnum).NPC(victim).vital(Vitals.HP) = 0
               
        ' send npc death packet to map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong victim
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
        
        If PetOwner > 0 Then
            PetFollowOwner PetOwner
        End If
    Else
        ' npc not dead, just do the damage
        mapnpc(mapnum).NPC(victim).vital(Vitals.HP) = mapnpc(mapnum).NPC(victim).vital(Vitals.HP) - Damage
       
        ' Say damage
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (mapnpc(mapnum).NPC(victim).x * 32), (mapnpc(mapnum).NPC(victim).y * 32)
        SendBlood mapnum, mapnpc(mapnum).NPC(victim).x, mapnpc(mapnum).NPC(victim).y
        
        ' set the regen timer
        mapnpc(mapnum).NPC(victim).stopRegen = True
        mapnpc(mapnum).NPC(victim).stopRegenTimer = GetTickCount
    End If
    
    'Send both Npc's Vitals to the client
    SendMapNpcVitals mapnum, attacker
    SendMapNpcVitals mapnum, victim

End Sub

Function CalculateLosenExp(ByVal attacker As Long, ByVal victim As Long, ByVal MaxFactor As Byte, ByVal MinFactor As Byte) As Long

Dim difference As Long

difference = GetPlayerLevel(attacker) - GetPlayerLevel(victim)

If difference <= 0 Then
    CalculateLosenExp = GetPlayerExp(victim) / 4
Else
    CalculateLosenExp = GetPlayerExp(victim) * (-difference / ((MAX_LEVELS - 10) * MinFactor) + CLng(MaxFactor) * -1)
End If

If CalculateLosenExp < 0 Then
    CalculateLosenExp = 0
End If

End Function



Public Function CanNpcDodgeNpc(ByVal mapnum As Long, ByVal attacker As Long, ByVal victim As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodgeNpc = False
    rate = (GetNpcStat(mapnpc(mapnum).NPC(attacker).Num, Agility, mapnpc(mapnum).NPC(attacker).PetData.owner) - GetNpcStat(mapnpc(mapnum).NPC(victim).Num, Agility, mapnpc(mapnum).NPC(victim).PetData.owner))
    If rate < 0 Then
        rate = 0
    Else
        rate = rate * 0.2
    End If
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcDodgeNpc = True
    End If


End Function

Public Function CanNpcParryNpc(ByVal mapnum As Long, ByVal attacker As Long, ByVal victim As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParryNpc = False
    rate = (GetNpcStat(mapnpc(mapnum).NPC(attacker).Num, Strength, mapnpc(mapnum).NPC(attacker).PetData.owner) - GetNpcStat(mapnpc(mapnum).NPC(victim).Num, Strength, mapnpc(mapnum).NPC(victim).PetData.owner))
    If rate < 0 Then
        rate = 0
    Else
        rate = rate * 0.07
    End If
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcParryNpc = True
    End If


End Function





Public Function CanPetAttackPlayer(ByVal mapnum As Long, ByVal MapNPCNum As Long, ByVal victim As Long) As Boolean
    CanPetAttackPlayer = False
    
    'Check if pet
    If mapnpc(mapnum).NPC(MapNPCNum).IsPet = NO Then Exit Function
    
    ' Check if map is attackable
    If Not Map(mapnum).Moral = MAP_MORAL_NONE Then
        If IsPlayerNeutral(victim) Then
            Exit Function
        End If
    End If

    ' Check to make sure that they dont have access
    If GetPlayerAccess(mapnpc(mapnum).NPC(MapNPCNum).PetData.owner) > ADMIN_MONITOR Then
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(mapnpc(mapnum).NPC(MapNPCNum).PetData.owner) < 10 Then
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Exit Function
    End If
    
    
    'make sure victim is not a guild partner
    If Player(mapnpc(mapnum).NPC(MapNPCNum).PetData.owner).GuildFileId > 1 Then
        If Player(mapnpc(mapnum).NPC(MapNPCNum).PetData.owner).GuildFileId = Player(victim).GuildFileId Then
        Exit Function
        End If
    End If
    
    CanPetAttackPlayer = True
    
End Function

Public Function CanPlayerAttackPet(ByVal mapnum As Long, ByVal MapNPCNum As Long, ByVal attacker As Long) As Boolean
    CanPlayerAttackPet = False
    
    'Check if pet
    If mapnpc(mapnum).NPC(MapNPCNum).IsPet = NO Then Exit Function
    
    ' Check if map is attackable
    If Not Map(mapnum).Moral = MAP_MORAL_NONE Then
        If IsPlayerNeutral(mapnpc(mapnum).NPC(MapNPCNum).PetData.owner) Then
            Exit Function
        End If
    End If

    ' Check to make sure that they dont have access
    If GetPlayerAccess(attacker) > ADMIN_MONITOR Then
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(mapnpc(mapnum).NPC(MapNPCNum).PetData.owner) > ADMIN_MONITOR Then
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(attacker) < 10 Then
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(mapnpc(mapnum).NPC(MapNPCNum).PetData.owner) < 10 Then
        Exit Function
    End If
    
    
    'make sure victim is not a guild partner
    If Player(attacker).GuildFileId > 1 Then
        If Player(attacker).GuildFileId = Player(mapnpc(mapnum).NPC(MapNPCNum).PetData.owner).GuildFileId Then
            Exit Function
        End If
    End If
    
    CanPlayerAttackPet = True
    
End Function

Public Function CanPetAttackPet(ByVal mapnum As Long, ByVal aMapNPCNum As Long, ByVal vMapNPCNum As Long) As Boolean
    CanPetAttackPet = False
    
    'Check if pet
    If mapnpc(mapnum).NPC(aMapNPCNum).IsPet = NO Then Exit Function
    If mapnpc(mapnum).NPC(vMapNPCNum).IsPet = NO Then Exit Function
    
    ' Check if map is attackable
    If Not Map(mapnum).Moral = MAP_MORAL_NONE Then
        If IsPlayerNeutral(mapnpc(mapnum).NPC(vMapNPCNum).PetData.owner) Then
            Exit Function
        End If
    End If

    ' Check to make sure that they dont have access
    If GetPlayerAccess(mapnpc(mapnum).NPC(aMapNPCNum).PetData.owner) > ADMIN_MONITOR Then
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(mapnpc(mapnum).NPC(vMapNPCNum).PetData.owner) > ADMIN_MONITOR Then
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(mapnpc(mapnum).NPC(aMapNPCNum).PetData.owner) < 10 Then
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(mapnpc(mapnum).NPC(vMapNPCNum).PetData.owner) < 10 Then
        Exit Function
    End If
    
    
    'make sure victim is not a guild partner
    If Player(mapnpc(mapnum).NPC(aMapNPCNum).PetData.owner).GuildFileId > 1 Then
        If Player(mapnpc(mapnum).NPC(aMapNPCNum).PetData.owner).GuildFileId = Player(mapnpc(mapnum).NPC(vMapNPCNum).PetData.owner).GuildFileId Then
        Exit Function
        End If
    End If
    
    CanPetAttackPet = True
    
End Function







