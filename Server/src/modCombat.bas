Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ## Basic Calculations ##
' ################################

Function GetPlayerMaxVital(ByVal index As Long, ByVal vital As Vitals) As Long
If index > MAX_PLAYERS Then Exit Function
Select Case vital
Case HP
'GetPlayerMaxVital = (GetPlayerLevel(index) * 10) + (GetPlayerStat(index, endurance) * 5 \ POINTS_PERLVL) + 50
    GetPlayerMaxVital = GetVitalByLevel(GetPlayerLevel(index))
    GetPlayerMaxVital = GetHPByTriforce(index, GetPlayerMaxVital)
    GetPlayerMaxVital = GetExtraHP(index, GetPlayerMaxVital)
Case MP
'GetPlayerMaxVital = (GetPlayerLevel(index) * 10) + (GetPlayerStat(index, Intelligence) * 5 \ POINTS_PERLVL) + 50
    GetPlayerMaxVital = GetVitalByLevel(GetPlayerLevel(index)) / 2 + GetMPByInt(GetPlayerStat(index, Intelligence))
    GetPlayerMaxVital = GetMPByTriforce(index, GetPlayerMaxVital)
End Select
End Function


Function GetExtraHP(ByVal index As Long, ByVal BaseVital As Long) As Long

    Dim N As Long

   If GetPlayerEquipment(index, Weapon) > 0 Then
        N = GetPlayerEquipment(index, Weapon)
        BaseVital = BaseVital + item(N).ExtraHP
    End If
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        N = GetPlayerEquipment(index, Armor)
        BaseVital = BaseVital + item(N).ExtraHP
    End If
    
    If GetPlayerEquipment(index, helmet) > 0 Then
        N = GetPlayerEquipment(index, helmet)
        BaseVital = BaseVital + item(N).ExtraHP
    End If
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        N = GetPlayerEquipment(index, Shield)
        BaseVital = BaseVital + item(N).ExtraHP
    End If
    
    GetExtraHP = BaseVital

End Function




Function GetHPByTriforce(ByVal index As Long, ByVal BaseVital As Long) As Long
Dim i As Byte
GetHPByTriforce = BaseVital
For i = 1 To TriforceType.TriforceType_Count - 1
    If GetPlayerTriforce(index, i) Then
        Select Case i
        Case TRIFORCE_COURAGE
            GetHPByTriforce = GetHPByTriforce + GetHPByTriforce \ 2
            'coraje tiene mas hp
        Case TRIFORCE_POWER
            GetHPByTriforce = GetHPByTriforce + GetHPByTriforce \ 4
        Case TRIFORCE_WISDOM
            GetHPByTriforce = GetHPByTriforce + GetHPByTriforce \ 6
            'sabiduria tiene menos
        End Select
    End If
Next

End Function

Function GetMPByTriforce(ByVal index As Long, ByVal BaseVital As Long) As Long
Dim i As Byte
GetMPByTriforce = BaseVital
For i = 1 To TriforceType.TriforceType_Count - 1
    If GetPlayerTriforce(index, i) Then
        Select Case i
        Case TRIFORCE_COURAGE
            GetMPByTriforce = GetMPByTriforce + GetMPByTriforce \ 6
            'coraje tiene menos mp
        Case TRIFORCE_POWER
            GetMPByTriforce = GetMPByTriforce + GetMPByTriforce \ 4
        Case TRIFORCE_WISDOM
            GetMPByTriforce = GetMPByTriforce + GetMPByTriforce \ 2
            'sabiduria tiene mas mp
        End Select
    End If
Next

End Function

Function GetVitalByLevel(ByVal level As Long) As Long
    Dim min As Long, max As Long
    min = 100
    max = 1000
    GetVitalByLevel = Line(MAX_LEVELS, 1, max, min, min, level)
End Function

Function GetMPByInt(ByVal i As Long) As Long
    Dim min As Long, max As Long
    min = 10
    max = 800
    GetMPByInt = Line(MAX_STAT, 0, max, min, min, i)
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
            i = GetPlayerMaxVital(index, HP) * GetVitalRegenPercent(GetPlayerStat(index, willpower))
        Case MP
            'i = (GetPlayerStat(index, Stats.willpower) / 4) + 60
            i = GetPlayerMaxVital(index, HP) * GetVitalRegenPercent(GetPlayerStat(index, willpower))
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

Function GetStunTime(ByVal attacker As Long, ByVal victim As Long) As Single
    Dim min As Single, max As Single
    min = 3#
    max = 1#
    
    GetStunTime = Line(MAX_LEVELS, 0, max, min, min, Abs(GetLevelDifference(attacker, victim)))

End Function

Function GetStunProbability(ByVal attacker As Long, ByVal victim As Long) As Integer
    Dim min1 As Integer, max1 As Integer, min2 As Integer, max2 As Integer
    min1 = 20
    max1 = 0
    
    min2 = 0
    max2 = 20
    
    Dim ChancesByLvl As Integer, ChancesByStrenght As Integer
    ChancesByLvl = (max1 - min1) / (MAX_LEVELS - 1) * Abs(GetLevelDifference(attacker, victim)) + min1

    ChancesByStrenght = (max2 - min2) / (MAX_STAT / 2) * GetStatDifference(attacker, victim, Strength, True) + min2

    
    GetStunProbability = ChancesByLvl + ChancesByStrenght
    If GetStunProbability < 0 Then
        GetStunProbability = 0
    End If
End Function

Sub CheckStunPlayer(ByVal attacker As Long, ByVal victim As Long)
    Dim Prob As Byte
    Prob = GetStunProbability(attacker, victim)
    If Prob > 0 Then
        If Prob >= RAND(1, 100) Then
            Dim Time As Single
            Time = GetStunTime(attacker, victim)
            StunPlayerByTime victim, Time
            ' map msg : Aturdido!
            SendActionMsg GetPlayerMap(victim), "Aturdido", Cyan, TARGET_TYPE_PLAYER, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32), , True
        End If
    End If
End Sub

Function GetStatDifference(ByVal index1 As Long, ByVal index2 As Long, ByVal stat As Stats, ByVal base As Boolean)
If base Then
    GetStatDifference = GetPlayerRawStat(index1, stat) - GetPlayerRawStat(index2, stat)
Else
    GetStatDifference = GetPlayerStat(index1, stat) - GetPlayerStat(index2, stat)
End If
End Function
Function GetPlayerDamageAgainstPlayer(ByVal attacker As Long, ByVal victim As Long) As Long
     Dim weaponNum As Long
    
    weaponNum = GetPlayerEquipment(attacker, Weapon)
    Dim WeaponDamage As Long
    If weaponNum = 0 Then
        WeaponDamage = 0
    Else
        WeaponDamage = item(weaponNum).Data1
    End If
    
    GetPlayerDamageAgainstPlayer = 4 * (GetPlayerStat(attacker, Strength) + WeaponDamage + GetPlayerLevel(attacker))
    Exit Function
    
End Function

Function GetPlayerDamageAgainstNPC(ByVal index As Long) As Long
    Dim weaponNum As Long
    
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(index, Weapon)
        GetPlayerDamageAgainstNPC = 4 * (GetPlayerStat(index, Strength) + item(weaponNum).Data2 + GetPlayerLevel(index))
    Else
        GetPlayerDamageAgainstNPC = 4 * (GetPlayerStat(index, Strength) + GetPlayerLevel(index))
    End If

End Function

Function GetPlayerDefenseAgainstNPC(ByVal index As Long, ByVal Damage As Long) As Long

Dim MinFactor As Single
Dim MaxFactor As Single
Dim factor As Single
    
MinFactor = 0.32
MaxFactor = 0.64
'LOS VALORES DE FACTOR OSCILAN ENTRE MINFACTOR Y MAXFACTOR, SIEMPRE ENTRE ESOS DOS VALORES
'OSCILAN SEGUN: STAT DEL NPC
factor = Line(MAX_STAT * 2, 0, MaxFactor, MinFactor, MinFactor, (GetPlayerStat(index, Endurance) + GetPlayerDef(index)))

'FUNCION DE DEFENSA, EL FACTOR DEBERIA OSCILAR ENTRE 0 Y 1
'0: EL DAÑO SE VE REDUCIDO AL MÍNIMO (NO SE REDUCE DAÑO BASE)
'1: EL DAÑO SE VE REDUCIDO AL MÁXIMO (NO HACE DAÑO)
GetPlayerDefenseAgainstNPC = Damage * factor

End Function

Function GetPlayerDefenseAgainstPlayer(ByVal index As Long, ByVal Damage As Long) As Long


Dim MinFactor As Single
Dim MaxFactor As Single
Dim factor As Single
    
MinFactor = 0.32
MaxFactor = 0.64

'LOS VALORES DE FACTOR OSCILAN ENTRE MINFACTOR Y MAXFACTOR, SIEMPRE ENTRE ESOS DOS VALORES
'OSCILAN SEGUN: STAT DEL NPC
factor = Line(MAX_STAT * 2, 0, MaxFactor, MinFactor, MinFactor, ((GetPlayerStat(index, Endurance) * 2) + GetPlayerDef(index)))

'FUNCION DE DEFENSA, EL FACTOR DEBERIA OSCILAR ENTRE 0 Y 1
'0: EL DAÑO SE VE REDUCIDO AL MÍNIMO (NO SE REDUCE DAÑO BASE)
'1: EL DAÑO SE VE REDUCIDO AL MÁXIMO (NO HACE DAÑO)

GetPlayerDefenseAgainstPlayer = Damage * factor

End Function

Function GetPlayerProjectileDamageAgainstNPC(ByVal index As Long, ByVal PlayerProjectile As Long) As Long
    If PlayerProjectile < 1 Then Exit Function
        '20 % more than Body PVN
    GetPlayerProjectileDamageAgainstNPC = 3.2 * (GetPlayerStat(index, Agility) + TempPlayer(index).ProjecTile(PlayerProjectile).Damage + GetPlayerLevel(index))
End Function

Function GetPlayerProjectileDamageAgainstPlayer(ByVal attacker As Long, ByVal victim As Long) As Long
    Dim weaponNum As Long
    
    weaponNum = GetPlayerEquipment(attacker, Weapon)
    
    Dim WeaponDamage As Long
    If weaponNum = 0 Then
        WeaponDamage = 0
    Else
        WeaponDamage = item(weaponNum).ProjecTile.Damage
    End If
    
    GetPlayerProjectileDamageAgainstPlayer = 3.2 * (WeaponDamage + GetPlayerStat(attacker, Agility) + GetPlayerLevel(attacker))
    
    
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
        Def = Def + item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, helmet) > 0 Then
        DefNum = GetPlayerEquipment(index, helmet)
        Def = Def + item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        DefNum = GetPlayerEquipment(index, Shield)
        Def = Def + item(DefNum).Data2
    End If
    
    GetPlayerDef = Def

End Function

Function GetNpcMaxVital(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal vital As Vitals) As Long
    Dim X As Long

    Dim PetOwner As Long, npcnum As Long
    'Prevent Pet System
    PetOwner = GetMapPetOwner(MapNum, mapnpcnum)
    npcnum = GetNPCNum(MapNum, mapnpcnum)
    If npcnum = 0 Then Exit Function
    
    If npcnum > MAX_NPCS Then
        MsgBox "Panic. getNpcMaxVital, npcnum over max npcs."
        End
    'GetNpcMaxVital = 0
    'Exit Function
    End If
    
    
    Select Case PetOwner
    
    Case Is > 0
        Select Case vital
            Case HP
                GetNpcMaxVital = NPC(npcnum).HP + ((GetNpcLevel(npcnum, PetOwner) / 2) + (GetNpcStat(MapNum, mapnpcnum, Endurance) / 2) * 10)
            Case MP
                GetNpcMaxVital = 30 + ((GetNpcLevel(npcnum, PetOwner) / 2) + (GetNpcStat(MapNum, mapnpcnum, Intelligence) / 2)) * 10
            End Select
    Case Else
            Select Case vital
            Case HP
                GetNpcMaxVital = NPC(npcnum).HP
            Case MP
                GetNpcMaxVital = 30 + (GetNPCBaseStat(npcnum, Intelligence) * 10) + 2
            End Select
    End Select

End Function

Function GetNpcVitalRegen(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal vital As Vitals) As Long
    Dim i As Long

    
    Dim PetOwner As Long
    'Prevent Pet System
    PetOwner = GetMapPetOwner(MapNum, mapnpcnum)
    
    Select Case PetOwner
    
    Case Is > 0
        i = GetNpcMaxVital(MapNum, mapnpcnum, vital) * GetVitalRegenPercent(GetNpcStat(MapNum, mapnpcnum, willpower, False))
    Case 0
    
        Select Case vital
            Case HP
                i = (GetNPCBaseStat(GetNPCNum(MapNum, mapnpcnum), willpower) * 0.8) + 6
            Case MP
                i = (GetNPCBaseStat(GetNPCNum(MapNum, mapnpcnum), willpower) / 4) + 12.5
        End Select
    
    End Select
    
    GetNpcVitalRegen = i

End Function

Function GetNpcDamage(ByVal MapNum As Long, ByVal mapnpcnum As Long) As Long
    Dim MinFactor As Single
    Dim MaxFactor As Single
    Dim npcnum As Long
    
    npcnum = GetNPCNum(MapNum, mapnpcnum)
    If npcnum < 1 Or npcnum > MAX_NPCS Then Exit Function
    
    If GetMapPetOwner(MapNum, mapnpcnum) > 0 Then
        MinFactor = 2
        MaxFactor = 3
    Else
        MinFactor = 1
        MaxFactor = 2
    End If
    Dim factor As Double
    factor = ((MaxFactor - MinFactor) / MAX_STAT * GetNpcStat(MapNum, mapnpcnum, Strength) + MinFactor)
    
    
    GetNpcDamage = NPC(npcnum).Damage * factor
End Function

Function GetNpcDefense(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal BaseDamage As Long) As Long
    Dim MinFactor As Single
    Dim MaxFactor As Single
    
    If GetMapPetOwner(MapNum, mapnpcnum) > 0 Then
        MinFactor = 0.4
        MaxFactor = 0.8
    Else
        MinFactor = 0.2
        MaxFactor = 0.5
    End If
    
    Dim factor As Double
    factor = Line(MAX_STAT, 0, MaxFactor, MinFactor, MinFactor, GetNpcStat(MapNum, mapnpcnum, Endurance))
    GetNpcDefense = BaseDamage * factor
    
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

    rate = GetPlayerStat(index, Agility) * 0.5
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
    
    If rate > 100 Then
        rate = 100
    End If
    
    If rate <= 0 Then
        rate = 0
    Else
        rate = rate * 0.3
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

Public Function CanNpcBlock(ByVal npcnum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcBlock = False

    rate = 0
    ' TODO : make it based on shield lol
End Function

Public Function CanNpcCrit(ByVal npcnum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = GetNPCBaseStat(npcnum, Strength) * 0.09
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal index As Long, ByVal npcnum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    If (GetNPCBaseStat(npcnum, Agility) < GetPlayerRawStat(index, Agility)) Then
        rate = 0
    Else
        rate = (GetNPCBaseStat(npcnum, Agility) - GetPlayerRawStat(index, Agility)) * 0.07
    End If
    
    
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal index As Long, ByVal npcnum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    If (GetNPCBaseStat(npcnum, Strength) < player(index).stat(Stats.Strength)) Then
        rate = 0
    Else
        rate = (GetNPCBaseStat(npcnum, Strength) - player(index).stat(Stats.Strength)) * 0.07
    End If
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal index As Long, ByVal mapnpcnum As Long)
Dim blockAmount As Long
Dim npcnum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(index, mapnpcnum) Then
    
        MapNum = GetPlayerMap(index)
        npcnum = MapNpc(MapNum).NPC(mapnpcnum).Num
        
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(index, npcnum) Then
            SendActionMsg MapNum, "Esquivado!", Pink, 1, (MapNpc(MapNum).NPC(mapnpcnum).X * 32), (MapNpc(MapNum).NPC(mapnpcnum).Y * 32), , True
            Exit Sub
        End If
        'Dim StunTime As Long
        'StunTime = CanStunNpc(index, mapnum, mapnpcnum)
        'If StunTime > 0 Then
            'Call StunNPCByTime(mapnum, mapnpcnum, StunTime)
            'SendActionMsg mapnum, "Aturdido!", Pink, 1, (mapnpc(mapnum).NPC(mapnpcnum).X * 32), (mapnpc(mapnum).NPC(mapnpcnum).Y * 32)
            'Exit Sub
        'End If

        ' Get the damage we can do
        Damage = GetPlayerDamageAgainstNPC(index)
        
        Damage = Damage - GetNpcDefense(MapNum, mapnpcnum, Damage)
        ' take away armour
        ' randomise from half to max hit
        Damage = RAND(Damage / 2, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCriticalHit(index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "¡Crítico!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32), , True
            SendSoundToMap MapNum, GetPlayerX(index), GetPlayerY(index), SoundEntity.seCritical, GetPlayerClass(index)
        Else
            SendSoundToMap MapNum, GetPlayerX(index), GetPlayerY(index), SoundEntity.seAttack, GetPlayerClass(index)
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(index, mapnpcnum, Damage)
        Else
            Call PlayerMsg(index, "Tu ataque no hace nada.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal mapnpcnum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim npcnum As Long
    Dim NPCX As Long
    Dim NPCY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(attacker)).NPC(mapnpcnum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(attacker)
    npcnum = MapNpc(MapNum).NPC(mapnpcnum).Num
    
    'Pet check
    If IsMapNPCaPet(MapNum, mapnpcnum) Then
        If Not (CanPlayerAttackPet(MapNum, mapnpcnum, attacker)) Then
            Exit Function
        End If
    End If
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    'Can't attack own pet
    If TempPlayer(attacker).TempPet.TempPetSlot = mapnpcnum Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
    
        ' exit out early
        If IsSpell Then
             If npcnum > 0 Then
                If CanNPCBeAttacked(npcnum) Then
                    TempPlayer(attacker).TargetType = TARGET_TYPE_NPC
                    TempPlayer(attacker).Target = mapnpcnum
                    SendTarget attacker
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            attackspeed = item(GetPlayerEquipment(attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If npcnum > 0 And GetRealTickCount > TempPlayer(attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(attacker)
                Case DIR_UP
                    NPCX = MapNpc(MapNum).NPC(mapnpcnum).X
                    NPCY = MapNpc(MapNum).NPC(mapnpcnum).Y + 1
                Case DIR_DOWN
                    NPCX = MapNpc(MapNum).NPC(mapnpcnum).X
                    NPCY = MapNpc(MapNum).NPC(mapnpcnum).Y - 1
                Case DIR_LEFT
                    NPCX = MapNpc(MapNum).NPC(mapnpcnum).X + 1
                    NPCY = MapNpc(MapNum).NPC(mapnpcnum).Y
                Case DIR_RIGHT
                    NPCX = MapNpc(MapNum).NPC(mapnpcnum).X - 1
                    NPCY = MapNpc(MapNum).NPC(mapnpcnum).Y
            End Select

            If NPCX = GetPlayerX(attacker) Then
                If NPCY = GetPlayerY(attacker) Then
                    If CanNPCBeAttacked(npcnum) Then
                        CanPlayerAttackNpc = True
                    Else
                        'ALATAR
                        If NPC(npcnum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(npcnum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                            
                            If Len(Trim$(NPC(npcnum).AttackSay)) > 0 Then
                                PlayerMsg attacker, Trim$(NPC(npcnum).TranslatedName) & ": " & GetTranslation((NPC(npcnum).AttackSay)), White, , False
                                'Call SendActionMsg(mapnum, Trim$(NPC(npcnum).Name) & ": " & Trim$(NPC(npcnum).AttackSay), SayColor, 1, mapnpc(mapnum).NPC(mapnpcnum).X * 32, mapnpc(mapnum).NPC(mapnpcnum).Y * 32)
                                Call SpeechWindow(attacker, GetTranslation(NPC(npcnum).AttackSay), npcnum)
                            End If
                            
                            SendMapSound (attacker), GetPlayerX(attacker), GetPlayerY(attacker), SoundEntity.seNpc, MapNpc(MapNum).NPC(mapnpcnum).Num
                            
                            Call CheckTasks(attacker, QUEST_TYPE_GOTALK, npcnum)
                            Call CheckTasks(attacker, QUEST_TYPE_GOGIVE, npcnum)
                            Call CheckTasks(attacker, QUEST_TYPE_GOGET, npcnum)
                            
                            If NPC(npcnum).Quest = YES Then 'Alatar v1.2
                                If player(attacker).PlayerQuest(NPC(npcnum).Quest).Status = QUEST_COMPLETED Then
                                    If Quest(NPC(npcnum).Quest).Repeat = YES Then
                                        player(attacker).PlayerQuest(NPC(npcnum).Quest).Status = QUEST_COMPLETED_BUT
                                        Exit Function
                                    End If
                                End If
                                If CanStartQuest(attacker, NPC(npcnum).questnum) Then
                                    'if can start show the request message (speech1)
                                    QuestMessage attacker, NPC(npcnum).questnum, GetTranslation(Quest(NPC(npcnum).questnum).Speech(1)), NPC(npcnum).questnum
                                    Exit Function
                                End If
                                If QuestInProgress(attacker, NPC(npcnum).questnum) Then
                                    'if the quest is in progress show the meanwhile message (speech2)
                                    QuestMessage attacker, NPC(npcnum).questnum, GetTranslation(Quest(NPC(npcnum).questnum).Speech(2)), 0
                                    Exit Function
                                End If
                            End If
                        End If
                        '/ALATAR
                    End If
                End If
            End If
        End If
    End If

End Function


Public Sub PlayerAttackNpc(ByVal attacker As Long, ByVal mapnpcnum As Long, ByVal Damage As Long, Optional ByVal spellnum As Long, Optional ByVal overTime As Boolean = False)
    'Dim Name As String
    Dim exp As Long
    Dim N As Long
    Dim i As Long
    Dim DropNum As Integer
    Dim STR As Long
    Dim Def As Long
    Dim MapNum As Long
    Dim npcnum As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(attacker)
    npcnum = MapNpc(MapNum).NPC(mapnpcnum).Num
    If npcnum < 1 Then Exit Sub
    'Name = Trim$(NPC(npcnum).Name)
    
    ' Check for weapon
    N = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        N = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetRealTickCount

    If Damage >= MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.HP) Then
        SendActionMsg GetPlayerMap(attacker), "-" & MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.HP), BrightRed, 1, (MapNpc(MapNum).NPC(mapnpcnum).X * 32), (MapNpc(MapNum).NPC(mapnpcnum).Y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y
        
        'Kill counter
        player(attacker).NpcKill = player(attacker).NpcKill + 1
        
        ' send the sound
        If spellnum > 0 Then SendMapSound attacker, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, SoundEntity.seSpell, spellnum
        
        ' send animation
            If N > 0 Then
                If Not overTime Then
                    If spellnum = 0 Then
                        Call SendAnimation(MapNum, item(GetPlayerEquipment(attacker, Weapon)).Animation, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y)
                    Else
                        Call SendAnimation(MapNum, Spell(spellnum).SpellAnim, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y)
                    End If
                End If
            End If

        'Pet ?
        Call ComputePlayerExp(attacker, TARGET_TYPE_PLAYER, mapnpcnum, TARGET_TYPE_NPC)
        
        'Auto Targetting
        If TempPlayer(attacker).TempPet.PetHasOwnTarget = mapnpcnum Then
            'Objective Finished
            TempPlayer(attacker).TempPet.PetHasOwnTarget = 0
        End If
        
        'begin of the new system
        'Call the calculator for deciding which item has to be chanced
        DropNum = CalculateDropChances(npcnum)
        'Drop the goods if they get it
        If DropNum > 0 Then
            N = Int(Rnd * NPC(npcnum).Drops(DropNum).DropChance) + 1

            'Drop the selected item
            If N = 1 Then
                Call SpawnItem(NPC(npcnum).Drops(DropNum).DropItem, NPC(npcnum).Drops(DropNum).DropItemValue, MapNum, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y)
            End If
        End If
        
        'If NPC was random spawned, disapear and respawn it
        'If IsTempNPC(mapnum, MapNPCNum) Then
            'Call ClearSingleMapNpc(MapNPCNum, mapnum)
            'Call SendMapNpcToMap(mapnum, MapNPCNum)
            'If NPCNum = NPC_SKULLTULA Then
                'Call RespawnRandomNPC(NPCNum)
            'End If
        'End If
        
        Call KillNpc(MapNum, mapnpcnum)
        
        Call CheckPlayerPartyTasks(attacker, QUEST_TYPE_GOSLAY, npcnum)
        'ALATAR
        'Call CheckTasks(attacker, QUEST_TYPE_GOSLAY, NPCNum)
        '/ALATAR
        
        'Player NPC info
        player(attacker).NPCKills = player(attacker).NPCKills + 1
        
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.HP) = MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.HP) - Damage
        
        'target the NPC
        TempPlayer(attacker).TargetType = TARGET_TYPE_NPC
        TempPlayer(attacker).Target = mapnpcnum
        SendTarget attacker
        
        ' Check for a weapon and say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).NPC(mapnpcnum).X * 32), (MapNpc(MapNum).NPC(mapnpcnum).Y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound attacker, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, SoundEntity.seSpell, spellnum
        
        ' send animation
        If N > 0 Then
            If Not overTime Then
                If spellnum = 0 Then Call SendAnimation(MapNum, item(GetPlayerEquipment(attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapnpcnum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(MapNum).NPC(mapnpcnum).TargetType = 1 ' player
        MapNpc(MapNum).NPC(mapnpcnum).Target = attacker
        
        'Set the NPC target to player's pet target
        If TempPlayer(attacker).TempPet.TempPetSlot > 0 And TempPlayer(attacker).TempPet.TempPetSlot < MAX_MAP_NPCS And TempPlayer(attacker).TempPet.PetHasOwnTarget = 0 Then
            MapNpc(MapNum).NPC(TempPlayer(attacker).TempPet.TempPetSlot).TargetType = TARGET_TYPE_NPC
            MapNpc(MapNum).NPC(TempPlayer(attacker).TempPet.TempPetSlot).Target = mapnpcnum
            'Auto Targetting
            TempPlayer(attacker).TempPet.PetHasOwnTarget = mapnpcnum
        End If
            

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(MapNum).NPC(mapnpcnum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).NPC(i).Num = MapNpc(MapNum).NPC(mapnpcnum).Num Then
                    MapNpc(MapNum).NPC(i).Target = attacker
                    MapNpc(MapNum).NPC(i).TargetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(MapNum).NPC(mapnpcnum).stopRegen = True
        MapNpc(MapNum).NPC(mapnpcnum).stopRegenTimer = GetRealTickCount
        
        ' if stunning spell, stun the npc
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunNPC mapnpcnum, MapNum, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Npc MapNum, mapnpcnum, spellnum, attacker
            End If
        End If
        
        
     
        
    
        SendMapNpcVitals MapNum, mapnpcnum
    End If

    If spellnum = 0 And Not overTime Then
        ' Reset attack timer
        TempPlayer(attacker).AttackTimer = GetRealTickCount
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal mapnpcnum As Long, ByVal index As Long)
Dim MapNum As Long, npcnum As Long, blockAmount As Long, Damage As Long
Dim buffer As clsBuffer
    ' Can the npc attack the player?
    If CanNpcAttackPlayer(mapnpcnum, index) Then
        MapNum = GetPlayerMap(index)
        npcnum = MapNpc(MapNum).NPC(mapnpcnum).Num
        
        
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seNpc, MapNpc(MapNum).NPC(mapnpcnum).Num
        
        ' Send this packet so they can see the npc attacking
        Set buffer = New clsBuffer
        buffer.WriteLong ServerPackets.SNpcAttack
        buffer.WriteLong mapnpcnum
        SendDataToMap MapNum, buffer.ToArray()
        Set buffer = Nothing
        
        MapNpc(MapNum).NPC(mapnpcnum).AttackTimer = GetRealTickCount + GetNPCAttackTimer(MapNum, mapnpcnum)
        
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg MapNum, "¡Esquivado!", Pink, 1, (player(index).X * 32), (player(index).Y * 32), , True
            Exit Sub
        End If
        If CanPlayerParry(index) Then
            SendActionMsg MapNum, "¡Bloqueado!", Pink, 1, (player(index).X * 32), (player(index).Y * 32), , True
            Exit Sub
        End If


        ' Get the damage we can do
        Damage = GetNpcDamage(MapNum, mapnpcnum)
        Damage = Damage - GetPlayerDefenseAgainstNPC(index, Damage)
        Damage = RAND(Damage * 0.8, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "¡Crítico!", BrightCyan, 1, (MapNpc(MapNum).NPC(mapnpcnum).X * 32), (MapNpc(MapNum).NPC(mapnpcnum).Y * 32), , True
        End If

        'Damage = Damage / (GetPlayerDef(index) / 20)
        
        'Damage = Damage * 0.8

        If Damage > 0 Then
            Call NpcAttackPlayer(mapnpcnum, index, Damage)
        Else
            SendActionMsg MapNum, "¡Evitado!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32), , True
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal mapnpcnum As Long, ByVal index As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim npcnum As Long
    Dim buffer As clsBuffer
    
    ' Check if player is loading
    If TempPlayer(index).IsLoading = True Then Exit Function
    
    ' Check for subscript out of range
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).NPC(mapnpcnum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(index)
    npcnum = MapNpc(MapNum).NPC(mapnpcnum).Num
    
    'Pet check
    If IsMapNPCaPet(MapNum, mapnpcnum) Then
        If Not (CanPetAttackPlayer(MapNum, mapnpcnum, index)) Then
            Exit Function
        End If
    Else
        If (Not CanNPCBehaviourAttack(npcnum)) Then Exit Function
    End If

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    

    

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If
    
    'check if the NPC attacking us is actually our pet.
    'We don't want a rebellion on our hands now do we?
    If MapNpc(MapNum).NPC(mapnpcnum).PetData.Owner = index Then Exit Function
    
    'Spell Check
    If IsSpell Then
        CanNpcAttackPlayer = True
    Else
        ' Make sure npcs dont attack more then once a second
        If GetRealTickCount < MapNpc(MapNum).NPC(mapnpcnum).AttackTimer Then
            Exit Function
        End If
        
        ' Make sure they are on the same map
        If IsPlaying(index) Then
            If npcnum > 0 Then

                ' Check if at same coordinates
                If (GetPlayerY(index) + 1 = MapNpc(MapNum).NPC(mapnpcnum).Y) And (GetPlayerX(index) = MapNpc(MapNum).NPC(mapnpcnum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) - 1 = MapNpc(MapNum).NPC(mapnpcnum).Y) And (GetPlayerX(index) = MapNpc(MapNum).NPC(mapnpcnum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(MapNum).NPC(mapnpcnum).Y) And (GetPlayerX(index) + 1 = MapNpc(MapNum).NPC(mapnpcnum).X) Then
                            CanNpcAttackPlayer = True
                        Else
                            If (GetPlayerY(index) = MapNpc(MapNum).NPC(mapnpcnum).Y) And (GetPlayerX(index) - 1 = MapNpc(MapNum).NPC(mapnpcnum).X) Then
                                CanNpcAttackPlayer = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
    End If

    
End Function


Sub NpcAttackPlayer(ByVal mapnpcnum As Long, ByVal victim As Long, ByVal Damage As Long)
    'Dim Name As String
    Dim exp As Long
    Dim MapNum As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(victim)).NPC(mapnpcnum).Num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(victim)
    'Name = (NPC(MapNpc(mapnum).NPC(mapnpcnum).Num).Name)
    
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
    MapNpc(MapNum).NPC(mapnpcnum).stopRegen = True
    MapNpc(MapNum).NPC(mapnpcnum).stopRegenTimer = GetRealTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(MapNum).NPC(mapnpcnum).Num
        
        'Drop Items If npc was a  pet
        If GetMapPetOwner(MapNum, mapnpcnum) > 0 Then
            If map(MapNum).moral <> MAP_MORAL_ARENA Then
                If Not (GetLevelDifference(GetMapPetOwner(MapNum, mapnpcnum), victim) > 20) Then
                    Call PlayerPVPDrops(victim)
                End If
                
                Call SetPlayerJustice(GetMapPetOwner(MapNum, mapnpcnum), victim)
                Call ComputeArmyPvP(GetMapPetOwner(MapNum, mapnpcnum), victim)
            End If
        End If
        
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        'Call GlobalMsg(GetPlayerName(victim) & " ha sido asesinado por " & Name, BrightRed)
        
        'Kill Counter
        player(victim).NpcDead = player(victim).NpcDead + 1

        ' Set NPC target to 0
        MapNpc(MapNum).NPC(mapnpcnum).Target = 0
        MapNpc(MapNum).NPC(mapnpcnum).TargetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        Call SendAnimation(MapNum, NPC(MapNpc(GetPlayerMap(victim)).NPC(mapnpcnum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, victim)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(MapNum).NPC(mapnpcnum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & Damage, Yellow, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetRealTickCount
        
        SendSoundToMap GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seHit, GetPlayerClass(victim)
    End If

End Sub

Sub PetSpellItself(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal SpellSlotNum As Long)
    Dim spellnum As Long
    Dim InitDamage As Long
    Dim DidCast As Boolean
    Dim MPCost As Long
    
    ' Check for subscript out of range
        If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Then
                Exit Sub
        End If

        ' Check for subscript out of range
        If MapNpc(MapNum).NPC(mapnpcnum).Num <= 0 Then
                Exit Sub
        End If
   
        If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub
        
        ' The Variables
        spellnum = NPC(MapNpc(MapNum).NPC(mapnpcnum).Num).Spell(SpellSlotNum)
        
        MPCost = GetSpellMPCost(MapNum, mapnpcnum, TARGET_TYPE_NPC, spellnum)

        ' Check if they have enough MP
        If MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) < MPCost Then
            Exit Sub
        End If
        
        If spellnum = 0 Then
        Exit Sub
        End If
        
        DidCast = False
   
        ' CoolDown Time
        If MapNpc(MapNum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) > GetRealTickCount Then Exit Sub
   
        ' Spell Types
                Select Case Spell(spellnum).Type
                    Case SPELL_TYPE_HEALHP
                        If Not Spell(spellnum).IsAoE Then
                                InitDamage = GetNPCSpellDamage(spellnum, MapNum, mapnpcnum)
                                SpellNpc_Effect Vitals.HP, True, mapnpcnum, InitDamage, spellnum, MapNum
                                DidCast = True
                        End If
                    Case SPELL_TYPE_HEALMP
                        If Not Spell(spellnum).IsAoE Then
                                InitDamage = GetNPCSpellDamage(spellnum, MapNum, mapnpcnum)
                                SpellNpc_Effect Vitals.MP, True, mapnpcnum, InitDamage, spellnum, MapNum
                                DidCast = True
                       End If
                    End Select
                    
                    If DidCast Then
                        MapNpc(MapNum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                        SendNpcAttackAnimation MapNum, mapnpcnum
                        If MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) - MPCost < 0 Then
                            MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) = 0
                        Else
                            MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) = MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) - MPCost
                        End If
                        SendMapNpcVitals MapNum, mapnpcnum
                    Else
                        MapNpc(MapNum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                    End If
End Sub

Sub PetSpellOwner(ByVal mapnpcnum As Long, ByVal index As Long, SpellSlotNum As Long)
    Dim MapNum As Long
    Dim spellnum As Long
    Dim InitDamage As Long
    Dim MPCost As Long
    Dim DidCast As Boolean
    
    ' Check for subscript out of range
        If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or IsPlaying(index) = False Then
                Exit Sub
        End If

        ' Check for subscript out of range
        If MapNpc(GetPlayerMap(index)).NPC(mapnpcnum).Num <= 0 Then
                Exit Sub
        End If
   
        If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub
        
        ' The Variables
        MapNum = GetPlayerMap(index)
        spellnum = NPC(MapNpc(MapNum).NPC(mapnpcnum).Num).Spell(SpellSlotNum)
        
        MPCost = GetSpellMPCost(MapNum, mapnpcnum, TARGET_TYPE_NPC, spellnum)

        ' Check if they have enough MP
        If MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) < MPCost Then
            Exit Sub
        End If
        
        If spellnum = 0 Then
        Exit Sub
        End If
        
        DidCast = False
   
        ' CoolDown Time
        If MapNpc(MapNum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) > GetRealTickCount Then Exit Sub
   
        ' Spell Types
                Select Case Spell(spellnum).Type
                    Case SPELL_TYPE_HEALHP
                        If Not Spell(spellnum).IsAoE Then
                            If IsinRange(Spell(spellnum).range + 3, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, GetPlayerX(index), GetPlayerY(index)) Then
                                InitDamage = GetNPCSpellDamage(spellnum, MapNum, mapnpcnum)
                                SpellPlayer_Effect Vitals.HP, True, index, InitDamage, spellnum
                                DidCast = True
                            End If
                        End If
                    Case SPELL_TYPE_HEALMP
                        If Not Spell(spellnum).IsAoE Then
                            If IsinRange(Spell(spellnum).range + 3, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, GetPlayerX(index), GetPlayerY(index)) Then
                                InitDamage = GetNPCSpellDamage(spellnum, MapNum, mapnpcnum)
                                SpellPlayer_Effect Vitals.MP, True, index, InitDamage, spellnum
                                DidCast = True
                            End If
                        End If
                    End Select
                    
                    If DidCast Then
                        MapNpc(MapNum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                        SendNpcAttackAnimation MapNum, mapnpcnum
                        If MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) - MPCost < 0 Then
                            MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) = 0
                        Else
                            MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) = MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) - MPCost
                        End If
                        SendMapNpcVitals MapNum, mapnpcnum
                    Else
                        MapNpc(MapNum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                    End If
End Sub
Sub NpcSpellPlayer(ByVal mapnpcnum As Long, ByVal victim As Long, SpellSlotNum As Long)
        Dim MapNum As Long
        Dim i As Long
        Dim N As Long
        Dim spellnum As Long
        Dim buffer As clsBuffer
        Dim InitDamage As Long
        Dim Damage As Long
        Dim index As Long
        Dim PetOwner As Long
        Dim MPCost As Long
        Dim DidCast As Boolean

        ' Check for subscript out of range
        If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
                Exit Sub
        End If

        ' Check for subscript out of range
        If MapNpc(GetPlayerMap(victim)).NPC(mapnpcnum).Num <= 0 Then
                Exit Sub
        End If
   
        If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub
        
        ' The Variables
        MapNum = GetPlayerMap(victim)
        spellnum = NPC(MapNpc(MapNum).NPC(mapnpcnum).Num).Spell(SpellSlotNum)
        
        If spellnum = 0 Then
        Exit Sub
        End If
        
        MPCost = GetSpellMPCost(MapNum, mapnpcnum, TARGET_TYPE_NPC, spellnum)

        ' Check if they have enough MP
        If MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) < MPCost Then
            Exit Sub
        End If
        
   
        ' CoolDown Time
        If MapNpc(MapNum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) > GetRealTickCount Then Exit Sub
   
        ' Spell Types
                Select Case Spell(spellnum).Type
                        ' AOE Healing Spells
                        Case SPELL_TYPE_HEALHP
                        
                        ' Make sure an npc waits for the spell to cooldown
                            If Spell(spellnum).IsAoE Then
                                'For i = 1 To MAX_PLAYERS
                                                   
                                                
                                'Next
                            Else
                                If Not HasNPCMaxVital(HP, MapNum, mapnpcnum) Then
                                    ' Non AOE Healing Spells
                                    InitDamage = GetNPCSpellDamage(spellnum, MapNum, mapnpcnum)
                                    SpellNpc_Effect Vitals.HP, True, mapnpcnum, InitDamage, spellnum, MapNum
                                    DidCast = True
                                End If
                            End If
                                
                           
                        ' AOE Damaging Spells
                        Case SPELL_TYPE_DAMAGEHP
                        ' Make sure an npc waits for the spell to cooldown
                                If Spell(spellnum).IsAoE Then
                                        For i = 1 To Player_HighIndex
                                            If GetPlayerMap(i) = MapNum Then
                                                If CanNpcAttackPlayer(mapnpcnum, i, True) Then
                                                        If IsinRange(Spell(spellnum).AoE, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, GetPlayerX(i), GetPlayerY(i)) Then
                                                                InitDamage = GetSpellDamage(MapNum, mapnpcnum, TARGET_TYPE_NPC, i, TARGET_TYPE_PLAYER, spellnum)
                                                                Damage = InitDamage - player(i).stat(Stats.willpower)
                                                                If Damage <= 0 Then
                                                                    SendActionMsg GetPlayerMap(i), "RESISTISTE!", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32), , True
                                                                Else
                                                                    If Spell(spellnum).StunDuration > 0 Then
                                                                        CheckSpellStunts victim, spellnum
                                                                    End If
                                                                    SendAnimation MapNum, Spell(spellnum).SpellAnim, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
                                                                    NpcAttackPlayer mapnpcnum, i, Damage
                                                                    DidCast = True
                                                                End If
                                                        End If
                                                End If
                                            End If
                                        Next
                                ' Non AoE Damaging Spells
                                Else
                                    If CanNpcAttackPlayer(mapnpcnum, victim, True) Then
                                        If IsinRange(Spell(spellnum).range, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, GetPlayerX(victim), GetPlayerY(victim)) Then
                                        InitDamage = GetSpellDamage(MapNum, mapnpcnum, TARGET_TYPE_NPC, victim, TARGET_TYPE_PLAYER, spellnum)
                                        Damage = InitDamage
                                                If Damage <= 0 Then
                                                        SendActionMsg GetPlayerMap(victim), "RESISTISTE!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32), , True
                                                Else
                                                    If Spell(spellnum).StunDuration > 0 Then
                                                        CheckSpellStunts victim, spellnum
                                                    End If
                                                    NpcAttackPlayer mapnpcnum, victim, Damage - player(victim).stat(Stats.willpower)
                                                    SendAnimation MapNum, Spell(spellnum).SpellAnim, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
                                                    DidCast = True
                                                End If
                                        End If
                                    End If
                                End If

                                Case SPELL_TYPE_DAMAGEMP
                                    ' Make sure an npc waits for the spell to cooldown
                                    If Spell(spellnum).IsAoE Then
                                        For i = 1 To Player_HighIndex
                                            If GetPlayerMap(i) = MapNum Then
                                                If CanNpcAttackPlayer(mapnpcnum, victim, True) Then
                                                    If IsinRange(Spell(spellnum).AoE, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, GetPlayerX(i), GetPlayerY(i)) Then
                    
                                                        Damage = GetSpellDamage(MapNum, mapnpcnum, TARGET_TYPE_NPC, i, TARGET_TYPE_PLAYER, spellnum)
                                                        If Damage <= 0 Then
                                                            SendActionMsg GetPlayerMap(i), "¡RESISTISTE!", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32), , True
                                                        Else
                                                            SpellPlayer_Effect Vitals.MP, False, victim, Damage - player(victim).stat(Stats.willpower), spellnum
                                                            DidCast = True
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Next
                                    ' Non AoE DamagingMP Spells
                                    Else
                                        If IsinRange(Spell(spellnum).range, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, GetPlayerX(victim), GetPlayerY(victim)) Then
                                            If CanNpcAttackPlayer(mapnpcnum, victim, True) Then
                                                Damage = GetSpellDamage(MapNum, mapnpcnum, TARGET_TYPE_NPC, victim, TARGET_TYPE_PLAYER, spellnum)
                                                If Damage <= 0 Then
                                                    SendActionMsg GetPlayerMap(victim), "¡RESISTISTE!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32), , True
                                                Else
                                                    SpellPlayer_Effect Vitals.MP, False, victim, Damage - player(victim).stat(Stats.willpower), spellnum
                                                    DidCast = True
                                                End If
                                            End If
                                        End If
                                    End If
                                    
                                    Case SPELL_TYPE_HEALMP
                                        ' Make sure an npc waits for the spell to cooldown
                                        If Spell(spellnum).IsAoE Then
                                            'For i = 1 To MAX_PLAYERS
                                                               
                                                            
                                            'Next
                                        Else
                                            If Not HasNPCMaxVital(MP, MapNum, mapnpcnum) Then
                                                ' Non AOE Healing Spells
                                                InitDamage = GetNPCSpellDamage(spellnum, MapNum, mapnpcnum)
                                                SpellNpc_Effect Vitals.MP, True, mapnpcnum, InitDamage, spellnum, MapNum
                                                DidCast = True
                                            End If
                                        End If
                                        
                                    
                                    
                                    End Select
                        
                                    If DidCast Then
                                        MapNpc(MapNum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                                        SendNpcAttackAnimation MapNum, mapnpcnum
                                        If MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) - MPCost < 0 Then
                                            MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) = 0
                                        Else
                                            MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) = MapNpc(MapNum).NPC(mapnpcnum).vital(Vitals.MP) - MPCost
                                        End If
                                        SendMapNpcVitals MapNum, mapnpcnum
                                    Else
                                        MapNpc(MapNum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                                    End If

End Sub

Sub NpcSpellNpc(ByVal MapNum As Long, ByVal aMapNPCNum As Long, ByVal vMapNPCNum As Long, SpellSlotNum As Long)
        Dim i As Long
        Dim N As Long
        Dim spellnum As Long
        Dim buffer As clsBuffer
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
        If MapNpc(MapNum).NPC(aMapNPCNum).Num <= 0 Or MapNpc(MapNum).NPC(vMapNPCNum).Num <= 0 Then
                Exit Sub
        End If
   
        If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub

        ' The Variables
        spellnum = NPC(MapNpc(MapNum).NPC(aMapNPCNum).Num).Spell(SpellSlotNum)
        
        If spellnum = 0 Then
        Exit Sub
        End If
        
        MPCost = GetSpellMPCost(MapNum, aMapNPCNum, TARGET_TYPE_NPC, spellnum)

        ' Check if they have enough MP
        If MapNpc(MapNum).NPC(aMapNPCNum).vital(Vitals.MP) < MPCost Then
            Exit Sub
        End If
        
        'set cast to false
        DidCast = False
   
        ' CoolDown Time
        If MapNpc(MapNum).NPC(aMapNPCNum).SpellTimer(SpellSlotNum) > GetRealTickCount Then Exit Sub
   
        ' Spell Types
                Select Case Spell(spellnum).Type
                        ' AOE Healing Spells
                        Case SPELL_TYPE_HEALHP
                            
                        ' Make sure an npc waits for the spell to cooldown
                            If Spell(spellnum).IsAoE Then
                                For i = 1 To MAX_MAP_NPCS
                                    If MapNpc(MapNum).NPC(i).Num > 0 Then
                                        If MapNpc(MapNum).NPC(i).vital(Vitals.HP) > 0 Then
                                            If IsinRange(Spell(spellnum).AoE, MapNpc(MapNum).NPC(aMapNPCNum).X, MapNpc(MapNum).NPC(aMapNPCNum).Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                                                InitDamage = GetNPCSpellDamage(spellnum, MapNum, aMapNPCNum)
                                                Select Case GetMapPetOwner(MapNum, aMapNPCNum)
                                                Case 0
                                                    SpellNpc_Effect Vitals.HP, True, i, InitDamage, spellnum, MapNum
                                                Case Is > 0
                                                    If i = aMapNPCNum Then
                                                        If Not HasNPCMaxVital(HP, MapNum, aMapNPCNum) Then
                                                            SpellNpc_Effect Vitals.HP, True, i, InitDamage, spellnum, MapNum
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
                                If Not HasNPCMaxVital(HP, MapNum, aMapNPCNum) Then
                                    InitDamage = GetNPCSpellDamage(spellnum, MapNum, aMapNPCNum)
                                    SpellNpc_Effect Vitals.HP, True, aMapNPCNum, InitDamage, spellnum, MapNum
                                    DidCast = True
                                End If
                            End If
                                
                           
                        ' AOE Damaging Spells
                        Case SPELL_TYPE_DAMAGEHP
                        ' Make sure an npc waits for the spell to cooldown
                            If Spell(spellnum).IsAoE Then
                                For i = 1 To MAX_MAP_NPCS
                                    If CanNpcAttackNpc(MapNum, aMapNPCNum, vMapNPCNum, True) Then
                                        If IsinRange(Spell(spellnum).AoE, MapNpc(MapNum).NPC(aMapNPCNum).X, MapNpc(MapNum).NPC(aMapNPCNum).Y, MapNpc(MapNum).NPC(vMapNPCNum).X, MapNpc(MapNum).NPC(vMapNPCNum).Y) Then
                                            
                                            Damage = GetSpellDamage(MapNum, aMapNPCNum, TARGET_TYPE_NPC, i, TARGET_TYPE_NPC, spellnum)
                                            If Damage <= 0 Then
                                                SendActionMsg MapNum, "¡RESISTISTE!", Pink, 1, (MapNpc(MapNum).NPC(vMapNPCNum).X) * 32, (MapNpc(MapNum).NPC(vMapNPCNum).Y * 32), , True
                                            Else
                                                If Spell(spellnum).StunDuration > 0 Then
                                                    StunNPC vMapNPCNum, MapNum, spellnum
                                                End If
                                                SendAnimation MapNum, Spell(spellnum).SpellAnim, MapNpc(MapNum).NPC(vMapNPCNum).X, MapNpc(MapNum).NPC(vMapNPCNum).Y, TARGET_TYPE_NPC, vMapNPCNum
                                                NpcAttackNpc MapNum, aMapNPCNum, vMapNPCNum, Damage
                                                DidCast = True
                                            End If
                                        End If
                                    End If
                                Next
                                
                            ' Non AoE Damaging Spells
                            Else
                                If CanNpcAttackNpc(MapNum, aMapNPCNum, vMapNPCNum, True) Then
                                    If IsinRange(Spell(spellnum).range, MapNpc(MapNum).NPC(aMapNPCNum).X, MapNpc(MapNum).NPC(aMapNPCNum).Y, MapNpc(MapNum).NPC(vMapNPCNum).X, MapNpc(MapNum).NPC(vMapNPCNum).Y) Then
                                        Damage = GetSpellDamage(MapNum, aMapNPCNum, TARGET_TYPE_NPC, vMapNPCNum, TARGET_TYPE_NPC, spellnum)
                                        If Damage <= 0 Then
                                            SendActionMsg MapNum, "¡RESISTISTE!", Pink, 1, (MapNpc(MapNum).NPC(vMapNPCNum).X) * 32, (MapNpc(MapNum).NPC(vMapNPCNum).Y * 32), , True
                                        Else
                                            If Spell(spellnum).StunDuration > 0 Then
                                                StunNPC vMapNPCNum, MapNum, spellnum
                                            End If
                                            SendAnimation MapNum, Spell(spellnum).SpellAnim, MapNpc(MapNum).NPC(vMapNPCNum).X, MapNpc(MapNum).NPC(vMapNPCNum).Y, TARGET_TYPE_NPC, vMapNPCNum
                                            NpcAttackNpc MapNum, aMapNPCNum, vMapNPCNum, Damage
                                            DidCast = True
                                        End If
                                    End If
                                End If
                            End If
                                

                        Case SPELL_TYPE_HEALMP
                            ' Make sure an npc waits for the spell to cooldown
                            If Spell(spellnum).IsAoE Then
                                'For i = 1 To MAX_PLAYERS
                                                   
                                                
                                'Next
                            Else
                                If Not HasNPCMaxVital(MP, MapNum, vMapNPCNum) Then
                                    ' Non AOE Healing Spells
                                    InitDamage = GetNPCSpellDamage(spellnum, MapNum, vMapNPCNum)
                                    SpellNpc_Effect Vitals.MP, True, vMapNPCNum, InitDamage, spellnum, MapNum
                                    DidCast = True
                                End If
                            End If
                        End Select
                        
                        If DidCast Then
                            MapNpc(MapNum).NPC(aMapNPCNum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                            SendNpcAttackAnimation MapNum, aMapNPCNum
                            If MapNpc(MapNum).NPC(aMapNPCNum).vital(Vitals.MP) - MPCost < 0 Then
                                MapNpc(MapNum).NPC(aMapNPCNum).vital(Vitals.MP) = 0
                            Else
                                MapNpc(MapNum).NPC(aMapNPCNum).vital(Vitals.MP) = MapNpc(MapNum).NPC(aMapNPCNum).vital(Vitals.MP) - MPCost
                            End If
                            SendMapNpcVitals MapNum, aMapNPCNum
                        Else
                            MapNpc(MapNum).NPC(aMapNPCNum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                        End If
                        
End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim npcnum As Long
Dim buffer As clsBuffer
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, victim) Then
    
        MapNum = GetPlayerMap(attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim, attacker) Then
            SendActionMsg MapNum, "¡Esquivado!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32), , True
            Exit Sub
        End If
        
        Call CheckStunPlayer(attacker, victim)

        ' Get the damage we can do
        Damage = GetPlayerDamageAgainstPlayer(attacker, victim)
        Damage = Damage - GetPlayerDefenseAgainstPlayer(victim, Damage)
        Damage = RAND(Damage * 0.8, Damage)
        
        ' * 1.5 if can crit
        If CanPlayerCriticalHit(attacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "¡Crítico!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32), , True
            SendSoundToMap MapNum, GetPlayerX(attacker), GetPlayerY(attacker), SoundEntity.seCritical, GetPlayerClass(attacker)
        Else
            SendSoundToMap MapNum, GetPlayerX(attacker), GetPlayerY(attacker), SoundEntity.seAttack, GetPlayerClass(attacker)
        End If

        
        If Damage > 0 Then
            Call Impactar(attacker, victim, Damage, TempPlayer(attacker).TargetType)
            Call PlayerAttackPlayer(attacker, victim, Damage)
        Else
            Call PlayerMsg(attacker, "Tu ataque no hizo nada.", BrightRed)
        End If
    End If
End Sub

Function CheckMapMorals(ByVal MapNum As Long, ByVal attacker As Long, ByVal victim As Long) As Boolean
    ' Check if map is attackable
    Dim moral As Byte
    moral = map(MapNum).moral
    Select Case moral
    Case MAP_MORAL_NONE
        CheckMapMorals = True
    Case MAP_MORAL_SAFE
        If IsPlayerNeutral(victim) Then
            Call PlayerMsg(attacker, "¡Ésta es una zona segura!", BrightRed)
        Else
            CheckMapMorals = True
        End If
    Case MAP_MORAL_ARENA
        CheckMapMorals = True
    Case MAP_MORAL_PK_SAFE
        If Not IsPlayerNeutral(victim) Then
            PlayerMsg attacker, "¡Esta es una zona de PK's!", BrightRed
            CheckMapMorals = False
        Else
            CheckMapMorals = True
        End If
    Case MAP_MORAL_PACIFIC
        Call PlayerMsg(attacker, "¡Ésta es una zona segura!", BrightRed)
        CheckMapMorals = False
    End Select
    
    
End Function

Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal IsProjectile As Boolean = False) As Boolean

    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            If GetRealTickCount < TempPlayer(attacker).AttackTimer + item(GetPlayerEquipment(attacker, Weapon)).Speed Then Exit Function
        Else
            If GetRealTickCount < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
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

    If Not CheckMapMorals(GetPlayerMap(attacker), attacker, victim) Then Exit Function
    
    If map(GetPlayerMap(attacker)).moral = MAP_MORAL_ARENA Then
        CanPlayerAttackPlayer = True
        Exit Function
    End If
    
    'Safe Mode
    If CheckSafeMode(attacker, victim) = True Then
        Exit Function
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess_Mode(attacker) > ADMIN_MONITOR Then
        If GetPlayerVisible(victim) = 0 Then: Call PlayerMsg(attacker, "Los administradores no pueden atacar a otros usuarios.", Cyan)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess_Mode(victim) > ADMIN_MONITOR Then
        'If GetPlayerVisible(victim) = 0 Then: Call PlayerMsg(attacker, "No puedes atacar a" & GetPlayerName(victim) & "!", BrightRed)
        Exit Function
    End If

    'Check if levels are correct
    If Not CheckLevels(attacker, victim) Then
        Exit Function
    End If
    
    If Not CanPlayerAttackByJustice(attacker, victim) Then
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
    If player(attacker).GuildFileId > 1 Then
        If player(attacker).GuildFileId = player(victim).GuildFileId Then
            'Call PlayerMsg(attacker, "¡" & GetPlayerName(victim) & " es miembro de tu clan!", BrightRed)
            Exit Function
        End If
    End If
    
    If TempPlayer(attacker).inParty > 0 Then
        If TempPlayer(victim).inParty = TempPlayer(attacker).inParty Then
            Exit Function
        End If
    End If
        
    TempPlayer(attacker).TargetType = TARGET_TYPE_PLAYER
    TempPlayer(attacker).Target = victim
    SendTarget attacker
    
    CanPlayerAttackPlayer = True
    
End Function

Public Function CheckLevels(ByVal attacker As Long, ByVal victim As Long, Optional ByVal sendmsg As Boolean = True) As Boolean
CheckLevels = False
If Not IsPlaying(attacker) Or Not IsPlaying(victim) Then Exit Function

' Make sure attacker is high enough level
If GetPlayerLevel(attacker) < 15 Then
    If sendmsg Then: Call PlayerMsg(attacker, "¡Estás bajo nivel 15, aún no puedes atacar a nadie!", BrightRed)
    Exit Function
End If

If GetPlayerLevel(victim) < 15 Then
    If sendmsg Then: Call PlayerMsg(attacker, "¡Tu objetivo esta debajo del nivel 15, no puedes atacarlo!", BrightRed)
    Exit Function
End If

'If GetPlayerLevel(attacker) >= 20 And GetPlayerLevel(victim) < 20 Then
    'If sendmsg Then: Call PlayerMsg(attacker, "¡Tu objetivo esta debajo del nivel 20 y tu estas por encima, no puedes atacarlo!", BrightRed)
    'Exit Function
'End If

'If GetPlayerLevel(attacker) < 20 And GetPlayerLevel(victim) >= 20 Then
    'If sendmsg Then: Call PlayerMsg(attacker, "¡Tu objetivo esta encima del nivel 20 y tu estas por debajo, no puedes atacarlo!", BrightRed)
    'Exit Function
'End If

CheckLevels = True
End Function
Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0)
    Dim exp As Long
    Dim N As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Dim MapNum As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    N = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        N = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetRealTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If spellnum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellnum
        
        If Not map(GetPlayerMap(attacker)).moral = MAP_MORAL_ARENA Then
            
            ' Player is dead
            Call GlobalMsg(GetPlayerName(victim) & " " & GetTranslation(" ha sido asesinado por ") & " " & GetPlayerName(attacker), BrightRed, False)
            ' Calculate exp to give attacker
            Call ComputePlayerExp(attacker, TARGET_TYPE_PLAYER, victim, TARGET_TYPE_PLAYER)
            
            
            'ALATAR
            Call CheckTasks(attacker, QUEST_TYPE_GOKILL, victim)
            '/ALATAR
            
            'Only If victim level + 20 >= attacker level
            If Not (GetLevelDifference(attacker, victim) > 20) Then
                Call PlayerPVPDrops(victim)
            End If
            
            Call SetPlayerJustice(attacker, victim)
            Call ComputeArmyPvP(attacker, victim)
        Else
            Call GlobalMsg(GetPlayerName(attacker) & " " & GetTranslation(" ha vencido a ") & " " & GetPlayerName(victim), White, False)
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If player(i).map = GetPlayerMap(attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
        
        
        'player death
        Call OnDeath(victim, 2)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)

        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        If spellnum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellnum
        
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetRealTickCount
        
        'if a stunning spell, stun the player
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then
                CheckSpellStunts victim, spellnum
            End If
            
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Player victim, spellnum, attacker
            End If
        End If
        
        ' send animation
        If spellnum = 0 Then
            If GetPlayerEquipment(attacker, Weapon) > 0 Then
                Call SendAnimation(GetPlayerMap(victim), item(GetPlayerEquipment(attacker, Weapon)).Animation, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim)
            End If
        End If
        
         SendSoundToMap GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seHit, GetPlayerClass(victim)
         
         
        
         
         Call SetPlayerHitJustice(attacker, victim)
    End If
    
    
    ' Reset attack timer
    TempPlayer(attacker).AttackTimer = GetRealTickCount
    
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal index As Long, ByVal spellslot As Long)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim range As Long
    Dim HasBuffered As Boolean
    
    Dim TargetType As Byte
    Dim Target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    spellnum = GetPlayerSpell(index, spellslot)
    MapNum = GetPlayerMap(index)
    
    If spellnum <= 0 Or spellnum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > GetRealTickCount Then
        'PlayerMsg index, "Habilidad recargándose", BrightRed
        Exit Sub
    End If
    
    If TempPlayer(index).spellBuffer.Spell <> 0 Then
        Exit Sub
    End If
    
    If Not CanSpell(index) Then Exit Sub

    If IsActionBlocked(index, aSpell) Then
    If Not Spell(spellnum).Type = SPELL_TYPE_PROTECT Then Exit Sub
    End If
    
    MPCost = GetSpellMPCost(GetPlayerMap(index), index, TARGET_TYPE_PLAYER, spellnum)

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        'Call PlayerMsg(index, "No tienes suficiente MP!", BrightRed)
        SendActionMsg MapNum, "No MP!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        Exit Sub
    End If
    
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, GetTranslation("Necesitas el lvl") & " " & LevelReq & " " & GetTranslation("para usar esta habilidad."), BrightRed, , False)
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
        Exit Sub
    End If
    
    AccessReq = Spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess_Mode(index) Then
        Call PlayerMsg(index, "Necesitas ser admin.", BrightRed)
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
        Exit Sub
    End If
    
    ClassReq = Spell(spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, GetTranslation("Solo la clase") & " " & CheckGrammar(Trim$(Class(ClassReq).TranslatedName)) & " " & GetTranslation("puede utilizar esta magia."), BrightRed, , False)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    TargetType = TempPlayer(index).TargetType
    Target = TempPlayer(index).Target
    range = Spell(spellnum).range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not Target > 0 Then
                'PlayerMsg index, "No tienes un objetivo.", BrightRed
            End If
            If TargetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not IsinRange(range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(Target), GetPlayerY(Target)) Then
                    'PlayerMsg index, "Objetivo fuera de rango.", BrightRed
                    SendActionMsg MapNum, "Fuera de rango.", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, , True
                Else
                    ' go through spell types
                    If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf TargetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not IsinRange(range, GetPlayerX(index), GetPlayerY(index), MapNpc(MapNum).NPC(Target).X, MapNpc(MapNum).NPC(Target).Y) Then
                    'PlayerMsg index, "Objetivo fuera de rango.", BrightRed
                    SendActionMsg MapNum, "Fuera de rango.", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, , True
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        If Spell(spellnum).Type = SPELL_TYPE_WARP Then
            SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellnum
        End If
        SendAnimation MapNum, Spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, index
        'here sound
        SendActionMsg MapNum, "Casting " & Trim$(Spell(spellnum).TranslatedName) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, , False
        TempPlayer(index).spellBuffer.Spell = spellslot
        TempPlayer(index).spellBuffer.Timer = GetRealTickCount
        TempPlayer(index).spellBuffer.Target = TempPlayer(index).Target
        TempPlayer(index).spellBuffer.tType = TempPlayer(index).TargetType
        Exit Sub
    Else
        SendClearSpellBuffer index
    End If
End Sub

Public Sub CastSpell(ByVal index As Long, ByVal spellslot As Long, ByVal Target As Long, ByVal TargetType As Byte)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim X As Long, Y As Long
   
    Dim buffer As clsBuffer
    Dim SpellCastType As Long
   
    DidCast = False
    

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellnum = GetPlayerSpell(index, spellslot)
    TempPlayer(index).LastSpell = spellnum
    MapNum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub

    MPCost = GetSpellMPCost(GetPlayerMap(index), index, TARGET_TYPE_PLAYER, spellnum)

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        'Call PlayerMsg(index, "¡No tienes suficiente MP!", BrightRed)
        SendActionMsg MapNum, "No MP!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        Exit Sub
    End If
   
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, GetTranslation("Debes ser de nivel") & " " & LevelReq & " " & GetTranslation("para usar ésta habilidad."), BrightRed, , False)
        Exit Sub
    End If
   
    AccessReq = Spell(spellnum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess_Mode(index) Then
        Call PlayerMsg(index, "Debes ser administrador para usar esta habilidad.", BrightRed)
        Exit Sub
    End If
   
    ClassReq = Spell(spellnum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, GetTranslation("Solo") & " " & Trim$((Class(ClassReq).TranslatedName)) & " " & GetTranslation("puede usar esta habilidad."), BrightRed, , False)
            Exit Sub
        End If
    End If
   
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
   
    ' set the vital
    vital = GetPlayerSpellDamageAgainstPlayer(index, spellnum)
    AoE = Spell(spellnum).AoE
    range = Spell(spellnum).range
   
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_HEALHP
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellPlayer_Effect Vitals.HP, True, index, vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellPlayer_Effect Vitals.MP, True, index, vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SetPlayerDir index, Spell(spellnum).dir
                    PlayerWarpBySpell index, spellnum
                    DidCast = True
                Case SPELL_TYPE_BUFFER
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellStatBuffer index, spellnum
                    DidCast = True
                Case SPELL_TYPE_PROTECT
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellProtect index, spellnum
                    DidCast = True
                Case SPELL_TYPE_CHANGESTATE
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellChangeState index, spellnum
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                X = GetPlayerX(index)
                Y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
                If TargetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
               
                If TargetType = TARGET_TYPE_PLAYER Then
                    X = GetPlayerX(Target)
                    Y = GetPlayerY(Target)
                Else
                    X = MapNpc(MapNum).NPC(Target).X
                    Y = MapNpc(MapNum).NPC(Target).Y
                End If
               
                If Not IsinRange(range, GetPlayerX(index), GetPlayerY(index), X, Y) Then
                    PlayerMsg index, "El objetivo no está al alcance.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If IsinRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, True) Then
                                            SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PlayerAttackPlayer index, i, GetSpellDamage(MapNum, index, TARGET_TYPE_PLAYER, i, TARGET_TYPE_PLAYER, spellnum), spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).NPC(i).Num > 0 Then
                            If MapNpc(MapNum).NPC(i).vital(HP) > 0 Then
                                If IsinRange(AoE, X, Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc index, i, GetSpellDamage(MapNum, index, TARGET_TYPE_PLAYER, i, TARGET_TYPE_NPC, spellnum), spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                   
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If IsinRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect VitalType, increment, i, vital, spellnum
                                    DidCast = True
                                End If
                            End If
                        End If
                    End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).NPC(i).Num > 0 Then
                            If MapNpc(MapNum).NPC(i).vital(HP) > 0 Then
                                If IsinRange(AoE, X, Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                                    SpellNpc_Effect VitalType, increment, i, vital, spellnum, MapNum
                                End If
                            End If
                        End If
                    Next
                    
                Case SPELL_TYPE_PROTECT
                
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If IsinRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellProtect i, spellnum
                                End If
                                End If
                            End If
                        End If
                    Next
            End Select
        Case 2 ' targetted
            If TargetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
           
            If TargetType = TARGET_TYPE_PLAYER Then
                X = GetPlayerX(Target)
                Y = GetPlayerY(Target)
            Else
                X = MapNpc(MapNum).NPC(Target).X
                Y = MapNpc(MapNum).NPC(Target).Y
            End If
               
            If Not IsinRange(range, GetPlayerX(index), GetPlayerY(index), X, Y) Then
                PlayerMsg index, "El objetivo no está al alcance.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
            
            If Spell(spellnum).Type = SPELL_TYPE_DAMAGEHP Then
                If TargetType = TARGET_TYPE_PLAYER And Target = index Then
                    PlayerMsg index, "No puedes atacarte a ti mismo.", BrightRed
                    Exit Sub
                End If
            End If
            
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, True) Then
                                SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                PlayerAttackPlayer index, Target, GetSpellDamage(MapNum, index, TARGET_TYPE_PLAYER, Target, TARGET_TYPE_PLAYER, spellnum), spellnum
                                DidCast = True
                        End If
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                                SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                PlayerAttackNpc index, Target, GetSpellDamage(MapNum, index, TARGET_TYPE_PLAYER, Target, TARGET_TYPE_NPC, spellnum), spellnum
                                DidCast = True
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    End If
                    
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                   
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, vital, spellnum
                                SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                                SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, vital, spellnum, MapNum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, vital, spellnum, MapNum
                        End If
                    End If
                Case SPELL_TYPE_PROTECT
                    SpellProtect Target, spellnum
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
            End Select
    End Select
   
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        TempPlayer(index).SpellCD(spellslot) = GetRealTickCount + (Spell(spellnum).CDTime * 1000)
        Call SendCooldown(index, spellslot)
        SendActionMsg MapNum, Trim$(Spell(spellnum).TranslatedName) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellnum As Long)
Dim sSymbol As String * 1
Dim colour As Long
Dim MaxVital As Long
Dim VitalComp As Long
            If Damage > 0 Then
        
                MaxVital = GetPlayerMaxVital(index, vital)
                If increment Then
                
                    If Spell(spellnum).Duration > 0 Then
                        AddHoT_Player index, spellnum
                    End If
                    'add vital
                    VitalComp = MaxVital
                    sSymbol = "+"
                    
                    If vital = Vitals.HP Then colour = BrightGreen
                    If vital = Vitals.MP Then colour = Cyan
                Else
                    'substract vital
                    VitalComp = 0
                    Damage = -Damage
                    
                    sSymbol = "-"
                    colour = Cyan
                End If
                
                If GetPlayerVital(index, vital) = VitalComp Then
                    'Time saver
                    Exit Sub
                ElseIf (GetPlayerVital(index, vital) + Damage >= VitalComp And increment) Or (GetPlayerVital(index, vital) + Damage <= VitalComp And Not increment) Then
                    SetPlayerVital index, vital, VitalComp
                Else
                    SetPlayerVital index, vital, GetPlayerVital(index, vital) + Damage
                End If
                
                
                If spellnum > 0 Then
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, GetPlayerX(index), GetPlayerY(index), TARGET_TYPE_PLAYER, index
                End If
                SendActionMsg GetPlayerMap(index), sSymbol & Damage, colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                
                ' send the sound
                SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellnum
                
                Call SendVital(index, vital)
                
        End If
End Sub

Public Sub SpellNpc_Effect(ByVal vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellnum As Long, ByVal MapNum As Long)
Dim sSymbol As String * 1
Dim colour As Long
Dim MaxVital As Long
Dim VitalComp As Long

'Do not us this procedeture on hp substracting

        If Damage > 0 Then
        
                MaxVital = GetNpcMaxVital(MapNum, index, vital)
                If increment Then
                    'add vital
                    VitalComp = MaxVital
                    sSymbol = "+"
                    
                    If vital = Vitals.HP Then colour = BrightGreen
                    If vital = Vitals.MP Then colour = Cyan
                Else
                    'substract vital
                    VitalComp = 0
                    Damage = -Damage
                    
                    sSymbol = "-"
                    colour = Cyan
                End If
                
                If MapNpc(MapNum).NPC(index).vital(vital) = VitalComp Then
                    'Time saver
                    Exit Sub
                ElseIf (MapNpc(MapNum).NPC(index).vital(vital) + Damage >= MaxVital And increment) Or (MapNpc(MapNum).NPC(index).vital(vital) + Damage <= VitalComp And Not increment) Then
                    MapNpc(MapNum).NPC(index).vital(vital) = VitalComp
                Else
                    MapNpc(MapNum).NPC(index).vital(vital) = MapNpc(MapNum).NPC(index).vital(vital) + Damage
                End If
                
                
                SendAnimation MapNum, Spell(spellnum).SpellAnim, MapNpc(MapNum).NPC(index).X, MapNpc(MapNum).NPC(index).Y, TARGET_TYPE_NPC, index
                SendActionMsg MapNum, sSymbol & Damage, colour, ACTIONMSG_SCROLL, MapNpc(MapNum).NPC(index).X * 32, MapNpc(MapNum).NPC(index).Y * 32
                
                ' send the sound
                SendMapSound index, MapNpc(MapNum).NPC(index).X, MapNpc(MapNum).NPC(index).Y, SoundEntity.seSpell, spellnum
                
                Call SendMapNpcVitals(MapNum, index)
                
        End If
End Sub

Public Sub AddDoT_Player(ByVal index As Long, ByVal spellnum As Long, ByVal caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            If .Spell = spellnum Then
                .Timer = GetRealTickCount
                .caster = caster
                .StartTime = GetRealTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = GetRealTickCount
                .caster = caster
                .Used = True
                .StartTime = GetRealTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal index As Long, ByVal spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).HoT(i)
            If .Spell = spellnum Then
                .Timer = GetRealTickCount
                .StartTime = GetRealTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = GetRealTickCount
                .Used = True
                .StartTime = GetRealTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal spellnum As Long, ByVal caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).NPC(index).DoT(i)
            If .Spell = spellnum Then
                .Timer = GetRealTickCount
                .caster = caster
                .StartTime = GetRealTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = GetRealTickCount
                .caster = caster
                .Used = True
                .StartTime = GetRealTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).NPC(index).HoT(i)
            If .Spell = spellnum Then
                .Timer = GetRealTickCount
                .StartTime = GetRealTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = GetRealTickCount
                .Used = True
                .StartTime = GetRealTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal index As Long, ByVal dotNum As Long)
    With TempPlayer(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetRealTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.caster, index, True) Then
                    PlayerAttackPlayer .caster, index, GetPlayerDotDamage(index, dotNum)
                End If
                .Timer = GetRealTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetRealTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Function GetPlayerDotDamage(ByVal index As Long, ByVal dotNum As Long) As Long
    'index > 0
    'tempplayer(index).dot(dotnum).spell > 0
    With TempPlayer(index).DoT(dotNum)
    Dim SpellDamage As Long
    If .caster > 0 Then
        SpellDamage = GetSpellDamage(GetPlayerMap(index), .caster, TARGET_TYPE_PLAYER, index, TARGET_TYPE_PLAYER, .Spell)
    End If
    
    
    Dim HappenedIntervals As Integer
    If Spell(.Spell).Interval = 0 Then
        HappenedIntervals = 1
    Else
        HappenedIntervals = (.Timer - .StartTime) / (Spell(.Spell).Interval * 1000)
        HappenedIntervals = HappenedIntervals + 1
    End If
    
    If HappenedIntervals > 0 And HappenedIntervals < 15 Then
        GetPlayerDotDamage = SpellDamage / (2 ^ HappenedIntervals)
    End If
        
    End With
    
End Function

Public Sub HandleHoT_Player(ByVal index As Long, ByVal hotNum As Long)
    With TempPlayer(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetRealTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                
                Dim spelltype As Byte
                spelltype = Spell(.Spell).Type
                Select Case spelltype
                Case SPELL_TYPE_HEALMP
                    SetPlayerVital index, Vitals.MP, GetPlayerVital(index, Vitals.MP) + Spell(.Spell).vital
                    SendVital index, Vitals.MP
                Case SPELL_TYPE_HEALHP
                    SetPlayerVital index, Vitals.HP, GetPlayerVital(index, Vitals.HP) + Spell(.Spell).vital
                    SendVital index, Vitals.HP
                Case Else
                    Exit Sub
                End Select
                'SendActionMsg Player(Index).Map, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, Player(Index).x * 32, Player(Index).y * 32
                
                .Timer = GetRealTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetRealTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal dotNum As Long)
    With MapNpc(MapNum).NPC(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetRealTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.caster, index, True) Then
                    PlayerAttackNpc .caster, index, GetNPCDotDamage(MapNum, index, dotNum), , True
                End If
                .Timer = GetRealTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetRealTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Function GetNPCDotDamage(ByVal MapNum As Long, ByVal index As Long, ByVal dotNum As Long) As Long
    'index > 0
    'tempplayer(index).dot(dotnum).spell > 0
    With MapNpc(MapNum).NPC(index).DoT(dotNum)
    Dim SpellDamage As Long
    If .caster > 0 Then
        SpellDamage = GetSpellDamage(MapNum, .caster, TARGET_TYPE_PLAYER, index, TARGET_TYPE_NPC, .Spell)
    End If
    
    Dim HappenedIntervals As Integer
    
    If Spell(.Spell).Interval = 0 Then Exit Function
    HappenedIntervals = (.Timer - .StartTime) / (1000 * Spell(.Spell).Interval)
    HappenedIntervals = HappenedIntervals + 1
    
    If HappenedIntervals > 0 And HappenedIntervals < 16 Then
        GetNPCDotDamage = SpellDamage / (2 ^ HappenedIntervals)
    End If
        
    End With
    
End Function

Public Sub HandleHoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal hotNum As Long)
        With MapNpc(MapNum).NPC(index).HoT(hotNum)
                If .Used And .Spell > 0 Then
                        ' time to tick?
                        If GetRealTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                                        SendActionMsg MapNum, "+" & Spell(.Spell).vital, BrightGreen, ACTIONMSG_SCROLL, MapNpc(MapNum).NPC(index).X * 32, MapNpc(MapNum).NPC(index).Y * 32
                                        MapNpc(MapNum).NPC(index).vital(Vitals.HP) = MapNpc(MapNum).NPC(index).vital(Vitals.HP) + Spell(.Spell).vital
                                Else
                                        SendActionMsg MapNum, "+" & Spell(.Spell).vital, Cyan, ACTIONMSG_SCROLL, MapNpc(MapNum).NPC(index).X * 32, MapNpc(MapNum).NPC(index).Y * 32
                                        MapNpc(MapNum).NPC(index).vital(Vitals.MP) = MapNpc(MapNum).NPC(index).vital(Vitals.MP) + Spell(.Spell).vital
                                        
                                        If MapNpc(MapNum).NPC(index).vital(Vitals.MP) > GetNpcMaxVital(MapNum, index, MP) Then
                                                MapNpc(MapNum).NPC(index).vital(Vitals.MP) = GetNpcMaxVital(MapNum, index, MP)
                                        End If
                                End If
                                
                                .Timer = GetRealTickCount
                                ' check if DoT is still active - if NPC died it'll have been purged
                                If .Used And .Spell > 0 Then
                                        ' destroy hoT if finished
                                        If GetRealTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                                                .Used = False
                                                .Spell = 0
                                                .Timer = 0
                                                .caster = 0
                                                .StartTime = 0
                                        End If
                                End If
                        End If
                End If
        End With
End Sub

Public Sub StunPlayer(ByVal index As Long, ByVal spellnum As Long)
    If Spell(spellnum).StunDuration > 0 Then
        If GPE(index) Then Exit Sub
        BlockPlayerAction index, aMove, CSng(Spell(spellnum).StunDuration)
        BlockPlayerAction index, aAttack, CSng(Spell(spellnum).StunDuration)
        PlayerMsg index, "¡Estás paralizado!", BrightRed
    End If
End Sub

Public Sub StunPlayerByTime(ByVal index As Long, ByVal Time As Single)
    If GPE(index) Then Exit Sub
    BlockPlayerAction index, aMove, Time
    BlockPlayerAction index, aAttack, Time
End Sub
Public Sub StunNPC(ByVal index As Long, ByVal MapNum As Long, ByVal spellnum As Long)
Dim npcnum As Long
    ' check if it's a stunning spell
    If Spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(MapNum).NPC(index).StunDuration = Spell(spellnum).StunDuration - NPC(index).level
        MapNpc(MapNum).NPC(index).StunTimer = GetRealTickCount
    End If
End Sub

Public Function CalculateDropChances(ByVal npcnum As Long) As Integer

'equal probability distribution
Dim i As Byte, N As Byte, j As Byte
Dim BoolVect(1 To MAX_NPC_DROPS) As Boolean

For i = 1 To CByte(MAX_NPC_DROPS)
    BoolVect(i) = False
Next

N = 0
For i = 1 To CByte(MAX_NPC_DROPS)
    If NPC(npcnum).Drops(i).DropItem > 0 Then
        N = N + 1
        BoolVect(i) = True
    End If
Next
If N = 0 Then
    CalculateDropChances = 0
Else
    i = RAND(1, N)
    j = 0
    For N = 1 To CByte(MAX_NPC_DROPS)
        If BoolVect(N) = True Then
            j = j + 1
        End If
        If j = i Then
            CalculateDropChances = N
            Exit Function
        End If
    Next
End If
            
End Function

' ###################################
' ##      NPC Attacking NPC     ##
' ###################################

Public Sub TryNpcAttackNpc(ByVal MapNum As Long, ByVal mapNPCnumAttacker As Long, ByVal mapNpcNumVictim As Long)
Dim NPCNumAttacker As Long, NPCNumVictim As Long, blockAmount As Long, Damage As Long

    ' Can the npc attack the player?
    If CanNpcAttackNpc(MapNum, mapNPCnumAttacker, mapNpcNumVictim) Then
    
        NPCNumAttacker = MapNpc(MapNum).NPC(mapNPCnumAttacker).Num
        NPCNumVictim = MapNpc(MapNum).NPC(mapNpcNumVictim).Num
        
        
        ' send the sound
        MapNpc(MapNum).NPC(mapNPCnumAttacker).AttackTimer = GetRealTickCount + GetNPCAttackTimer(MapNum, mapNPCnumAttacker)
        ' send the sound
        SendSoundToMap MapNum, MapNpc(MapNum).NPC(mapNpcNumVictim).X, MapNpc(MapNum).NPC(mapNpcNumVictim).Y, SoundEntity.seNpc, MapNpc(MapNum).NPC(mapNPCnumAttacker).Num

        ' check if NPC can avoid the attack
        If CanNpcDodgeNpc(MapNum, mapNPCnumAttacker, mapNpcNumVictim) Then
            SendActionMsg MapNum, "¡Esquivado!", Pink, 1, (MapNpc(MapNum).NPC(mapNpcNumVictim).X * 32), (MapNpc(MapNum).NPC(mapNpcNumVictim).Y * 32), , True
            Exit Sub
        End If
        If CanNpcParryNpc(MapNum, mapNPCnumAttacker, mapNpcNumVictim) Then
            SendActionMsg MapNum, "¡Bloqueado!", Pink, 1, (MapNpc(MapNum).NPC(mapNpcNumVictim).X * 32), (MapNpc(MapNum).NPC(mapNpcNumVictim).Y * 32), , True
            Exit Sub
        End If

        Damage = GetNpcDamage(MapNum, mapNPCnumAttacker)
        Damage = Damage - GetNpcDefense(MapNum, mapNpcNumVictim, Damage)
        Damage = RAND(Damage * 0.8, Damage)
        
        If CanNpcCrit(NPCNumAttacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "¡Crítico!", BrightCyan, 1, (MapNpc(MapNum).NPC(mapNPCnumAttacker).X * 32), (MapNpc(MapNum).NPC(mapNPCnumAttacker).Y * 32), , True
        End If

        If Damage > 0 Then
            Call NpcAttackNpc(MapNum, mapNPCnumAttacker, mapNpcNumVictim, Damage)
        Else
            SendActionMsg MapNum, "¡Evitado!", Cyan, 1, MapNpc(MapNum).NPC(mapNpcNumVictim).X * 32, MapNpc(MapNum).NPC(mapNpcNumVictim).Y * 32, , True
        End If
    End If
End Sub



Function CanNpcAttackNpc(ByVal MapNum As Long, ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
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
    If MapNpc(MapNum).NPC(attacker).Num <= 0 Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(MapNum).NPC(victim).Num <= 0 Then
        Exit Function
    End If

    aNPCNum = MapNpc(MapNum).NPC(attacker).Num
    vNPCNum = MapNpc(MapNum).NPC(victim).Num
    
    
    If aNPCNum <= 0 Then Exit Function
    If vNPCNum <= 0 Then Exit Function
    
    'Pet check
    If IsMapNPCaPet(MapNum, attacker) And IsMapNPCaPet(MapNum, victim) Then
        If Not (CanPetAttackPet(MapNum, attacker, victim)) Then
            Exit Function
        End If
    End If
    
    'Check npc type
    If CanNPCBeAttacked(vNPCNum) = False Then Exit Function

    ' Make sure the npcs arent already dead
    If MapNpc(MapNum).NPC(attacker).vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).NPC(victim).vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    If IsSpell Then
        CanNpcAttackNpc = True
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetRealTickCount < MapNpc(MapNum).NPC(attacker).AttackTimer Then
        Exit Function
    End If
    
    AttackerX = MapNpc(MapNum).NPC(attacker).X
    AttackerY = MapNpc(MapNum).NPC(attacker).Y
    VictimX = MapNpc(MapNum).NPC(victim).X
    VictimY = MapNpc(MapNum).NPC(victim).Y

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
Sub NpcAttackNpc(ByVal MapNum As Long, ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Dim aNPCNum As Long
    Dim vNPCNum As Long
    Dim N As Long
    Dim PetOwner As Long
    Dim DropNum As Integer
    
    If attacker <= 0 Or attacker > MAX_MAP_NPCS Then Exit Sub
    If victim <= 0 Or victim > MAX_MAP_NPCS Then Exit Sub
    
    If Damage <= 0 Then Exit Sub
    
    aNPCNum = MapNpc(MapNum).NPC(attacker).Num
    vNPCNum = MapNpc(MapNum).NPC(victim).Num
    
    If aNPCNum <= 0 Then Exit Sub
    If vNPCNum <= 0 Then Exit Sub
    
    'set the victim's target to the pet attacking it
    MapNpc(MapNum).NPC(victim).TargetType = 2 'Npc
    MapNpc(MapNum).NPC(victim).Target = attacker
    
    ' set the regen timer
    MapNpc(MapNum).NPC(attacker).stopRegen = True
    MapNpc(MapNum).NPC(victim).stopRegenTimer = GetRealTickCount
    
    ' Send this packet so they can see the person attacking
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcAttack
    buffer.WriteLong attacker
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing

    If Damage >= MapNpc(MapNum).NPC(victim).vital(Vitals.HP) Then
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).NPC(victim).X * 32), (MapNpc(MapNum).NPC(victim).Y * 32)
        SendBlood MapNum, MapNpc(MapNum).NPC(victim).X, MapNpc(MapNum).NPC(victim).Y
        
        ' Set NPC target to 0
        MapNpc(MapNum).NPC(attacker).Target = 0
        MapNpc(MapNum).NPC(attacker).TargetType = 0
        'reset the targetter for the player
        
        'Call the calculator for deciding which item has to be chanced
        DropNum = CalculateDropChances(vNPCNum)
        
        'Drop the goods if they get it
        If DropNum > 0 Then
            N = Int(Rnd * NPC(vNPCNum).Drops(DropNum).DropChance) + 1

            'Drop the selected item
            If N = 1 Then
                Call SpawnItem(NPC(vNPCNum).Drops(DropNum).DropItem, NPC(vNPCNum).Drops(DropNum).DropItemValue, MapNum, MapNpc(MapNum).NPC(victim).X, MapNpc(MapNum).NPC(victim).Y)
            End If
        End If
        
        
        If IsMapNPCaPet(MapNum, attacker) Then
            TempPlayer(MapNpc(MapNum).NPC(attacker).PetData.Owner).Target = 0
            TempPlayer(MapNpc(MapNum).NPC(attacker).PetData.Owner).TargetType = TARGET_TYPE_NONE
            
            PetOwner = MapNpc(MapNum).NPC(attacker).PetData.Owner
            
            SendTarget PetOwner
            
            Call ComputePlayerExp(PetOwner, TARGET_TYPE_PLAYER, victim, TARGET_TYPE_NPC)
            'objective finished
            TempPlayer(PetOwner).TempPet.PetHasOwnTarget = 0
            
            Call CheckPlayerPartyTasks(PetOwner, QUEST_TYPE_GOSLAY, vNPCNum)
            
        End If
                      
        If IsMapNPCaPet(MapNum, victim) Then
            'Get the pet owners' index
            PetOwner = MapNpc(MapNum).NPC(victim).PetData.Owner
            'Set the NPC's target on the owner now
            MapNpc(MapNum).NPC(attacker).TargetType = 1 'player
            MapNpc(MapNum).NPC(attacker).Target = PetOwner
            
            'objective finished
            TempPlayer(PetOwner).TempPet.PetHasOwnTarget = 0
            'Set Spawn time
            TempPlayer(PetOwner).TempPet.PetSpawnWait = GetRealTickCount
                   
        End If
        
        KillNpc MapNum, victim
        
        If PetOwner > 0 Then
            PetFollowOwner PetOwner
        End If
    Else
        ' npc not dead, just do the damage
        MapNpc(MapNum).NPC(victim).vital(Vitals.HP) = MapNpc(MapNum).NPC(victim).vital(Vitals.HP) - Damage
       
        ' Say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).NPC(victim).X * 32), (MapNpc(MapNum).NPC(victim).Y * 32)
        SendBlood MapNum, MapNpc(MapNum).NPC(victim).X, MapNpc(MapNum).NPC(victim).Y
        
        ' set the regen timer
        MapNpc(MapNum).NPC(victim).stopRegen = True
        MapNpc(MapNum).NPC(victim).stopRegenTimer = GetRealTickCount
    End If
    
    'Send both Npc's Vitals to the client
    SendMapNpcVitals MapNum, attacker
    SendMapNpcVitals MapNum, victim

End Sub

Function CalculateLosenExp(ByVal attacker As Long, ByVal victim As Long) As Long

Dim Difference As Long

Difference = GetLevelDifference(attacker, victim)

Dim reduction As Double
reduction = Line(MAX_LEVELS, 0, 10, 4, 4, Difference)

'can't run time 6
CalculateLosenExp = CDbl(GetPlayerExp(victim)) / reduction

End Function



Public Function CanNpcDodgeNpc(ByVal MapNum As Long, ByVal attacker As Long, ByVal victim As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodgeNpc = False
    rate = (GetNpcStat(MapNum, victim, Agility) - GetNpcStat(MapNum, attacker, Agility))
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

Public Function CanNpcParryNpc(ByVal MapNum As Long, ByVal attacker As Long, ByVal victim As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParryNpc = False
    rate = (GetNpcStat(MapNum, attacker, Strength) - GetNpcStat(MapNum, victim, Strength))
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
Public Function CanPetAttackPlayer(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal victim As Long) As Boolean
    CanPetAttackPlayer = False
    
    'Check if pet
    If Not IsMapNPCaPet(MapNum, mapnpcnum) Then Exit Function
    
    Dim Owner As Long
    Owner = GetMapPetOwner(MapNum, mapnpcnum)
    ' Check if map is attackable
    If Not CheckMapMorals(MapNum, Owner, victim) Then
        Exit Function
    End If

    ' Check to make sure that they dont have access
    If GetPlayerAccess_Mode(Owner) > ADMIN_MONITOR Then
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess_Mode(victim) > ADMIN_MONITOR Then
        Exit Function
    End If
    
    If Not CheckLevels(Owner, victim, False) Then
        Exit Function
    End If
    
    If Not CanPlayerAttackByJustice(Owner, victim, False) Then
        Exit Function
    End If
    
    
    'make sure victim is not a guild partner
    If player(Owner).GuildFileId > 1 Then
        If player(Owner).GuildFileId = player(victim).GuildFileId Then
        Exit Function
        End If
    End If
    
    If TempPlayer(Owner).inParty > 1 Then
        If TempPlayer(Owner).inParty = TempPlayer(victim).inParty Then
            Exit Function
        End If
    End If
    
    CanPetAttackPlayer = True
    
End Function

Public Function CanPlayerAttackPet(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal attacker As Long) As Boolean
    CanPlayerAttackPet = False
    
    'Check if pet
    If Not IsMapNPCaPet(MapNum, mapnpcnum) Then Exit Function
    
    Dim Owner As Long
    Owner = GetMapPetOwner(MapNum, mapnpcnum)
    
    ' Check if map is attackable
    If Not CheckMapMorals(MapNum, attacker, Owner) Then
        Exit Function
    End If

    ' Check to make sure that they dont have access
    If GetPlayerAccess_Mode(attacker) > ADMIN_MONITOR Then
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess_Mode(Owner) > ADMIN_MONITOR Then
        Exit Function
    End If

    If Not CheckLevels(attacker, Owner, False) Then
        Exit Function
    End If
    
    If CanPlayerAttackByJustice(attacker, Owner, False) Then
        Exit Function
    End If
    
    
    
    
    'make sure victim is not a guild partner
    If player(attacker).GuildFileId > 1 Then
        If player(attacker).GuildFileId = player(Owner).GuildFileId Then
            Exit Function
        End If
    End If
    
    If TempPlayer(attacker).inParty > 1 Then
        If TempPlayer(attacker).inParty = TempPlayer(Owner).inParty Then
            Exit Function
        End If
    End If
    
    CanPlayerAttackPet = True
    
End Function

Public Function CanPetAttackPet(ByVal MapNum As Long, ByVal aMapNPCNum As Long, ByVal vMapNPCNum As Long) As Boolean
    CanPetAttackPet = False
    
    'Check if pet
    If Not IsMapNPCaPet(MapNum, aMapNPCNum) Then Exit Function
    If Not IsMapNPCaPet(MapNum, vMapNPCNum) Then Exit Function
    
    Dim aOwner As Long
    Dim vOwner As Long
    
    aOwner = GetMapPetOwner(MapNum, aMapNPCNum)
    vOwner = GetMapPetOwner(MapNum, vMapNPCNum)
    
    ' Check if map is attackable
    If Not map(MapNum).moral = MAP_MORAL_NONE Then
        If IsPlayerNeutral(vOwner) Then
            Exit Function
        End If
    End If

    ' Check to make sure that they dont have access
    If GetPlayerAccess_Mode(aOwner) > ADMIN_MONITOR Then
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess_Mode(vOwner) > ADMIN_MONITOR Then
        Exit Function
    End If

    If Not CheckLevels(aOwner, vOwner, False) Then
        Exit Function
    End If
    
    If Not CanPlayerAttackByJustice(aOwner, vOwner, False) Then
        Exit Function
    End If
    
    
    'make sure victim is not a guild partner
    If player(aOwner).GuildFileId > 1 Then
        If player(aOwner).GuildFileId = player(vOwner).GuildFileId Then
        Exit Function
        End If
    End If
    
    If TempPlayer(aOwner).inParty > 1 Then
        If TempPlayer(aOwner).inParty = TempPlayer(vOwner).inParty Then
            Exit Function
        End If
    End If
    
    CanPetAttackPet = True
    
End Function
Public Function GetSpellDamage(ByVal MapNum As Long, ByVal attacker As Long, ByVal attackertype As Byte, ByVal victim As Long, ByVal victimtype As Byte, ByVal spellnum As Long) As Long
    Dim Damage As Long
    Dim Protection As Long
    If MapNum < 1 Or MapNum > MAX_MAPS Or attacker < 1 Or victim < 1 Or spellnum < 1 Then Exit Function
    
    If attackertype = TARGET_TYPE_PLAYER And victimtype = TARGET_TYPE_NPC Then
    
        Damage = GetPlayerSpellDamageAgainstNPC(attacker, spellnum)
        Damage = Damage - GetNpcSpellDefense(MapNum, victim, Damage, spellnum)
        
    ElseIf attackertype = TARGET_TYPE_PLAYER And victimtype = TARGET_TYPE_PLAYER Then
    
        Damage = GetPlayerSpellDamageAgainstPlayer(attacker, spellnum)
        Damage = Damage - GetPlayerSpellDefenseAgainstPlayer(victim, Damage, spellnum)
        
    ElseIf attackertype = TARGET_TYPE_NPC And victimtype = TARGET_TYPE_PLAYER Then
    
        Damage = GetNPCSpellDamage(spellnum, MapNum, attacker)
        Damage = Damage - GetPlayerSpellDefenseAgainstNPC(victim, Damage, spellnum)
        
    ElseIf attackertype = TARGET_TYPE_NPC And victimtype = TARGET_TYPE_NPC Then
    
        Damage = GetNPCSpellDamage(spellnum, MapNum, attacker)
        Damage = Damage - GetNpcSpellDefense(MapNum, victim, Damage, spellnum)
        
    End If
    
    GetSpellDamage = Damage
    
    If GetSpellDamage < 0 Then
        GetSpellDamage = 0
    End If
        
End Function

Public Function GetPlayerSpellDamageAgainstNPC(ByVal index As Long, ByVal spellnum As Long) As Long
    Dim vitalval As Long
    vitalval = Spell(spellnum).vital
    
    Dim MinFactor As Single
    Dim MaxFactor As Single
    Dim factor As Single
    Dim X As Long
    

    MinFactor = 0.32
    MaxFactor = 0.64
    X = GetPlayerStat(index, GetSpellDamageStat(spellnum))

    'LOS VALORES DE FACTOR OSCILAN ENTRE MINFACTOR Y MAXFACTOR, SIEMPRE ENTRE ESOS DOS VALORES
    'OSCILAN SEGUN: STAT DEL JUGADOR
    factor = Line(MAX_STAT, 0, MaxFactor, MinFactor, MinFactor, X)
    
    
    GetPlayerSpellDamageAgainstNPC = vitalval * factor
End Function

Public Function GetPlayerSpellDamageAgainstPlayer(ByVal index As Long, ByVal spellnum As Long) As Long
    Dim vitalval As Long
    vitalval = Spell(spellnum).vital
    
    Dim MinFactor As Single
    Dim MaxFactor As Single
    Dim factor As Single
    Dim X As Long
    
    MinFactor = 0.32
    MaxFactor = 0.64
    X = GetPlayerStat(index, GetSpellDamageStat(spellnum))
    
    'LOS VALORES DE FACTOR OSCILAN ENTRE MINFACTOR Y MAXFACTOR, SIEMPRE ENTRE ESOS DOS VALORES
    'OSCILAN SEGUN: STAT DEL JUGADOR
    factor = Line(MAX_STAT, 0, MaxFactor, MinFactor, MinFactor, X)

    GetPlayerSpellDamageAgainstPlayer = vitalval * factor
End Function

Public Function GetPlayerSpellDefenseAgainstNPC(ByVal index As Long, ByVal BaseDamage As Long, spellnum As Long) As Long

    Dim MinFactor As Single
    Dim MaxFactor As Single
    Dim factor As Single
    Dim X As Long
    
    
    MinFactor = 0.32
    MaxFactor = 0.64
    X = GetPlayerStat(index, willpower)
    
    'LOS VALORES DE FACTOR OSCILAN ENTRE MINFACTOR Y MAXFACTOR, SIEMPRE ENTRE ESOS DOS VALORES
    'OSCILAN SEGUN: STAT DEL JUGADOR
    factor = Line(MAX_STAT, 0, MaxFactor, MinFactor, MinFactor, X)
    
    'FUNCION DE DEFENSA, EL FACTOR DEBERIA OSCILAR ENTRE 0 Y 1
    '0: EL DAÑO SE VE REDUCIDO AL MÍNIMO (NO SE REDUCE DAÑO BASE)
    '1: EL DAÑO SE VE REDUCIDO AL MÁXIMO (NO HACE DAÑO)
    GetPlayerSpellDefenseAgainstNPC = BaseDamage * factor
End Function

Public Function GetPlayerSpellDefenseAgainstPlayer(ByVal index As Long, ByVal BaseDamage As Long, spellnum As Long) As Long
    Dim Desv As Double
    
    Dim MinFactor As Single
    Dim MaxFactor As Single
    Dim factor As Single
    Dim X As Long
    
    MinFactor = 0.32
    MaxFactor = 0.64
    X = GetPlayerStat(index, willpower)
    
    'LOS VALORES DE FACTOR OSCILAN ENTRE MINFACTOR Y MAXFACTOR, SIEMPRE ENTRE ESOS DOS VALORES
    'OSCILAN SEGUN: STAT DEL JUGADOR
    factor = Line(MAX_STAT, 0, MaxFactor, MinFactor, MinFactor, X)
    
    'FUNCION DE DEFENSA, EL FACTOR DEBERIA OSCILAR ENTRE 0 Y 1
    '0: EL DAÑO SE VE REDUCIDO AL MÍNIMO (NO SE REDUCE DAÑO BASE)
    '1: EL DAÑO SE VE REDUCIDO AL MÁXIMO (NO HACE DAÑO)
    GetPlayerSpellDefenseAgainstPlayer = BaseDamage * factor
End Function

Public Function GetNpcSpellDefense(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal BaseDamage As Long, spellnum As Long) As Long
    If MapNum < 1 Or MapNum > MAX_MAPS Or mapnpcnum < 1 Or mapnpcnum > MAX_MAP_NPCS Then Exit Function
    Dim Desv As Double
    
    Dim MinFactor As Single
    Dim MaxFactor As Single
    Dim factor As Single
    Dim X As Long
    If GetMapPetOwner(MapNum, mapnpcnum) > 0 Then
        MinFactor = 0.5
        MaxFactor = 0.75
    Else
        MinFactor = 0.1
        MaxFactor = 0.6
    End If
    
    X = GetNpcStat(MapNum, mapnpcnum, willpower)
    
    'LOS VALORES DE FACTOR OSCILAN ENTRE MINFACTOR Y MAXFACTOR, SIEMPRE ENTRE ESOS DOS VALORES
    'OSCILAN SEGUN: STAT DEL NPC
    factor = Line(MAX_STAT, 0, MaxFactor, MinFactor, MinFactor, X)
    
    'FUNCION DE DEFENSA, EL FACTOR DEBERIA OSCILAR ENTRE 0 Y 1
    '0: EL DAÑO SE VE REDUCIDO AL MÍNIMO (NO SE REDUCE DAÑO BASE)
    '1: EL DAÑO SE VE REDUCIDO AL MÁXIMO (NO HACE DAÑO)

    GetNpcSpellDefense = BaseDamage * factor
End Function

Public Function GetNPCSpellDamage(ByVal spellnum As Long, ByVal MapNum As Long, ByVal mapnpcnum As Long) As Long
    If MapNum < 1 Or MapNum > MAX_MAPS Or mapnpcnum < 1 Or mapnpcnum > MAX_MAP_NPCS Then Exit Function
    Dim vitalval As Long
    
    vitalval = Spell(spellnum).vital
    
    Dim MinFactor As Single
    Dim MaxFactor As Single
    Dim factor As Single
    Dim X As Long
    
   
    MinFactor = 1.2
    MaxFactor = 2.4
    X = GetNpcStat(MapNum, mapnpcnum, Intelligence)

    'LOS VALORES DE FACTOR OSCILAN ENTRE MINFACTOR Y MAXFACTOR, SIEMPRE ENTRE ESOS DOS VALORES
    'OSCILAN SEGUN: STAT DEL NPC
    factor = Line(MAX_STAT, 0, MaxFactor, MinFactor, MinFactor, X)
    

    GetNPCSpellDamage = vitalval * factor

End Function



Public Sub Impactar(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, ByVal TargetType As Byte)
    Dim spellnum As Long
    
    Dim N As Long
    Dim Auto As Long
    
  ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Then
        'Exit Sub
    End If

    ' Check for weapon
    N = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        N = GetPlayerEquipment(attacker, Weapon)
        spellnum = item(N).Impactar.Spell
        Auto = item(N).Impactar.Auto
    End If
    
   
    'TempPlayer(index).LastSpell = spellnum
    
    If spellnum > 0 Or Auto = 1 Then
        CastSpell2 attacker, spellnum, victim, TargetType
    End If
     
End Sub


Public Sub CastSpell2(ByVal index As Long, ByVal spellnum As Long, ByVal Target As Long, TargetType)
    Dim MapNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim X As Long, Y As Long
   
    Dim buffer As clsBuffer
    Dim SpellCastType As Long
   
    DidCast = False
    


     MapNum = GetPlayerMap(index)
     
      MPCost = GetSpellMPCost(GetPlayerMap(index), index, TARGET_TYPE_PLAYER, spellnum)
      

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        'Call PlayerMsg(index, "¡No tienes suficiente MP!", BrightRed)
        SendActionMsg MapNum, "No MP!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        Exit Sub
    End If
    
    
   
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
   
    ' set the vital
    vital = GetPlayerSpellDamageAgainstPlayer(index, spellnum)
    AoE = Spell(spellnum).AoE
    range = Spell(spellnum).range
   
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_HEALHP
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellPlayer_Effect Vitals.HP, True, index, vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellPlayer_Effect Vitals.MP, True, index, vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SetPlayerDir index, Spell(spellnum).dir
                    PlayerWarpBySpell index, spellnum
                    DidCast = True
                Case SPELL_TYPE_BUFFER
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellStatBuffer index, spellnum
                    DidCast = True
                Case SPELL_TYPE_PROTECT
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellProtect index, spellnum
                    DidCast = True
                Case SPELL_TYPE_CHANGESTATE
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellChangeState index, spellnum
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                X = GetPlayerX(index)
                Y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
                If TargetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
               
                If TargetType = TARGET_TYPE_PLAYER Then
                    X = GetPlayerX(Target)
                    Y = GetPlayerY(Target)
                Else
                    X = MapNpc(MapNum).NPC(Target).X
                    Y = MapNpc(MapNum).NPC(Target).Y
                End If
               
            End If
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If IsinRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, True) Then
                                            SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PlayerAttackPlayer index, i, GetSpellDamage(MapNum, index, TARGET_TYPE_PLAYER, i, TARGET_TYPE_PLAYER, spellnum), spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).NPC(i).Num > 0 Then
                            If MapNpc(MapNum).NPC(i).vital(HP) > 0 Then
                                If IsinRange(AoE, X, Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc index, i, GetSpellDamage(MapNum, index, TARGET_TYPE_PLAYER, i, TARGET_TYPE_NPC, spellnum), spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                   
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If IsinRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect VitalType, increment, i, vital, spellnum
                                    DidCast = True
                                End If
                            End If
                        End If
                    End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).NPC(i).Num > 0 Then
                            If MapNpc(MapNum).NPC(i).vital(HP) > 0 Then
                                If IsinRange(AoE, X, Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                                    SpellNpc_Effect VitalType, increment, i, vital, spellnum, MapNum
                                End If
                            End If
                        End If
                    Next
                    
                Case SPELL_TYPE_PROTECT
                
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If IsinRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellProtect i, spellnum
                                End If
                                End If
                            End If
                        End If
                    Next
            End Select
        Case 2 ' targetted
            If TargetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
           
            If TargetType = TARGET_TYPE_PLAYER Then
                X = GetPlayerX(Target)
                Y = GetPlayerY(Target)
            Else
                X = MapNpc(MapNum).NPC(Target).X
                Y = MapNpc(MapNum).NPC(Target).Y
            End If
               
            If Not IsinRange(range, GetPlayerX(index), GetPlayerY(index), X, Y) Then
                PlayerMsg index, "El objetivo no está al alcance.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
            
            If Spell(spellnum).Type = SPELL_TYPE_DAMAGEHP Then
                If TargetType = TARGET_TYPE_PLAYER And Target = index Then
                    PlayerMsg index, "No puedes atacarte a ti mismo.", BrightRed
                    Exit Sub
                End If
            End If
            
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, True) Then
                                SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                PlayerAttackPlayer index, Target, GetSpellDamage(MapNum, index, TARGET_TYPE_PLAYER, Target, TARGET_TYPE_PLAYER, spellnum), spellnum
                                DidCast = True
                        End If
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                                SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                PlayerAttackNpc index, Target, GetSpellDamage(MapNum, index, TARGET_TYPE_PLAYER, Target, TARGET_TYPE_NPC, spellnum), spellnum
                                DidCast = True
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    End If
                    
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                   
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, vital, spellnum
                                SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                                SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, vital, spellnum, MapNum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, vital, spellnum, MapNum
                        End If
                    End If
                Case SPELL_TYPE_PROTECT
                    SpellProtect Target, spellnum
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
            End Select
    End Select
   
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        'TempPlayer(Index).SpellCD(spellslot) = GetRealTickCount + (Spell(spellnum).CDTime * 1000)
        'Call SendCooldown(Index, spellslot)
        SendActionMsg MapNum, Trim$(Spell(spellnum).TranslatedName) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
    End If
End Sub











