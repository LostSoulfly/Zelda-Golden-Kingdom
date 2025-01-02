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
            SendActionMsg GetPlayerMap(victim), "Dazed", Cyan, TARGET_TYPE_PLAYER, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
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
'0: EL DA�O SE VE REDUCIDO AL M�NIMO (NO SE REDUCE DA�O BASE)
'1: EL DA�O SE VE REDUCIDO AL M�XIMO (NO HACE DA�O)
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
'0: EL DA�O SE VE REDUCIDO AL M�NIMO (NO SE REDUCE DA�O BASE)
'1: EL DA�O SE VE REDUCIDO AL M�XIMO (NO HACE DA�O)

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

Function GetNpcMaxVital(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal vital As Vitals) As Long
    Dim X As Long

    Dim PetOwner As Long, npcnum As Long
    'Prevent Pet System
    PetOwner = GetMapPetOwner(mapnum, mapnpcnum)
    npcnum = GetNPCNum(mapnum, mapnpcnum)
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
                GetNpcMaxVital = NPC(npcnum).HP + ((GetNpcLevel(npcnum, PetOwner) / 2) + (GetNpcStat(mapnum, mapnpcnum, Endurance) / 2) * 10)
            Case MP
                GetNpcMaxVital = 30 + ((GetNpcLevel(npcnum, PetOwner) / 2) + (GetNpcStat(mapnum, mapnpcnum, Intelligence) / 2)) * 10
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

Function GetNpcVitalRegen(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal vital As Vitals) As Long
    Dim i As Long

    
    Dim PetOwner As Long
    'Prevent Pet System
    PetOwner = GetMapPetOwner(mapnum, mapnpcnum)
    
    Select Case PetOwner
    
    Case Is > 0
        i = GetNpcMaxVital(mapnum, mapnpcnum, vital) * GetVitalRegenPercent(GetNpcStat(mapnum, mapnpcnum, willpower, False))
    Case 0
    
        Select Case vital
            Case HP
                i = (GetNPCBaseStat(GetNPCNum(mapnum, mapnpcnum), willpower) * 0.8) + 6
            Case MP
                i = (GetNPCBaseStat(GetNPCNum(mapnum, mapnpcnum), willpower) / 4) + 12.5
        End Select
    
    End Select
    
    GetNpcVitalRegen = i

End Function

Function GetNpcDamage(ByVal mapnum As Long, ByVal mapnpcnum As Long) As Long
    Dim MinFactor As Single
    Dim MaxFactor As Single
    Dim npcnum As Long
    
    npcnum = GetNPCNum(mapnum, mapnpcnum)
    If npcnum < 1 Or npcnum > MAX_NPCS Then Exit Function
    
    If GetMapPetOwner(mapnum, mapnpcnum) > 0 Then
        MinFactor = 2
        MaxFactor = 3
    Else
        MinFactor = 1
        MaxFactor = 2
    End If
    Dim factor As Double
    factor = ((MaxFactor - MinFactor) / MAX_STAT * GetNpcStat(mapnum, mapnpcnum, Strength) + MinFactor)
    
    
    GetNpcDamage = NPC(npcnum).Damage * factor
End Function

Function GetNpcDefense(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal BaseDamage As Long) As Long
    Dim MinFactor As Single
    Dim MaxFactor As Single
    
    If GetMapPetOwner(mapnum, mapnpcnum) > 0 Then
        MinFactor = 0.4
        MaxFactor = 0.8
    Else
        MinFactor = 0.2
        MaxFactor = 0.5
    End If
    
    Dim factor As Double
    factor = Line(MAX_STAT, 0, MaxFactor, MinFactor, MinFactor, GetNpcStat(mapnum, mapnpcnum, Endurance))
    GetNpcDefense = BaseDamage * factor
    
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

'Public Function CanPlayerBlock(ByVal index As Long) As Boolean
'Dim rate As Long
'Dim rndNum As Long

'    CanPlayerBlock = False

'    rate = 0
    ' TODO : make it based on shield lulz
'End Function

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

'Public Function CanNpcBlock(ByVal npcnum As Long) As Boolean
'Dim rate As Long
'Dim rndNum As Long

'    CanNpcBlock = False

'    rate = 0
    ' TODO : make it based on shield lol
'End Function

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
Dim mapnum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(index, mapnpcnum) Then
    
        mapnum = GetPlayerMap(index)
        npcnum = MapNpc(mapnum).NPC(mapnpcnum).Num
        
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(index, npcnum) Then
            SendActionMsg mapnum, "Dodged!", Pink, 1, (MapNpc(mapnum).NPC(mapnpcnum).X * 32), (MapNpc(mapnum).NPC(mapnpcnum).Y * 32)
            Exit Sub
        End If
        'Dim StunTime As Long
        'StunTime = CanStunNpc(index, mapnum, mapnpcnum)
        'If StunTime > 0 Then
            'Call StunNPCByTime(mapnum, mapnpcnum, StunTime)
            'SendActionMsg mapnum, "Stunned!", Pink, 1, (mapnpc(mapnum).NPC(mapnpcnum).X * 32), (mapnpc(mapnum).NPC(mapnpcnum).Y * 32)
            'Exit Sub
        'End If

        ' Get the damage we can do
        Damage = GetPlayerDamageAgainstNPC(index)
        
        Damage = Damage - GetNpcDefense(mapnum, mapnpcnum, Damage)
        ' take away armour
        ' randomise from half to max hit
        Damage = RAND(Damage / 2, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCriticalHit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
            SendSoundToMap mapnum, GetPlayerX(index), GetPlayerY(index), SoundEntity.seCritical, GetPlayerClass(index)
        Else
            SendSoundToMap mapnum, GetPlayerX(index), GetPlayerY(index), SoundEntity.seAttack, GetPlayerClass(index)
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(index, mapnpcnum, Damage)
        Else
            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal mapnpcnum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapnum As Long
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

    mapnum = GetPlayerMap(attacker)
    npcnum = MapNpc(mapnum).NPC(mapnpcnum).Num
    
    'Pet check
    If IsMapNPCaPet(mapnum, mapnpcnum) Then
        If Not (CanPlayerAttackPet(mapnum, mapnpcnum, attacker)) Then
            Exit Function
        End If
    End If
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.HP) <= 0 Then
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
                    NPCX = MapNpc(mapnum).NPC(mapnpcnum).X
                    NPCY = MapNpc(mapnum).NPC(mapnpcnum).Y + 1
                Case DIR_DOWN
                    NPCX = MapNpc(mapnum).NPC(mapnpcnum).X
                    NPCY = MapNpc(mapnum).NPC(mapnpcnum).Y - 1
                Case DIR_LEFT
                    NPCX = MapNpc(mapnum).NPC(mapnpcnum).X + 1
                    NPCY = MapNpc(mapnum).NPC(mapnpcnum).Y
                Case DIR_RIGHT
                    NPCX = MapNpc(mapnum).NPC(mapnpcnum).X - 1
                    NPCY = MapNpc(mapnum).NPC(mapnpcnum).Y
            End Select

            If NPCX = GetPlayerX(attacker) Then
                If NPCY = GetPlayerY(attacker) Then
                    If CanNPCBeAttacked(npcnum) Then
                        CanPlayerAttackNpc = True
                    Else
                        'ALATAR
                        If NPC(npcnum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(npcnum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                            
                            If Len(Trim$(NPC(npcnum).AttackSay)) > 0 Then
                                PlayerMsg attacker, Trim$(NPC(npcnum).Name) & ": " & NPC(npcnum).AttackSay, White
                                'Call SendActionMsg(mapnum, Trim$(NPC(npcnum).Name) & ": " & Trim$(NPC(npcnum).AttackSay), SayColor, 1, mapnpc(mapnum).NPC(mapnpcnum).X * 32, mapnpc(mapnum).NPC(mapnpcnum).Y * 32)
                                Call SpeechWindow(attacker, NPC(npcnum).AttackSay, npcnum)
                            End If
                            
                            SendMapSound (attacker), GetPlayerX(attacker), GetPlayerY(attacker), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapnpcnum).Num
                            
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
                                    QuestMessage attacker, NPC(npcnum).questnum, Quest(NPC(npcnum).questnum).Speech(1), NPC(npcnum).questnum
                                    Exit Function
                                End If
                                If QuestInProgress(attacker, NPC(npcnum).questnum) Then
                                    'if the quest is in progress show the meanwhile message (speech2)
                                    QuestMessage attacker, NPC(npcnum).questnum, Quest(NPC(npcnum).questnum).Speech(2), 0
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
    Dim mapnum As Long
    Dim npcnum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(attacker)
    npcnum = MapNpc(mapnum).NPC(mapnpcnum).Num
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

    If Damage >= MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.HP) Then
        SendActionMsg GetPlayerMap(attacker), "-" & MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.HP), BrightRed, 1, (MapNpc(mapnum).NPC(mapnpcnum).X * 32), (MapNpc(mapnum).NPC(mapnpcnum).Y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y
        
        'Kill counter
        player(attacker).NpcKill = player(attacker).NpcKill + 1
        
        ' send the sound
        If spellnum > 0 Then SendMapSound attacker, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y, SoundEntity.seSpell, spellnum
        
        ' send animation
            If N > 0 Then
                If Not overTime Then
                    If spellnum = 0 Then
                        Call SendAnimation(mapnum, item(GetPlayerEquipment(attacker, Weapon)).Animation, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y)
                    Else
                        Call SendAnimation(mapnum, Spell(spellnum).SpellAnim, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y)
                    End If
                End If
            End If

        'Pet ?
        Call ComputePlayerExp(attacker, TARGET_TYPE_PLAYER, mapnpcnum, TARGET_TYPE_NPC)
        
        'Auto Targetting
        If TempPlayer(attacker).TempPet.PetHasOwnTarget = mapnpcnum Then
            'Objective Finished
            TempPlayer(attacker).TempPet.PetHasOwnTarget = 0
            PetFollowOwner attacker
        End If
        
        'begin of the new system
        'Call the calculator for deciding which item has to be chanced
        DropNum = CalculateDropChances(npcnum)
        'Drop the goods if they get it
        If DropNum > 0 Then
            N = Int(Rnd * NPC(npcnum).Drops(DropNum).DropChance) + 1

            'Drop the selected item
            If N = 1 Then
                Call SpawnItem(NPC(npcnum).Drops(DropNum).DropItem, NPC(npcnum).Drops(DropNum).DropItemValue, mapnum, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y)
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
        
        Call KillNpc(mapnum, mapnpcnum)
        
        Call CheckPlayerPartyTasks(attacker, QUEST_TYPE_GOSLAY, npcnum)
        'ALATAR
        'Call CheckTasks(attacker, QUEST_TYPE_GOSLAY, NPCNum)
        '/ALATAR
        
        'Player NPC info
        player(attacker).NPCKills = player(attacker).NPCKills + 1
        
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.HP) = MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.HP) - Damage
        
        'target the NPC
        TempPlayer(attacker).TargetType = TARGET_TYPE_NPC
        TempPlayer(attacker).Target = mapnpcnum
        SendTarget attacker

        ' Check for a weapon and say damage
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (MapNpc(mapnum).NPC(mapnpcnum).X * 32), (MapNpc(mapnum).NPC(mapnpcnum).Y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound attacker, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y, SoundEntity.seSpell, spellnum
        
        ' send animation
        If N > 0 Then
            If Not overTime Then
                If spellnum = 0 Then Call SendAnimation(mapnum, item(GetPlayerEquipment(attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapnpcnum)
            End If
        End If

        'If the attacker is a pet and does not have a target
        If TempPlayer(attacker).TempPet.TempPetSlot > 0 And TempPlayer(attacker).TempPet.TempPetSlot < MAX_MAP_NPCS And TempPlayer(attacker).TempPet.PetHasOwnTarget = 0 Then
            If TempPlayer(attacker).TempPet.PetState = Assist Then
                MapNpc(mapnum).NPC(TempPlayer(attacker).TempPet.TempPetSlot).TargetType = TARGET_TYPE_NPC
                MapNpc(mapnum).NPC(TempPlayer(attacker).TempPet.TempPetSlot).Target = mapnpcnum
                'Auto Targetting
                TempPlayer(attacker).TempPet.PetHasOwnTarget = mapnpcnum
            End If
        End If

        If MapNpc(mapnum).NPC(mapnpcnum).PetData.Owner > 0 Then
            If Not TempPlayer(MapNpc(mapnum).NPC(mapnpcnum).PetData.Owner).TempPet.PetState = Passive Then
                MapNpc(mapnum).NPC(mapnpcnum).TargetType = TARGET_TYPE_PLAYER ' player
                MapNpc(mapnum).NPC(mapnpcnum).Target = attacker
            End If
        Else
            If MapNpc(mapnum).NPC(mapnpcnum).TargetType = TARGET_TYPE_PLAYER Then
                If Not GetPlayerMap(attacker) = mapnum Then
                    'try and prevent players from kiting?
                    MapNpc(mapnum).NPC(mapnpcnum).TargetType = TARGET_TYPE_PLAYER ' player
                    MapNpc(mapnum).NPC(mapnpcnum).Target = attacker
                End If
            Else
                    MapNpc(mapnum).NPC(mapnpcnum).TargetType = TARGET_TYPE_PLAYER ' player
                    MapNpc(mapnum).NPC(mapnpcnum).Target = attacker
            End If
        End If

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapnum).NPC(mapnpcnum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).NPC(i).Num = MapNpc(mapnum).NPC(mapnpcnum).Num Then
                    MapNpc(mapnum).NPC(i).Target = attacker
                    MapNpc(mapnum).NPC(i).TargetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(mapnum).NPC(mapnpcnum).stopRegen = True
        MapNpc(mapnum).NPC(mapnpcnum).stopRegenTimer = GetRealTickCount
        
        ' if stunning spell, stun the npc
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunNPC mapnpcnum, mapnum, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Npc mapnum, mapnpcnum, spellnum, attacker
            End If
        End If
        
        
     
        
    
        SendMapNpcVitals mapnum, mapnpcnum
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
Dim mapnum As Long, npcnum As Long, blockAmount As Long, Damage As Long
Dim Buffer As clsBuffer
    ' Can the npc attack the player?
    If CanNpcAttackPlayer(mapnpcnum, index) Then
        mapnum = GetPlayerMap(index)
        npcnum = MapNpc(mapnum).NPC(mapnpcnum).Num
                
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapnpcnum).Num
        
        ' Send this packet so they can see the npc attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong ServerPackets.SNpcAttack
        Buffer.WriteLong mapnpcnum
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
        
        MapNpc(mapnum).NPC(mapnpcnum).AttackTimer = GetRealTickCount + GetNPCAttackTimer(mapnum, mapnpcnum)
        
        If TempPlayer(index).TempPet.TempPetSlot > 0 And TempPlayer(index).TempPet.TempPetSlot < MAX_MAP_NPCS And TempPlayer(index).TempPet.PetHasOwnTarget = 0 Then
            If TempPlayer(index).TempPet.PetState <> Passive Then
                MapNpc(mapnum).NPC(TempPlayer(index).TempPet.TempPetSlot).TargetType = TARGET_TYPE_NPC
                MapNpc(mapnum).NPC(TempPlayer(index).TempPet.TempPetSlot).Target = mapnpcnum
                'Auto Targetting
                TempPlayer(index).TempPet.PetHasOwnTarget = mapnpcnum
            End If
        End If
        
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg mapnum, "Dodged!", Pink, 1, (player(index).X * 32), (player(index).Y * 32)
            Exit Sub
        End If
        If CanPlayerParry(index) Then
            SendActionMsg mapnum, "Blocked!", Pink, 1, (player(index).X * 32), (player(index).Y * 32)
            Exit Sub
        End If


        ' Get the damage we can do
        Damage = GetNpcDamage(mapnum, mapnpcnum)
        Damage = Damage - GetPlayerDefenseAgainstNPC(index, Damage)
        Damage = RAND(Damage * 0.8, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (MapNpc(mapnum).NPC(mapnpcnum).X * 32), (MapNpc(mapnum).NPC(mapnpcnum).Y * 32)
        End If

        'Damage = Damage / (GetPlayerDef(index) / 20)
        
        'Damage = Damage * 0.8

        If Damage > 0 Then
            Call NpcAttackPlayer(mapnpcnum, index, Damage)
        Else
            SendActionMsg mapnum, "Evaded!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal mapnpcnum As Long, ByVal index As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapnum As Long
    Dim npcnum As Long
    Dim Buffer As clsBuffer
    
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

    mapnum = GetPlayerMap(index)
    npcnum = MapNpc(mapnum).NPC(mapnpcnum).Num
    
    'Pet check
    If IsMapNPCaPet(mapnum, mapnpcnum) Then
        If Not (CanPetAttackPlayer(mapnum, mapnpcnum, index)) Then
            Exit Function
        End If
    Else
        If (Not CanNPCBehaviourAttack(npcnum)) Then Exit Function
    End If

    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If
    
    'check if the NPC attacking us is actually our pet.
    'We don't want a rebellion on our hands now do we?
    If MapNpc(mapnum).NPC(mapnpcnum).PetData.Owner = index Then Exit Function
    
    'Spell Check
    If IsSpell Then
        CanNpcAttackPlayer = True
    Else
        ' Make sure npcs dont attack more then once a second
        If GetRealTickCount < MapNpc(mapnum).NPC(mapnpcnum).AttackTimer Then
            Exit Function
        End If
        
        ' Make sure they are on the same map
        If IsPlaying(index) Then
            If npcnum > 0 Then

                ' Check if at same coordinates
                If (GetPlayerY(index) + 1 = MapNpc(mapnum).NPC(mapnpcnum).Y) And (GetPlayerX(index) = MapNpc(mapnum).NPC(mapnpcnum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) - 1 = MapNpc(mapnum).NPC(mapnpcnum).Y) And (GetPlayerX(index) = MapNpc(mapnum).NPC(mapnpcnum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(mapnum).NPC(mapnpcnum).Y) And (GetPlayerX(index) + 1 = MapNpc(mapnum).NPC(mapnpcnum).X) Then
                            CanNpcAttackPlayer = True
                        Else
                            If (GetPlayerY(index) = MapNpc(mapnum).NPC(mapnpcnum).Y) And (GetPlayerX(index) - 1 = MapNpc(mapnum).NPC(mapnpcnum).X) Then
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
    Dim mapnum As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(victim)).NPC(mapnpcnum).Num <= 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(victim)
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
    
    'attacker = GetMapPetOwner(mapnum, mapnpcnum)
    If IsMapNPCaPet(mapnum, mapnpcnum) = True Then
        If CheckSafeMode(GetMapPetOwner(mapnum, mapnpcnum), victim) = True Then
        PetFollowOwner GetMapPetOwner(mapnum, mapnpcnum)
        Exit Sub
        End If
    End If
    
    If TempPlayer(victim).TempPet.PetState <> Passive And TempPlayer(victim).TempPet.PetHasOwnTarget = NO Then
        
        PetAttack victim
    End If
    
    ' set the regen timer
    MapNpc(mapnum).NPC(mapnpcnum).stopRegen = True
    MapNpc(mapnum).NPC(mapnpcnum).stopRegenTimer = GetRealTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapnpcnum).Num
        
        'Drop Items If npc was a  pet
        If GetMapPetOwner(mapnum, mapnpcnum) > 0 Then
            If map(mapnum).moral <> MAP_MORAL_ARENA Then
                If (GetLevelDifference(GetMapPetOwner(mapnum, mapnpcnum), victim) <= 10) Then
                    Call PlayerPVPDrops(victim)
                End If
                
                Call SetPlayerJustice(GetMapPetOwner(mapnum, mapnpcnum), victim)
                Call ComputeArmyPvP(GetMapPetOwner(mapnum, mapnpcnum), victim)
            End If
            PetFollowOwner GetMapPetOwner(mapnum, mapnpcnum)
        Else
        
        ' Set NPC target to 0
        MapNpc(mapnum).NPC(mapnpcnum).Target = 0
        MapNpc(mapnum).NPC(mapnpcnum).TargetType = 0
        End If
        
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by a " & Trim$(NPC(GetNPCNum(mapnum, mapnpcnum)).Name) & "!", BrightRed, False)
        ForwardGlobalMsg "[Hub - " & SERVER_NAME & "] " & GetPlayerName(victim) & " has been killed by a " & Trim$(NPC(GetNPCNum(mapnum, mapnpcnum)).Name) & "!"
        Call AddLog(victim, GetPlayerName(victim) & " has been killed by a " & Trim$(NPC(GetNPCNum(mapnum, mapnpcnum)).Name) & "!", PLAYER_LOG)
        
        'Kill Counter
        player(victim).NpcDead = player(victim).NpcDead + 1

    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        Call SendAnimation(mapnum, NPC(MapNpc(GetPlayerMap(victim)).NPC(mapnpcnum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, victim)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapnpcnum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & Damage, Yellow, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetRealTickCount
        
        SendSoundToMap GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seHit, GetPlayerClass(victim)
    End If

End Sub

Sub PetSpellItself(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal SpellSlotNum As Long)
    Dim spellnum As Long
    Dim InitDamage As Long
    Dim DidCast As Boolean
    Dim MPCost As Long
    
    ' Check for subscript out of range
        If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Then
                Exit Sub
        End If

        ' Check for subscript out of range
        If MapNpc(mapnum).NPC(mapnpcnum).Num <= 0 Then
                Exit Sub
        End If
   
        If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub
        
        ' The Variables
        spellnum = NPC(MapNpc(mapnum).NPC(mapnpcnum).Num).Spell(SpellSlotNum)
        
        MPCost = GetSpellMPCost(mapnum, mapnpcnum, TARGET_TYPE_NPC, spellnum)

        ' Check if they have enough MP
        If MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) < MPCost Then
            Exit Sub
        End If
        
        If spellnum = 0 Then
        Exit Sub
        End If
        
        DidCast = False
   
        ' CoolDown Time
        If MapNpc(mapnum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) > GetRealTickCount Then Exit Sub
   
        ' Spell Types
                Select Case Spell(spellnum).Type
                    Case SPELL_TYPE_HEALHP
                        If Not Spell(spellnum).IsAoE Then
                                InitDamage = GetNPCSpellDamage(spellnum, mapnum, mapnpcnum)
                                SpellNpc_Effect Vitals.HP, True, mapnpcnum, InitDamage, spellnum, mapnum
                                DidCast = True
                        End If
                    Case SPELL_TYPE_HEALMP
                        If Not Spell(spellnum).IsAoE Then
                                InitDamage = GetNPCSpellDamage(spellnum, mapnum, mapnpcnum)
                                SpellNpc_Effect Vitals.MP, True, mapnpcnum, InitDamage, spellnum, mapnum
                                DidCast = True
                       End If
                    End Select
                    
                    If DidCast Then
                        MapNpc(mapnum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                        SendNpcAttackAnimation mapnum, mapnpcnum
                        If MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) - MPCost < 0 Then
                            MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) = 0
                        Else
                            MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) = MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) - MPCost
                        End If
                        SendMapNpcVitals mapnum, mapnpcnum
                    Else
                        MapNpc(mapnum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                    End If
End Sub

Sub PetSpellOwner(ByVal mapnpcnum As Long, ByVal index As Long, SpellSlotNum As Long)
    Dim mapnum As Long
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
        mapnum = GetPlayerMap(index)
        spellnum = NPC(MapNpc(mapnum).NPC(mapnpcnum).Num).Spell(SpellSlotNum)
        
        MPCost = GetSpellMPCost(mapnum, mapnpcnum, TARGET_TYPE_NPC, spellnum)

        ' Check if they have enough MP
        If MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) < MPCost Then
            Exit Sub
        End If
        
        If spellnum = 0 Then
        Exit Sub
        End If
        
        DidCast = False
   
        ' CoolDown Time
        If MapNpc(mapnum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) > GetRealTickCount Then Exit Sub
   
        ' Spell Types
                Select Case Spell(spellnum).Type
                    Case SPELL_TYPE_HEALHP
                        If Not Spell(spellnum).IsAoE Then
                            If IsinRange(Spell(spellnum).range + 3, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y, GetPlayerX(index), GetPlayerY(index)) Then
                                InitDamage = GetNPCSpellDamage(spellnum, mapnum, mapnpcnum)
                                SpellPlayer_Effect Vitals.HP, True, index, InitDamage, spellnum
                                DidCast = True
                            End If
                        End If
                    Case SPELL_TYPE_HEALMP
                        If Not Spell(spellnum).IsAoE Then
                            If IsinRange(Spell(spellnum).range + 3, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y, GetPlayerX(index), GetPlayerY(index)) Then
                                InitDamage = GetNPCSpellDamage(spellnum, mapnum, mapnpcnum)
                                SpellPlayer_Effect Vitals.MP, True, index, InitDamage, spellnum
                                DidCast = True
                            End If
                        End If
                    End Select
                    
                    If DidCast Then
                        MapNpc(mapnum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                        SendNpcAttackAnimation mapnum, mapnpcnum
                        If MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) - MPCost < 0 Then
                            MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) = 0
                        Else
                            MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) = MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) - MPCost
                        End If
                        SendMapNpcVitals mapnum, mapnpcnum
                    Else
                        MapNpc(mapnum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                    End If
End Sub
Sub NpcSpellPlayer(ByVal mapnpcnum As Long, ByVal victim As Long, SpellSlotNum As Long)
        Dim mapnum As Long
        Dim i As Long
        Dim N As Long
        Dim spellnum As Long
        Dim Buffer As clsBuffer
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
        mapnum = GetPlayerMap(victim)
        spellnum = NPC(MapNpc(mapnum).NPC(mapnpcnum).Num).Spell(SpellSlotNum)
        
        If spellnum = 0 Then
        Exit Sub
        End If
        
        MPCost = GetSpellMPCost(mapnum, mapnpcnum, TARGET_TYPE_NPC, spellnum)

        ' Check if they have enough MP
        If MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) < MPCost Then
            Exit Sub
        End If
        
   
        ' CoolDown Time
        If MapNpc(mapnum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) > GetRealTickCount Then Exit Sub
   
        ' Spell Types
                Select Case Spell(spellnum).Type
                        ' AOE Healing Spells
                        Case SPELL_TYPE_HEALHP
                        
                        ' Make sure an npc waits for the spell to cooldown
                            If Spell(spellnum).IsAoE Then
                                'For i = 1 To MAX_PLAYERS
                                 GlobalMsg "AOE HealHP is not implemented..", Green, False
                                                
                                'Next
                            Else
                                If Not HasNPCMaxVital(HP, mapnum, mapnpcnum) Then
                                    ' Non AOE Healing Spells
                                    InitDamage = GetNPCSpellDamage(spellnum, mapnum, mapnpcnum)
                                    SpellNpc_Effect Vitals.HP, True, mapnpcnum, InitDamage, spellnum, mapnum
                                    DidCast = True
                                End If
                            End If
                                
                           
                        ' AOE Damaging Spells
                        Case SPELL_TYPE_DAMAGEHP
                        ' Make sure an npc waits for the spell to cooldown
                                If Spell(spellnum).IsAoE Then
                                        For i = 1 To Player_HighIndex
                                            If GetPlayerMap(i) = mapnum Then
                                                If CanNpcAttackPlayer(mapnpcnum, i, True) Then
                                                        If IsinRange(Spell(spellnum).AoE, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y, GetPlayerX(i), GetPlayerY(i)) Then
                                                                InitDamage = GetSpellDamage(mapnum, mapnpcnum, TARGET_TYPE_NPC, i, TARGET_TYPE_PLAYER, spellnum)
                                                                Damage = InitDamage - player(i).stat(Stats.willpower)
                                                                If Damage <= 0 Then
                                                                    SendActionMsg GetPlayerMap(i), "Resisted!", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
                                                                Else
                                                                    If Spell(spellnum).StunDuration > 0 Then
                                                                        CheckSpellStunts victim, spellnum
                                                                    End If
                                                                    SendAnimation mapnum, Spell(spellnum).SpellAnim, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
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
                                        If IsinRange(Spell(spellnum).range, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y, GetPlayerX(victim), GetPlayerY(victim)) Then
                                        InitDamage = GetSpellDamage(mapnum, mapnpcnum, TARGET_TYPE_NPC, victim, TARGET_TYPE_PLAYER, spellnum)
                                        Damage = InitDamage
                                                If Damage <= 0 Then
                                                        SendActionMsg GetPlayerMap(victim), "Resisted!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
                                                Else
                                                    If Spell(spellnum).StunDuration > 0 Then
                                                        CheckSpellStunts victim, spellnum
                                                    End If
                                                    NpcAttackPlayer mapnpcnum, victim, Damage - player(victim).stat(Stats.willpower)
                                                    SendAnimation mapnum, Spell(spellnum).SpellAnim, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim
                                                    DidCast = True
                                                End If
                                        End If
                                    End If
                                End If

                                Case SPELL_TYPE_DAMAGEMP
                                    ' Make sure an npc waits for the spell to cooldown
                                    If Spell(spellnum).IsAoE Then
                                        For i = 1 To Player_HighIndex
                                            If GetPlayerMap(i) = mapnum Then
                                                If CanNpcAttackPlayer(mapnpcnum, victim, True) Then
                                                    If IsinRange(Spell(spellnum).AoE, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y, GetPlayerX(i), GetPlayerY(i)) Then
                    
                                                        Damage = GetSpellDamage(mapnum, mapnpcnum, TARGET_TYPE_NPC, i, TARGET_TYPE_PLAYER, spellnum)
                                                        If Damage <= 0 Then
                                                            SendActionMsg GetPlayerMap(i), "Resisted!", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
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
                                        If IsinRange(Spell(spellnum).range, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y, GetPlayerX(victim), GetPlayerY(victim)) Then
                                            If CanNpcAttackPlayer(mapnpcnum, victim, True) Then
                                                Damage = GetSpellDamage(mapnum, mapnpcnum, TARGET_TYPE_NPC, victim, TARGET_TYPE_PLAYER, spellnum)
                                                If Damage <= 0 Then
                                                    SendActionMsg GetPlayerMap(victim), "Resisted!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
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
                                            If Not HasNPCMaxVital(MP, mapnum, mapnpcnum) Then
                                                ' Non AOE Healing Spells
                                                InitDamage = GetNPCSpellDamage(spellnum, mapnum, mapnpcnum)
                                                SpellNpc_Effect Vitals.MP, True, mapnpcnum, InitDamage, spellnum, mapnum
                                                DidCast = True
                                            End If
                                        End If
                                        
                                    
                                    
                                    End Select
                        
                                    If DidCast Then
                                        MapNpc(mapnum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                                        SendNpcAttackAnimation mapnum, mapnpcnum
                                        If MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) - MPCost < 0 Then
                                            MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) = 0
                                        Else
                                            MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) = MapNpc(mapnum).NPC(mapnpcnum).vital(Vitals.MP) - MPCost
                                        End If
                                        SendMapNpcVitals mapnum, mapnpcnum
                                    Else
                                        MapNpc(mapnum).NPC(mapnpcnum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                                    End If

End Sub

Sub NpcSpellNpc(ByVal mapnum As Long, ByVal aMapNPCNum As Long, ByVal vMapNPCNum As Long, SpellSlotNum As Long)
        Dim i As Long
        Dim N As Long
        Dim spellnum As Long
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
        If MapNpc(mapnum).NPC(aMapNPCNum).Num <= 0 Or MapNpc(mapnum).NPC(vMapNPCNum).Num <= 0 Then
                Exit Sub
        End If
   
        If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub

        ' The Variables
        spellnum = NPC(MapNpc(mapnum).NPC(aMapNPCNum).Num).Spell(SpellSlotNum)
        
        If spellnum = 0 Then
        Exit Sub
        End If
        
        MPCost = GetSpellMPCost(mapnum, aMapNPCNum, TARGET_TYPE_NPC, spellnum)

        ' Check if they have enough MP
        If MapNpc(mapnum).NPC(aMapNPCNum).vital(Vitals.MP) < MPCost Then
            Exit Sub
        End If
        
        'set cast to false
        DidCast = False
   
        ' CoolDown Time
        If MapNpc(mapnum).NPC(aMapNPCNum).SpellTimer(SpellSlotNum) > GetRealTickCount Then Exit Sub
   
        ' Spell Types
                Select Case Spell(spellnum).Type
                        ' AOE Healing Spells
                        Case SPELL_TYPE_HEALHP
                            
                        ' Make sure an npc waits for the spell to cooldown
                            If Spell(spellnum).IsAoE Then
                                For i = 1 To MAX_MAP_NPCS
                                    If MapNpc(mapnum).NPC(i).Num > 0 Then
                                        If MapNpc(mapnum).NPC(i).vital(Vitals.HP) > 0 Then
                                            If IsinRange(Spell(spellnum).AoE, MapNpc(mapnum).NPC(aMapNPCNum).X, MapNpc(mapnum).NPC(aMapNPCNum).Y, MapNpc(mapnum).NPC(i).X, MapNpc(mapnum).NPC(i).Y) Then
                                                InitDamage = GetNPCSpellDamage(spellnum, mapnum, aMapNPCNum)
                                                Select Case GetMapPetOwner(mapnum, aMapNPCNum)
                                                Case 0
                                                    SpellNpc_Effect Vitals.HP, True, i, InitDamage, spellnum, mapnum
                                                Case Is > 0
                                                    If i = aMapNPCNum Then
                                                        If Not HasNPCMaxVital(HP, mapnum, aMapNPCNum) Then
                                                            SpellNpc_Effect Vitals.HP, True, i, InitDamage, spellnum, mapnum
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
                                If Not HasNPCMaxVital(HP, mapnum, aMapNPCNum) Then
                                    InitDamage = GetNPCSpellDamage(spellnum, mapnum, aMapNPCNum)
                                    SpellNpc_Effect Vitals.HP, True, aMapNPCNum, InitDamage, spellnum, mapnum
                                    DidCast = True
                                End If
                            End If
                                
                           
                        ' AOE Damaging Spells
                        Case SPELL_TYPE_DAMAGEHP
                        ' Make sure an npc waits for the spell to cooldown
                            If Spell(spellnum).IsAoE Then
                                For i = 1 To MAX_MAP_NPCS
                                    If CanNpcAttackNpc(mapnum, aMapNPCNum, vMapNPCNum, True) Then
                                        If IsinRange(Spell(spellnum).AoE, MapNpc(mapnum).NPC(aMapNPCNum).X, MapNpc(mapnum).NPC(aMapNPCNum).Y, MapNpc(mapnum).NPC(vMapNPCNum).X, MapNpc(mapnum).NPC(vMapNPCNum).Y) Then
                                            
                                            Damage = GetSpellDamage(mapnum, aMapNPCNum, TARGET_TYPE_NPC, i, TARGET_TYPE_NPC, spellnum)
                                            If Damage <= 0 Then
                                                SendActionMsg mapnum, "Resisted!", Pink, 1, (MapNpc(mapnum).NPC(vMapNPCNum).X) * 32, (MapNpc(mapnum).NPC(vMapNPCNum).Y * 32)
                                            Else
                                                If Spell(spellnum).StunDuration > 0 Then
                                                    StunNPC vMapNPCNum, mapnum, spellnum
                                                End If
                                                SendAnimation mapnum, Spell(spellnum).SpellAnim, MapNpc(mapnum).NPC(vMapNPCNum).X, MapNpc(mapnum).NPC(vMapNPCNum).Y, TARGET_TYPE_NPC, vMapNPCNum
                                                NpcAttackNpc mapnum, aMapNPCNum, vMapNPCNum, Damage
                                                DidCast = True
                                            End If
                                        End If
                                    End If
                                Next
                                
                            ' Non AoE Damaging Spells
                            Else
                                If CanNpcAttackNpc(mapnum, aMapNPCNum, vMapNPCNum, True) Then
                                    If IsinRange(Spell(spellnum).range, MapNpc(mapnum).NPC(aMapNPCNum).X, MapNpc(mapnum).NPC(aMapNPCNum).Y, MapNpc(mapnum).NPC(vMapNPCNum).X, MapNpc(mapnum).NPC(vMapNPCNum).Y) Then
                                        Damage = GetSpellDamage(mapnum, aMapNPCNum, TARGET_TYPE_NPC, vMapNPCNum, TARGET_TYPE_NPC, spellnum)
                                        If Damage <= 0 Then
                                            SendActionMsg mapnum, "Resisted!", Pink, 1, (MapNpc(mapnum).NPC(vMapNPCNum).X) * 32, (MapNpc(mapnum).NPC(vMapNPCNum).Y * 32)
                                        Else
                                            If Spell(spellnum).StunDuration > 0 Then
                                                StunNPC vMapNPCNum, mapnum, spellnum
                                            End If
                                            SendAnimation mapnum, Spell(spellnum).SpellAnim, MapNpc(mapnum).NPC(vMapNPCNum).X, MapNpc(mapnum).NPC(vMapNPCNum).Y, TARGET_TYPE_NPC, vMapNPCNum
                                            NpcAttackNpc mapnum, aMapNPCNum, vMapNPCNum, Damage
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
                                If Not HasNPCMaxVital(MP, mapnum, vMapNPCNum) Then
                                    ' Non AOE Healing Spells
                                    InitDamage = GetNPCSpellDamage(spellnum, mapnum, vMapNPCNum)
                                    SpellNpc_Effect Vitals.MP, True, vMapNPCNum, InitDamage, spellnum, mapnum
                                    DidCast = True
                                End If
                            End If
                        End Select
                        
                        If DidCast Then
                            MapNpc(mapnum).NPC(aMapNPCNum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                            SendNpcAttackAnimation mapnum, aMapNPCNum
                            If MapNpc(mapnum).NPC(aMapNPCNum).vital(Vitals.MP) - MPCost < 0 Then
                                MapNpc(mapnum).NPC(aMapNPCNum).vital(Vitals.MP) = 0
                            Else
                                MapNpc(mapnum).NPC(aMapNPCNum).vital(Vitals.MP) = MapNpc(mapnum).NPC(aMapNPCNum).vital(Vitals.MP) - MPCost
                            End If
                            SendMapNpcVitals mapnum, aMapNPCNum
                        Else
                            MapNpc(mapnum).NPC(aMapNPCNum).SpellTimer(SpellSlotNum) = GetRealTickCount + Spell(spellnum).CDTime * 1000
                        End If
                        
End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim npcnum As Long
Dim Buffer As clsBuffer
Dim mapnum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, victim) Then
    
        'check if players have pets
        'check if the attacker's pet is assist? idk yet. will think.
        'check if victims pet is defensive
        'TempPlayer(index).TempPet.PetState
        
        mapnum = GetPlayerMap(attacker)
     
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim, attacker) Then
            SendActionMsg mapnum, "Dodged!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
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
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
            SendSoundToMap mapnum, GetPlayerX(attacker), GetPlayerY(attacker), SoundEntity.seCritical, GetPlayerClass(attacker)
        Else
            SendSoundToMap mapnum, GetPlayerX(attacker), GetPlayerY(attacker), SoundEntity.seAttack, GetPlayerClass(attacker)
        End If

        
        If Damage > 0 Then
            Call Impactar(attacker, victim, Damage, TempPlayer(attacker).TargetType)
            Call PlayerAttackPlayer(attacker, victim, Damage)
        Else
            Call PlayerMsg(attacker, "Your attack did nothing.", BrightRed)
        End If
    End If
End Sub

Function CheckMapMorals(ByVal mapnum As Long, ByVal attacker As Long, ByVal victim As Long) As Boolean
    ' Check if map is attackable
    Dim moral As Byte
    moral = map(mapnum).moral
    Select Case moral
    Case MAP_MORAL_NONE
        CheckMapMorals = True
    Case MAP_MORAL_SAFE
        If IsPlayerNeutral(victim) Then
            Call PlayerMsg(attacker, "This is a safe zone.", BrightRed)
        Else
            CheckMapMorals = True
        End If
    Case MAP_MORAL_ARENA
        CheckMapMorals = True
    Case MAP_MORAL_PK_SAFE
        If Not IsPlayerNeutral(victim) Then
            PlayerMsg attacker, "This is a PVP zone!", BrightRed
            CheckMapMorals = False
        Else
            CheckMapMorals = True
        End If
    Case MAP_MORAL_PACIFIC
        Call PlayerMsg(attacker, "This is a safe zone.", BrightRed)
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
        If GetPlayerVisible(victim) = 0 Then: Call PlayerMsg(attacker, "Administrators cannot attack other users.", Cyan)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess_Mode(victim) > ADMIN_MONITOR Then
        'If GetPlayerVisible(victim) = 0 Then: Call PlayerMsg(attacker, "You can't attack a" & GetPlayerName(victim) & "!", BrightRed)
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
        'Call PlayerMsg(attacker, "Your levels are too different!", BrightRed)
        'Exit Function
    'End If
    'Make sure the attacker's level isn't too low
    'If GetPlayerLevel(victim) - 10 > GetPlayerLevel(attacker) Then
        'Call PlayerMsg(attacker, "Your levels are too different!", BrightRed)
        'Exit Function
    'End If
    
    'make sure victim is not a guild partner
    If player(attacker).GuildFileId > 1 Then
        If player(attacker).GuildFileId = player(victim).GuildFileId Then
            'Call PlayerMsg(attacker, "�" & GetPlayerName(victim) & "he is a member of your clan!", BrightRed)
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
    If sendmsg Then: Call PlayerMsg(attacker, "You're under level 15, you cannot attack anyone!", BrightRed)
    Exit Function
End If

If GetPlayerLevel(victim) < 15 Then
    If sendmsg Then: Call PlayerMsg(attacker, "Your target is below level 15, you can't attack them!", BrightRed)
    Exit Function
End If

'If GetPlayerLevel(attacker) >= 20 And GetPlayerLevel(victim) < 20 Then
    'If sendmsg Then: Call PlayerMsg(attacker, "If your target is below level 20 and you are above, you cannot attack it!", BrightRed)
    'Exit Function
'End If

'If GetPlayerLevel(attacker) < 20 And GetPlayerLevel(victim) >= 20 Then
    'If sendmsg Then: Call PlayerMsg(attacker, "If your target is above level 20 and you are below, you cannot attack it!", BrightRed)
    'Exit Function
'End If

CheckLevels = True
End Function
Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0)
    Dim exp As Long
    Dim N As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim mapnum As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    N = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        N = GetPlayerEquipment(attacker, Weapon)
    End If
    
    If TempPlayer(attacker).TempPet.TempPetSlot > 0 And TempPlayer(attacker).TempPet.TempPetSlot < MAX_MAP_NPCS And TempPlayer(attacker).TempPet.PetHasOwnTarget = 0 Then
        If TempPlayer(attacker).TempPet.PetState = Assist Then
                'this uses the index of the player to start them attacking!
            PetAttack attacker
        End If
    End If

    If TempPlayer(victim).TempPet.TempPetSlot > 0 And TempPlayer(victim).TempPet.TempPetSlot < MAX_MAP_NPCS And TempPlayer(victim).TempPet.PetHasOwnTarget = 0 Then
        If TempPlayer(victim).TempPet.PetState <> Passive Then
            
        MapNpc(GetPlayerMap(victim)).NPC(TempPlayer(victim).TempPet.TempPetSlot).TargetType = TARGET_TYPE_PLAYER
        MapNpc(GetPlayerMap(victim)).NPC(TempPlayer(victim).TempPet.TempPetSlot).Target = attacker
        TempPlayer(victim).TempPet.PetHasOwnTarget = attacker
            
        End If
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
            Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(attacker), BrightRed, False)
            ' Calculate exp to give attacker
            Call ComputePlayerExp(attacker, TARGET_TYPE_PLAYER, victim, TARGET_TYPE_PLAYER)
            
            
            'ALATAR
            Call CheckTasks(attacker, QUEST_TYPE_GOKILL, victim)
            '/ALATAR
            
            'Only If the level difference is less than 10
            If (GetLevelDifference(attacker, victim) <= 10) Then
                Call PlayerPVPDrops(victim)
            End If
            
            Call SetPlayerJustice(attacker, victim)
            Call ComputeArmyPvP(attacker, victim)
        Else
            Call GlobalMsg(GetPlayerName(attacker) & " has defeated " & GetPlayerName(victim), White, False)
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
        
        TempPlayer(attacker).TempPet.PetHasOwnTarget = 0
        TempPlayer(victim).TempPet.PetHasOwnTarget = 0
        PetFollowOwner attacker
        
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
    Dim mapnum As Long
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
    mapnum = GetPlayerMap(index)
    
    If spellnum <= 0 Or spellnum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > GetRealTickCount Then
        PlayerMsg index, "Ability to reload quickly", BrightRed
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
        'Call PlayerMsg(index, "You don't have enough MP!", BrightRed)
        SendActionMsg mapnum, "No MP!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        Exit Sub
    End If
    
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to use this ability.", BrightRed)
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
        Exit Sub
    End If
    
    AccessReq = Spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess_Mode(index) Then
        Call PlayerMsg(index, "You aren't an admin!.", BrightRed)
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
        Exit Sub
    End If
    
    ClassReq = Spell(spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only a " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
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
                'PlayerMsg index, "You don't have a goal.", BrightRed
            End If
            If TargetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not IsinRange(range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(Target), GetPlayerY(Target)) Then
                    'PlayerMsg index, "Target out of range.", BrightRed
                    SendActionMsg mapnum, "Out of range.", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
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
                If Not IsinRange(range, GetPlayerX(index), GetPlayerY(index), MapNpc(mapnum).NPC(Target).X, MapNpc(mapnum).NPC(Target).Y) Then
                    'PlayerMsg index, "Target out of range.", BrightRed
                    SendActionMsg mapnum, "Out of range.", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
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
        SendAnimation mapnum, Spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, index
        'here sound
        SendActionMsg mapnum, "Casting " & Trim$(Spell(spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
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
    Dim mapnum As Long
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
   
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
   
    DidCast = False
    

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellnum = GetPlayerSpell(index, spellslot)
    TempPlayer(index).LastSpell = spellnum
    mapnum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub

    MPCost = GetSpellMPCost(GetPlayerMap(index), index, TARGET_TYPE_PLAYER, spellnum)

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        'Call PlayerMsg(index, "If you don't have enough MP!", BrightRed)
        SendActionMsg mapnum, "No MP!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        Exit Sub
    End If
   
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to use this ability.", BrightRed)
        Exit Sub
    End If
   
    AccessReq = Spell(spellnum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess_Mode(index) Then
        Call PlayerMsg(index, "You must be an Admin to do that.", BrightRed)
        Exit Sub
    End If
   
    ClassReq = Spell(spellnum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only a " & Trim$((Class(ClassReq).Name)) & " can use this ability!", BrightRed)
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
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellPlayer_Effect Vitals.HP, True, index, vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellPlayer_Effect Vitals.MP, True, index, vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SetPlayerDir index, Spell(spellnum).dir
                    PlayerWarpBySpell index, spellnum
                    DidCast = True
                Case SPELL_TYPE_BUFFER
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellStatBuffer index, spellnum
                    DidCast = True
                Case SPELL_TYPE_PROTECT
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellProtect index, spellnum
                    DidCast = True
                Case SPELL_TYPE_CHANGESTATE
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
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
                    X = MapNpc(mapnum).NPC(Target).X
                    Y = MapNpc(mapnum).NPC(Target).Y
                End If
               
                If Not IsinRange(range, GetPlayerX(index), GetPlayerY(index), X, Y) Then
                    PlayerMsg index, "The goal is not within reach.", BrightRed
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
                                            SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PlayerAttackPlayer index, i, GetSpellDamage(mapnum, index, TARGET_TYPE_PLAYER, i, TARGET_TYPE_PLAYER, spellnum), spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).Num > 0 Then
                            If MapNpc(mapnum).NPC(i).vital(HP) > 0 Then
                                If IsinRange(AoE, X, Y, MapNpc(mapnum).NPC(i).X, MapNpc(mapnum).NPC(i).Y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc index, i, GetSpellDamage(mapnum, index, TARGET_TYPE_PLAYER, i, TARGET_TYPE_NPC, spellnum), spellnum
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
                        If MapNpc(mapnum).NPC(i).Num > 0 Then
                            If MapNpc(mapnum).NPC(i).vital(HP) > 0 Then
                                If IsinRange(AoE, X, Y, MapNpc(mapnum).NPC(i).X, MapNpc(mapnum).NPC(i).Y) Then
                                    SpellNpc_Effect VitalType, increment, i, vital, spellnum, mapnum
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
                X = MapNpc(mapnum).NPC(Target).X
                Y = MapNpc(mapnum).NPC(Target).Y
            End If
               
            If Not IsinRange(range, GetPlayerX(index), GetPlayerY(index), X, Y) Then
                PlayerMsg index, "The goal is not within reach.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
            
            If Spell(spellnum).Type = SPELL_TYPE_DAMAGEHP Then
                If TargetType = TARGET_TYPE_PLAYER And Target = index Then
                    PlayerMsg index, "You can't attack yourself.", BrightRed
                    Exit Sub
                End If
            End If
            
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, True) Then
                                SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                PlayerAttackPlayer index, Target, GetSpellDamage(mapnum, index, TARGET_TYPE_PLAYER, Target, TARGET_TYPE_PLAYER, spellnum), spellnum
                                DidCast = True
                        End If
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                                SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                PlayerAttackNpc index, Target, GetSpellDamage(mapnum, index, TARGET_TYPE_PLAYER, Target, TARGET_TYPE_NPC, spellnum), spellnum
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
                    
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                   
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, vital, spellnum
                                SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                                SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, vital, spellnum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, vital, spellnum, mapnum
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
        SendActionMsg mapnum, Trim$(Spell(spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
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

Public Sub SpellNpc_Effect(ByVal vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellnum As Long, ByVal mapnum As Long)
Dim sSymbol As String * 1
Dim colour As Long
Dim MaxVital As Long
Dim VitalComp As Long

'Do not us this procedeture on hp substracting

        If Damage > 0 Then
        
                MaxVital = GetNpcMaxVital(mapnum, index, vital)
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
                
                If MapNpc(mapnum).NPC(index).vital(vital) = VitalComp Then
                    'Time saver
                    Exit Sub
                ElseIf (MapNpc(mapnum).NPC(index).vital(vital) + Damage >= MaxVital And increment) Or (MapNpc(mapnum).NPC(index).vital(vital) + Damage <= VitalComp And Not increment) Then
                    MapNpc(mapnum).NPC(index).vital(vital) = VitalComp
                Else
                    MapNpc(mapnum).NPC(index).vital(vital) = MapNpc(mapnum).NPC(index).vital(vital) + Damage
                End If
                
                
                SendAnimation mapnum, Spell(spellnum).SpellAnim, MapNpc(mapnum).NPC(index).X, MapNpc(mapnum).NPC(index).Y, TARGET_TYPE_NPC, index
                SendActionMsg mapnum, sSymbol & Damage, colour, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(index).X * 32, MapNpc(mapnum).NPC(index).Y * 32
                
                ' send the sound
                SendMapSound index, MapNpc(mapnum).NPC(index).X, MapNpc(mapnum).NPC(index).Y, SoundEntity.seSpell, spellnum
                
                Call SendMapNpcVitals(mapnum, index)
                
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

Public Sub AddDoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal spellnum As Long, ByVal caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).NPC(index).DoT(i)
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

Public Sub AddHoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).NPC(index).HoT(i)
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

Public Sub HandleDoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal dotNum As Long)
    With MapNpc(mapnum).NPC(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetRealTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.caster, index, True) Then
                    PlayerAttackNpc .caster, index, GetNPCDotDamage(mapnum, index, dotNum)
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

Public Function GetNPCDotDamage(ByVal mapnum As Long, ByVal index As Long, ByVal dotNum As Long) As Long
    'index > 0
    'tempplayer(index).dot(dotnum).spell > 0
    With MapNpc(mapnum).NPC(index).DoT(dotNum)
    Dim SpellDamage As Long
    If .caster > 0 Then
        SpellDamage = GetSpellDamage(mapnum, .caster, TARGET_TYPE_PLAYER, index, TARGET_TYPE_NPC, .Spell)
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

Public Sub HandleHoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal hotNum As Long)
        With MapNpc(mapnum).NPC(index).HoT(hotNum)
                If .Used And .Spell > 0 Then
                        ' time to tick?
                        If GetRealTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                                        SendActionMsg mapnum, "+" & Spell(.Spell).vital, BrightGreen, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(index).X * 32, MapNpc(mapnum).NPC(index).Y * 32
                                        MapNpc(mapnum).NPC(index).vital(Vitals.HP) = MapNpc(mapnum).NPC(index).vital(Vitals.HP) + Spell(.Spell).vital
                                Else
                                        SendActionMsg mapnum, "+" & Spell(.Spell).vital, Cyan, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(index).X * 32, MapNpc(mapnum).NPC(index).Y * 32
                                        MapNpc(mapnum).NPC(index).vital(Vitals.MP) = MapNpc(mapnum).NPC(index).vital(Vitals.MP) + Spell(.Spell).vital
                                        
                                        If MapNpc(mapnum).NPC(index).vital(Vitals.MP) > GetNpcMaxVital(mapnum, index, MP) Then
                                                MapNpc(mapnum).NPC(index).vital(Vitals.MP) = GetNpcMaxVital(mapnum, index, MP)
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
        PlayerMsg index, "You are paralyzed!", BrightRed
    End If
End Sub

Public Sub StunPlayerByTime(ByVal index As Long, ByVal Time As Single)
    If GPE(index) Then Exit Sub
    BlockPlayerAction index, aMove, Time
    BlockPlayerAction index, aAttack, Time
End Sub
Public Sub StunNPC(ByVal index As Long, ByVal mapnum As Long, ByVal spellnum As Long)
Dim npcnum As Long
    ' check if it's a stunning spell
    If Spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(mapnum).NPC(index).StunDuration = Spell(spellnum).StunDuration - NPC(index).level
        MapNpc(mapnum).NPC(index).StunTimer = GetRealTickCount
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

Public Sub TryNpcAttackNpc(ByVal mapnum As Long, ByVal mapNPCnumAttacker As Long, ByVal mapNpcNumVictim As Long)
Dim NPCNumAttacker As Long, NPCNumVictim As Long, blockAmount As Long, Damage As Long

    ' Can the npc attack the player?
    If CanNpcAttackNpc(mapnum, mapNPCnumAttacker, mapNpcNumVictim) Then
    
        NPCNumAttacker = MapNpc(mapnum).NPC(mapNPCnumAttacker).Num
        NPCNumVictim = MapNpc(mapnum).NPC(mapNpcNumVictim).Num
        
        
        ' send the sound
        MapNpc(mapnum).NPC(mapNPCnumAttacker).AttackTimer = GetRealTickCount + GetNPCAttackTimer(mapnum, mapNPCnumAttacker)
        ' send the sound
        SendSoundToMap mapnum, MapNpc(mapnum).NPC(mapNpcNumVictim).X, MapNpc(mapnum).NPC(mapNpcNumVictim).Y, SoundEntity.seNpc, MapNpc(mapnum).NPC(mapNPCnumAttacker).Num

        ' check if NPC can avoid the attack
        If CanNpcDodgeNpc(mapnum, mapNPCnumAttacker, mapNpcNumVictim) Then
            SendActionMsg mapnum, "Dodged!", Pink, 1, (MapNpc(mapnum).NPC(mapNpcNumVictim).X * 32), (MapNpc(mapnum).NPC(mapNpcNumVictim).Y * 32)
            Exit Sub
        End If
        If CanNpcParryNpc(mapnum, mapNPCnumAttacker, mapNpcNumVictim) Then
            SendActionMsg mapnum, "Blocked!", Pink, 1, (MapNpc(mapnum).NPC(mapNpcNumVictim).X * 32), (MapNpc(mapnum).NPC(mapNpcNumVictim).Y * 32)
            Exit Sub
        End If

        Damage = GetNpcDamage(mapnum, mapNPCnumAttacker)
        Damage = Damage - GetNpcDefense(mapnum, mapNpcNumVictim, Damage)
        Damage = RAND(Damage * 0.8, Damage)
        
        If CanNpcCrit(NPCNumAttacker) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (MapNpc(mapnum).NPC(mapNPCnumAttacker).X * 32), (MapNpc(mapnum).NPC(mapNPCnumAttacker).Y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackNpc(mapnum, mapNPCnumAttacker, mapNpcNumVictim, Damage)
        Else
            SendActionMsg mapnum, "Evaded!", Cyan, 1, MapNpc(mapnum).NPC(mapNpcNumVictim).X * 32, MapNpc(mapnum).NPC(mapNpcNumVictim).Y * 32
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
    If MapNpc(mapnum).NPC(attacker).Num <= 0 Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(mapnum).NPC(victim).Num <= 0 Then
        Exit Function
    End If

    aNPCNum = MapNpc(mapnum).NPC(attacker).Num
    vNPCNum = MapNpc(mapnum).NPC(victim).Num
    
    
    If aNPCNum <= 0 Then Exit Function
    If vNPCNum <= 0 Then Exit Function
    
    'Pet check
    If IsMapNPCaPet(mapnum, attacker) And IsMapNPCaPet(mapnum, victim) Then
        If Not (CanPetAttackPet(mapnum, attacker, victim)) Then
            Exit Function
        End If
    End If
    
    'Check npc type
    If CanNPCBeAttacked(vNPCNum) = False Then Exit Function

    ' Make sure the npcs arent already dead
    If MapNpc(mapnum).NPC(attacker).vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(victim).vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    If IsSpell Then
        CanNpcAttackNpc = True
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetRealTickCount < MapNpc(mapnum).NPC(attacker).AttackTimer Then
        Exit Function
    End If
    
    AttackerX = MapNpc(mapnum).NPC(attacker).X
    AttackerY = MapNpc(mapnum).NPC(attacker).Y
    VictimX = MapNpc(mapnum).NPC(victim).X
    VictimY = MapNpc(mapnum).NPC(victim).Y

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
    Dim N As Long
    Dim PetOwner As Long
    Dim DropNum As Integer
    
    If attacker <= 0 Or attacker > MAX_MAP_NPCS Then Exit Sub
    If victim <= 0 Or victim > MAX_MAP_NPCS Then Exit Sub
    
    If Damage <= 0 Then Exit Sub
    
    aNPCNum = MapNpc(mapnum).NPC(attacker).Num
    vNPCNum = MapNpc(mapnum).NPC(victim).Num
    
    If aNPCNum <= 0 Then Exit Sub
    If vNPCNum <= 0 Then Exit Sub
    
    'set the victim's target to the pet attacking it
    If MapNpc(mapnum).NPC(victim).PetData.Owner > 0 Then
        'if the victim is a pet
        If MapNpc(mapnum).NPC(attacker).PetData.Owner > 0 Then
            If Not TempPlayer(MapNpc(mapnum).NPC(victim).PetData.Owner).TempPet.PetHasOwnTarget > 0 Then
                If RAND(1, 2) <> 2 Then 'randomly choose to attack owner or pet.
                'for 2, we'll attack the owner.
                    MapNpc(mapnum).NPC(victim).TargetType = TARGET_TYPE_PLAYER
                    MapNpc(mapnum).NPC(victim).Target = MapNpc(mapnum).NPC(attacker).PetData.Owner
                    TempPlayer(MapNpc(mapnum).NPC(victim).PetData.Owner).TempPet.PetHasOwnTarget = MapNpc(mapnum).NPC(attacker).PetData.Owner
                Else
                    MapNpc(mapnum).NPC(victim).TargetType = 2 'Npc
                    MapNpc(mapnum).NPC(victim).Target = attacker
                    TempPlayer(MapNpc(mapnum).NPC(victim).PetData.Owner).TempPet.PetHasOwnTarget = attacker
                End If
            End If
        Else
        'attacker is not an pet, but the victim is.
            If TempPlayer(MapNpc(mapnum).NPC(victim).PetData.Owner).TempPet.PetHasOwnTarget = 0 Then
                MapNpc(mapnum).NPC(victim).TargetType = 2 'Npc
                MapNpc(mapnum).NPC(victim).Target = attacker
            End If
        End If
    Else
        'victim is not a pet
        MapNpc(mapnum).NPC(victim).TargetType = 2 'Npc
        MapNpc(mapnum).NPC(victim).Target = attacker
    End If

    
    ' set the regen timer
    MapNpc(mapnum).NPC(attacker).stopRegen = True
    MapNpc(mapnum).NPC(victim).stopRegenTimer = GetRealTickCount
    
    ' Send this packet so they can see the person attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong attacker
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing

    If Damage >= MapNpc(mapnum).NPC(victim).vital(Vitals.HP) Then
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (MapNpc(mapnum).NPC(victim).X * 32), (MapNpc(mapnum).NPC(victim).Y * 32)
        SendBlood mapnum, MapNpc(mapnum).NPC(victim).X, MapNpc(mapnum).NPC(victim).Y
        
        ' Set NPC target to 0
        MapNpc(mapnum).NPC(attacker).Target = 0
        MapNpc(mapnum).NPC(attacker).TargetType = 0
        'reset the targetter for the player
        
        'Call the calculator for deciding which item has to be chanced
        DropNum = CalculateDropChances(vNPCNum)
        
        'Drop the goods if they get it
        If DropNum > 0 Then
            N = Int(Rnd * NPC(vNPCNum).Drops(DropNum).DropChance) + 1

            'Drop the selected item
            If N = 1 Then
                Call SpawnItem(NPC(vNPCNum).Drops(DropNum).DropItem, NPC(vNPCNum).Drops(DropNum).DropItemValue, mapnum, MapNpc(mapnum).NPC(victim).X, MapNpc(mapnum).NPC(victim).Y)
            End If
        End If
        
        
        If IsMapNPCaPet(mapnum, attacker) Then
            TempPlayer(MapNpc(mapnum).NPC(attacker).PetData.Owner).Target = 0
            TempPlayer(MapNpc(mapnum).NPC(attacker).PetData.Owner).TargetType = TARGET_TYPE_NONE
            
            PetOwner = MapNpc(mapnum).NPC(attacker).PetData.Owner
            
            SendTarget PetOwner
            
            Call ComputePlayerExp(PetOwner, TARGET_TYPE_PLAYER, victim, TARGET_TYPE_NPC)
            'objective finished
            TempPlayer(PetOwner).TempPet.PetHasOwnTarget = 0
            PetFollowOwner PetOwner
            Call CheckPlayerPartyTasks(PetOwner, QUEST_TYPE_GOSLAY, vNPCNum)
            
        End If
                      
        If IsMapNPCaPet(mapnum, victim) Then
            'Get the pet owners' index
            PetOwner = MapNpc(mapnum).NPC(victim).PetData.Owner
            'Set the NPC's target on the owner now
            MapNpc(mapnum).NPC(attacker).TargetType = 1 'player
            MapNpc(mapnum).NPC(attacker).Target = PetOwner
            
            'objective finished
            TempPlayer(PetOwner).TempPet.PetHasOwnTarget = 0
            'Set Spawn time
            'TempPlayer(PetOwner).TempPet.PetSpawnWait = GetRealTickCount
            'PetDisband PetOwner, mapnum, True
        End If
        
        KillNpc mapnum, victim
        
        If PetOwner > 0 Then
            PetFollowOwner PetOwner
        End If
    Else
        ' npc not dead, just do the damage
        MapNpc(mapnum).NPC(victim).vital(Vitals.HP) = MapNpc(mapnum).NPC(victim).vital(Vitals.HP) - Damage
       
        ' Say damage
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (MapNpc(mapnum).NPC(victim).X * 32), (MapNpc(mapnum).NPC(victim).Y * 32)
        SendBlood mapnum, MapNpc(mapnum).NPC(victim).X, MapNpc(mapnum).NPC(victim).Y
        
        ' set the regen timer
        MapNpc(mapnum).NPC(victim).stopRegen = True
        MapNpc(mapnum).NPC(victim).stopRegenTimer = GetRealTickCount
    End If
    
    'Send both Npc's Vitals to the client
    SendMapNpcVitals mapnum, attacker
    SendMapNpcVitals mapnum, victim

End Sub

Function CalculateLosenExp(ByVal attacker As Long, ByVal victim As Long) As Long

Dim Difference As Long

Difference = GetLevelDifference(attacker, victim)

Dim reduction As Double
reduction = Line(MAX_LEVELS, 0, 10, 4, 4, Difference)

'can't run time 6
CalculateLosenExp = CDbl(GetPlayerExp(victim)) / reduction

End Function



Public Function CanNpcDodgeNpc(ByVal mapnum As Long, ByVal attacker As Long, ByVal victim As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodgeNpc = False
    rate = (GetNpcStat(mapnum, victim, Agility) - GetNpcStat(mapnum, attacker, Agility))
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
    rate = (GetNpcStat(mapnum, attacker, Strength) - GetNpcStat(mapnum, victim, Strength))
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
Public Function CanPetAttackPlayer(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal victim As Long) As Boolean
    CanPetAttackPlayer = False
    
    'Check if pet
    If Not IsMapNPCaPet(mapnum, mapnpcnum) Then Exit Function
    
    Dim Owner As Long
    Owner = GetMapPetOwner(mapnum, mapnpcnum)
    ' Check if map is attackable
    
    'If MapNpc(mapnum).NPC(mapnpcnum).PetData.Owner = Owner Then Exit Function
    If TempPlayer(victim).TempPet.TempPetSlot = mapnpcnum Then Exit Function
    
    If Not CheckMapMorals(mapnum, Owner, victim) Then
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

Public Function CanPlayerAttackPet(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal attacker As Long) As Boolean
    CanPlayerAttackPet = False
    
    'Check if pet
    If Not IsMapNPCaPet(mapnum, mapnpcnum) Then Exit Function
    
    Dim Owner As Long
    Owner = GetMapPetOwner(mapnum, mapnpcnum)
    
    ' Check if map is attackable
    If Not CheckMapMorals(mapnum, attacker, Owner) Then
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
    
    'If CanPlayerAttackByJustice(attacker, Owner, False) Then
    '    Exit Function
    'End If
    
    
    
    
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

Public Function CanPetAttackPet(ByVal mapnum As Long, ByVal aMapNPCNum As Long, ByVal vMapNPCNum As Long) As Boolean
    CanPetAttackPet = False
    
    'Check if pet
    If Not IsMapNPCaPet(mapnum, aMapNPCNum) Then Exit Function
    If Not IsMapNPCaPet(mapnum, vMapNPCNum) Then Exit Function
    
    Dim aOwner As Long
    Dim vOwner As Long
    
    aOwner = GetMapPetOwner(mapnum, aMapNPCNum)
    vOwner = GetMapPetOwner(mapnum, vMapNPCNum)
    
    ' Check if map is attackable
    If Not map(mapnum).moral = MAP_MORAL_NONE Then
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
Public Function GetSpellDamage(ByVal mapnum As Long, ByVal attacker As Long, ByVal attackertype As Byte, ByVal victim As Long, ByVal victimtype As Byte, ByVal spellnum As Long) As Long
    Dim Damage As Long
    Dim Protection As Long
    If mapnum < 1 Or mapnum > MAX_MAPS Or attacker < 1 Or victim < 1 Or spellnum < 1 Then Exit Function
    
    If attackertype = TARGET_TYPE_PLAYER And victimtype = TARGET_TYPE_NPC Then
    
        Damage = GetPlayerSpellDamageAgainstNPC(attacker, spellnum)
        Damage = Damage - GetNpcSpellDefense(mapnum, victim, Damage, spellnum)
        
    ElseIf attackertype = TARGET_TYPE_PLAYER And victimtype = TARGET_TYPE_PLAYER Then
    
        Damage = GetPlayerSpellDamageAgainstPlayer(attacker, spellnum)
        Damage = Damage - GetPlayerSpellDefenseAgainstPlayer(victim, Damage, spellnum)
        
    ElseIf attackertype = TARGET_TYPE_NPC And victimtype = TARGET_TYPE_PLAYER Then
    
        Damage = GetNPCSpellDamage(spellnum, mapnum, attacker)
        Damage = Damage - GetPlayerSpellDefenseAgainstNPC(victim, Damage, spellnum)
        
    ElseIf attackertype = TARGET_TYPE_NPC And victimtype = TARGET_TYPE_NPC Then
    
        Damage = GetNPCSpellDamage(spellnum, mapnum, attacker)
        Damage = Damage - GetNpcSpellDefense(mapnum, victim, Damage, spellnum)
        
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
    '0: EL DA�O SE VE REDUCIDO AL M�NIMO (NO SE REDUCE DA�O BASE)
    '1: EL DA�O SE VE REDUCIDO AL M�XIMO (NO HACE DA�O)
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
    '0: EL DA�O SE VE REDUCIDO AL M�NIMO (NO SE REDUCE DA�O BASE)
    '1: EL DA�O SE VE REDUCIDO AL M�XIMO (NO HACE DA�O)
    GetPlayerSpellDefenseAgainstPlayer = BaseDamage * factor
End Function

Public Function GetNpcSpellDefense(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal BaseDamage As Long, spellnum As Long) As Long
    If mapnum < 1 Or mapnum > MAX_MAPS Or mapnpcnum < 1 Or mapnpcnum > MAX_MAP_NPCS Then Exit Function
    Dim Desv As Double
    
    Dim MinFactor As Single
    Dim MaxFactor As Single
    Dim factor As Single
    Dim X As Long
    If GetMapPetOwner(mapnum, mapnpcnum) > 0 Then
        MinFactor = 0.5
        MaxFactor = 0.75
    Else
        MinFactor = 0.1
        MaxFactor = 0.6
    End If
    
    X = GetNpcStat(mapnum, mapnpcnum, willpower)
    
    'LOS VALORES DE FACTOR OSCILAN ENTRE MINFACTOR Y MAXFACTOR, SIEMPRE ENTRE ESOS DOS VALORES
    'OSCILAN SEGUN: STAT DEL NPC
    factor = Line(MAX_STAT, 0, MaxFactor, MinFactor, MinFactor, X)
    
    'FUNCION DE DEFENSA, EL FACTOR DEBERIA OSCILAR ENTRE 0 Y 1
    '0: EL DA�O SE VE REDUCIDO AL M�NIMO (NO SE REDUCE DA�O BASE)
    '1: EL DA�O SE VE REDUCIDO AL M�XIMO (NO HACE DA�O)

    GetNpcSpellDefense = BaseDamage * factor
End Function

Public Function GetNPCSpellDamage(ByVal spellnum As Long, ByVal mapnum As Long, ByVal mapnpcnum As Long) As Long
    If mapnum < 1 Or mapnum > MAX_MAPS Or mapnpcnum < 1 Or mapnpcnum > MAX_MAP_NPCS Then Exit Function
    Dim vitalval As Long
    
    vitalval = Spell(spellnum).vital
    
    Dim MinFactor As Single
    Dim MaxFactor As Single
    Dim factor As Single
    Dim X As Long
    
   
    MinFactor = 1.2
    MaxFactor = 2.4
    X = GetNpcStat(mapnum, mapnpcnum, Intelligence)

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
    Dim mapnum As Long
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
   
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
   
    DidCast = False
    


     mapnum = GetPlayerMap(index)
     
      MPCost = GetSpellMPCost(GetPlayerMap(index), index, TARGET_TYPE_PLAYER, spellnum)
      

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        'Call PlayerMsg(index, "If you don't have enough MP!", BrightRed)
        SendActionMsg mapnum, "No MP!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
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
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellPlayer_Effect Vitals.HP, True, index, vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellPlayer_Effect Vitals.MP, True, index, vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SetPlayerDir index, Spell(spellnum).dir
                    PlayerWarpBySpell index, spellnum
                    DidCast = True
                Case SPELL_TYPE_BUFFER
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellStatBuffer index, spellnum
                    DidCast = True
                Case SPELL_TYPE_PROTECT
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellProtect index, spellnum
                    DidCast = True
                Case SPELL_TYPE_CHANGESTATE
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
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
                    X = MapNpc(mapnum).NPC(Target).X
                    Y = MapNpc(mapnum).NPC(Target).Y
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
                                            SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PlayerAttackPlayer index, i, GetSpellDamage(mapnum, index, TARGET_TYPE_PLAYER, i, TARGET_TYPE_PLAYER, spellnum), spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).Num > 0 Then
                            If MapNpc(mapnum).NPC(i).vital(HP) > 0 Then
                                If IsinRange(AoE, X, Y, MapNpc(mapnum).NPC(i).X, MapNpc(mapnum).NPC(i).Y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc index, i, GetSpellDamage(mapnum, index, TARGET_TYPE_PLAYER, i, TARGET_TYPE_NPC, spellnum), spellnum
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
                        If MapNpc(mapnum).NPC(i).Num > 0 Then
                            If MapNpc(mapnum).NPC(i).vital(HP) > 0 Then
                                If IsinRange(AoE, X, Y, MapNpc(mapnum).NPC(i).X, MapNpc(mapnum).NPC(i).Y) Then
                                    SpellNpc_Effect VitalType, increment, i, vital, spellnum, mapnum
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
                X = MapNpc(mapnum).NPC(Target).X
                Y = MapNpc(mapnum).NPC(Target).Y
            End If
               
            If Not IsinRange(range, GetPlayerX(index), GetPlayerY(index), X, Y) Then
                PlayerMsg index, "The goal is not within reach.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
            
            If Spell(spellnum).Type = SPELL_TYPE_DAMAGEHP Then
                If TargetType = TARGET_TYPE_PLAYER And Target = index Then
                    PlayerMsg index, "You can't attack yourself.", BrightRed
                    Exit Sub
                End If
            End If
            
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, True) Then
                                SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                PlayerAttackPlayer index, Target, GetSpellDamage(mapnum, index, TARGET_TYPE_PLAYER, Target, TARGET_TYPE_PLAYER, spellnum), spellnum
                                DidCast = True
                        End If
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                                SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                PlayerAttackNpc index, Target, GetSpellDamage(mapnum, index, TARGET_TYPE_PLAYER, Target, TARGET_TYPE_NPC, spellnum), spellnum
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
                    
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                   
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, vital, spellnum
                                SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                                SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, vital, spellnum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, vital, spellnum, mapnum
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
        SendActionMsg mapnum, Trim$(Spell(spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
    End If
End Sub

















