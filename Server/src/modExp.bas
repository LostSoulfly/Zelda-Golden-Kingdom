Attribute VB_Name = "modExp"
Option Explicit
Public NeededNPCS(1 To MAX_LEVELS) As Long
Public NeededExp(1 To MAX_LEVELS) As Long

Private Const NeededNPCSFirstLevel As Long = 5
Private Const Growth As Single = 7 / 5
Private Const Divisions As Long = 4

Private Const LowerExpReductionFactor As Long = 5
Private Const UpperExpReductionFactor As Long = 7

Private Const ExpBase As Long = 30

Function NeededNPCSByLvl(ByVal level As Long) As Long
    If level = 1 Then
        NeededNPCSByLvl = NeededNPCSFirstLevel
    ElseIf level > 1 Then
        NeededNPCSByLvl = NeededNPCS(level - 1) + Growth * (CSng(level) / (MAX_LEVELS \ Divisions) + 1)
    End If
    
    NeededNPCS(level) = NeededNPCSByLvl
End Function

Function LevelExp(ByVal level As Long) As Long

        LevelExp = ExpBase * level * NeededNPCS(level)
        NeededExp(level) = LevelExp
End Function

Sub GenerateExp()
    Dim i As Long
    For i = 1 To MAX_LEVELS
        Call NeededNPCSByLvl(i)
        Call LevelExp(i)
    Next
End Sub

Function GetExpReduction(ByVal leveldifference As Long, ByVal exp As Long) As Long
    If exp <= 0 Then Exit Function
    'returns experience
    Dim Difference As Long
    Difference = Abs(leveldifference)
    
    GetExpReduction = Line(MAX_LEVELS, 0, 100, 0, 0, Difference)
    
    GetExpReduction = exp - exp * (GetExpReduction / 100)
    
End Function

Public Function Line(ByVal MaxX As Variant, ByVal MinX As Variant, ByVal MaxY As Variant, ByVal MinY As Variant, ByVal N As Variant, ByVal X As Variant, Optional ByVal AjustLimits As Boolean = True) As Variant
    Dim VX As Variant
    Dim VY As Variant
    VX = MaxX - MinX
    VY = MaxY - MinY
    
    If VX = 0 Then Exit Function
    
    Dim m As Variant
    m = VY / VX
    Line = m * X + N
    
    If AjustLimits Then
        If m > 0 Then
            If Line > MaxY Then Line = MaxY
            If Line < MinY Then Line = MinY
        Else
            If Line < MaxY Then Line = MaxY
            If Line > MinY Then Line = MinY
        End If
    End If
End Function

Public Sub GivePlayerEXP(ByVal index As Long, ByVal exp As Long)
        ' give the exp
        If exp < 0 Then exp = 0
        If GetPlayerLevel(index) >= MAX_LEVELS Then Exit Sub
        
        Call SetPlayerExp(index, GetPlayerExp(index) + exp)
        SendEXP index
        SendActionMsg GetPlayerMap(index), "+" & exp & " EXP", White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        ' check if we've leveled
        CheckPlayerLevelUp index
End Sub

Function GetPlayerNextLevel(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    'GetPlayerNextLevel = (GetPlayerTriforcesNum(index) + 1) * ((50 / 3) * ((GetPlayerLevel(index) + 1) ^ 3 - (6 * (GetPlayerLevel(index) + 1) ^ 2) + 17 * (GetPlayerLevel(index) + 1)) - 12)
    GetPlayerNextLevel = (GetPlayerTriforcesNum(index) + 1) * LevelExp(GetPlayerLevel(index))
End Function
Sub ComputePlayerExp(ByVal attacker As Long, ByVal attackertype As Byte, ByVal victim As Long, ByVal victimtype As Byte)
    Dim exp As Long
    Dim Share As Collection
    Select Case victimtype
    Case TARGET_TYPE_PLAYER
        Exit Sub 'out of system
        exp = CalculateLosenExp(attacker, victim)
    
        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If
    
        If exp = 0 Then
            Call PlayerMsg(victim, "No perdiste experiencia.", BrightRed)
            Call PlayerMsg(attacker, "No recibiste experiencia.", Cyan)
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - exp)
            SendEXP victim
            Call PlayerMsg(victim, GetTranslation("¡Has perdido", , UnTrimBack) & exp & GetTranslation("de experiencia!", , UnTrimFront), BrightRed, , False)
            
            'Kill Counter
            player(attacker).Kill = player(attacker).Kill + 1
            player(victim).Dead = player(victim).Dead + 1
            
            ' check if we're in a party
            If TempPlayer(attacker).inParty > 0 Then

                Set Share = GetPartyMembersToShareExp(TempPlayer(attacker).inParty, attacker)
                SharePartyMembersExp attacker, Share, exp
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, exp
            End If
        End If
    Case TARGET_TYPE_NPC
        Dim npcnum As Long
        npcnum = MapNpc(GetPlayerMap(attacker)).NPC(victim).Num
        exp = (NPC(npcnum).exp) * Options.ExpMultiplier
        Dim level As Long, Difference As Long
        level = GetNpcLevel(npcnum)
        If attackertype = TARGET_TYPE_NPC Then
            Difference = GetNpcLevel(GetNPCNum(GetPlayerMap(attacker), PlayerHasPetInMap(attacker)), attacker) - level
        Else
            Difference = GetPlayerLevel(attacker) - level
        End If
        If exp > 0 Then
            If TempPlayer(attacker).inParty > 0 Then
                Set Share = GetPartyMembersToShareExp(TempPlayer(attacker).inParty, attacker)
                SharePartyMembersExp attacker, Share, exp, level
            Else
                exp = GetExpReduction(Difference, exp)
                Call ComputeOnlyPlayerExp(attacker, exp)
            End If
        End If
    End Select
End Sub
Sub ComputeOnlyPlayerExp(ByVal index As Long, ByVal exp As Long)
    If TempPlayer(index).TempPet.TempPetSlot > 0 Then
        Call SharePetExp(index, GetPlayerPetSlot(index), exp, TempPlayer(index).TempPet.PetExpPercent)
    Else
        Call GivePlayerEXP(index, exp)
    End If
End Sub

Sub CheckPlayerOutOfExp(ByVal index As Long)
    If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
        SetPlayerExp index, 0
        SendEXP index
    End If
End Sub

Sub CheckPlayerPetsOutOfExp(ByVal index As Long)
    Dim i As Byte
    For i = 1 To MAX_PLAYER_PETS
        With player(index).Pet(i)
            If .Experience >= GetPlayerPetExpByLevel(.NumPet, .level) Then
                .Experience = 0
                SendPetData index, i
            End If
        End With
    Next
End Sub

Public Function GetPartyMembersToShareExp(ByVal partynum As Long, ByVal index As Long) As Collection

    If Party(partynum).MemberCount <= 0 Then Exit Function

    
    Dim Share As Collection
    Set Share = New Collection
    Dim MembersToShare As Byte
    
    Dim i As Long
    For i = 1 To MAX_PARTY_MEMBERS
        Dim tmpIndex As Long
        tmpIndex = Party(partynum).Member(i)
        If IsPlaying(tmpIndex) Then
            If Party(partynum).Member(i) = index Or Not GetLevelDifference(index, Party(partynum).Member(i)) > 10 Then
                Share.Add tmpIndex
            End If
        End If
    Next
    
    Set GetPartyMembersToShareExp = Share
End Function

Public Sub SharePartyMembersExp(ByVal index As Long, ByRef Share As Collection, ByVal TotalExp As Long, Optional ByVal level As Long = 0)
    Dim IndividualExp As Long
    Dim LeftOver As Long
    
    If Share.Count > 0 Then
        IndividualExp = TotalExp \ Share.Count
        LeftOver = TotalExp Mod Share.Count
        
        Dim i As Variant
        Dim exp As Long
        For Each i In Share
            If i = index Then
                exp = IndividualExp + LeftOver
            Else
                exp = IndividualExp
            End If
            
            If level > 0 Then
                exp = GetExpReduction(GetPlayerLevel(i) - level, exp)
            End If
            ComputeOnlyPlayerExp i, exp
        Next
    
    End If

End Sub
