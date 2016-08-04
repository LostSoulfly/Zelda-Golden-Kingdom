Attribute VB_Name = "modArmy"
Option Explicit

Public Enum HeroRangesType
    Soldado = 1
    Escolta
    Teniente
    Capitan
    Protector
    Caballero
    HeroRangesTypeCount
End Enum

Public Enum PKRangesType
    Mercenario = 1
    Aniquilador
    Devastador
    Asolador
    Comandante
    Elite
    PkRangesTypeCount
End Enum

Private Const MAX_HERO_RANGE_POINTS As Long = 600
Private Const MAX_PK_RANGE_POINTS As Long = 600


Public Const MAX_FREE_KILL_POINTS As Long = 10
Public Const KILL_POINTS_BY_TIME As Single = 1
Public Const KILL_POINTS_LAPSE As Long = 300 ' 5 minutes
Public Const NEUTRAL_POINTS_LAPSE As Long = 600 ' 10 minutes
'Public Const KILL_POINTS_LAPSE As Long = 5 ' 5 minutes
'Public Const NEUTRAL_POINTS_LAPSE As Long = 10 ' 10 minutes

Public KillPointsTimer As Long
Public NeutralPlayerPointsTimer As Long

Sub SetPlayerKillPoints(ByVal index As Long, ByVal points As Single, ByVal Status As Byte)
    If index = 0 Then Exit Sub
    Select Case Status
    Case PK_PLAYER
        player(index).PKKillPoints = points
    Case HERO_PLAYER
        player(index).HeroKillPoints = points
    End Select
End Sub

Function GetPlayerKillPoints(ByVal index As Long, ByVal Status As Byte) As Single
    If index = 0 Then Exit Function
    Select Case Status
    Case PK_PLAYER
        GetPlayerKillPoints = player(index).PKKillPoints
    Case HERO_PLAYER
        GetPlayerKillPoints = player(index).HeroKillPoints
    End Select
End Function

Function GetPlayerRangePoints(ByVal index As Long, ByVal Status As Byte) As Long
    If index = 0 Then Exit Function
    Select Case Status
    Case PK_PLAYER
        GetPlayerRangePoints = player(index).PKPoints
    Case HERO_PLAYER
        GetPlayerRangePoints = player(index).HeroPoints
    End Select
End Function


Sub AddPlayerPointsByTimeLapse(ByVal index As Long)
    If index = 0 Then Exit Sub
    Dim Status As Byte
    Status = GetPlayerPK(index)
    If Status = NONE_PLAYER Then Exit Sub


    If GetPlayerKillPoints(index, Status) < GetPlayerArmyRange(index) Then
        Call SetPlayerKillPoints(index, GetPlayerKillPoints(index, Status) + KILL_POINTS_BY_TIME, Status)
    End If

End Sub

Sub AddNeutralPlayerPoints(ByVal index As Long)
    If index = 0 Then Exit Sub
    Dim Status As Byte
    Status = GetPlayerPK(index)
    
    If Status <> NONE_PLAYER Then Exit Sub
    
    SetNeutral index, True
End Sub


Sub ResetPlayerArmy(ByVal index As Long, ByVal Status As Byte)
    If index = 0 Then Exit Sub

    SetPlayerJusticePoints index, Status, 0
    SetPlayerKillPoints index, 0, Status

End Sub


Function GetPlayerArmyRange(ByVal index As Long) As Byte
    Dim Status As Byte
    Status = GetPlayerPK(index)
    If Status = NONE_PLAYER Then Exit Function
    
    Dim MaxRangePoints As Long, MinRangePoints As Long
    MaxRangePoints = GetArmyMaxRangePoints(Status)
    MinRangePoints = 0
    
    Dim Temp As Long
    Temp = GetArmyMaxRange(Status) / (MaxRangePoints - MinRangePoints) * GetPlayerJusticePoints(index, Status) + MinRangePoints
    
    If Temp > GetArmyMaxRange(Status) Then Temp = GetArmyMaxRange(Status)
    If Temp < 1 Then Temp = 1
    
    GetPlayerArmyRange = Temp
End Function

Function GetArmyMaxRangePoints(ByVal army As Byte)
    Select Case army
    Case HERO_PLAYER
        GetArmyMaxRangePoints = MAX_HERO_RANGE_POINTS
    Case PK_PLAYER
        GetArmyMaxRangePoints = MAX_PK_RANGE_POINTS
    End Select
End Function

Function AreArmyRivals(ByVal army1 As Byte, ByVal army2 As Byte) As Boolean
    If army1 = HERO_PLAYER And army2 = PK_PLAYER Then
        AreArmyRivals = True
    ElseIf army1 = PK_PLAYER And army2 = HERO_PLAYER Then
        AreArmyRivals = True
    End If
End Function

Sub PartyShareKillPoints(ByVal index As Long, ByVal partynum As Byte, ByVal points As Single)
    If partynum < 1 Or partynum > MAX_PARTYS Then Exit Sub
    Dim i As Long
    
    With Party(partynum)
    
    Dim ShareVect(1 To MAX_PARTY_MEMBERS) As Boolean
    Dim MembersToShare As Byte
    Dim MemberIndex As Long
    Dim PartyStatus As Byte
    PartyStatus = GetPlayerPK(index)
    
    If PartyStatus = NONE_PLAYER Then Exit Sub
    
    For i = 1 To MAX_PARTY_MEMBERS
        MemberIndex = .Member(i)
        If IsPlaying(MemberIndex) Then
            
            If MemberIndex = index Then
                ShareVect(i) = True
                MembersToShare = MembersToShare + 1
            Else
                If GetPlayerMap(index) = GetPlayerMap(MemberIndex) And GetLevelDifference(index, MemberIndex) <= 10 And GetPlayerPK(MemberIndex) = PartyStatus Then
                    ShareVect(i) = True
                    MembersToShare = MembersToShare + 1
                Else
                    ShareVect(i) = False
                End If
            End If
        Else
            ShareVect(i) = False
        End If
    Next
    
    If MembersToShare = 0 Then Exit Sub
    
    ' find out the equal share
    points = points / MembersToShare
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        MemberIndex = Party(partynum).Member(i)
        If ShareVect(i) Then
            SetPlayerKillPoints MemberIndex, GetPlayerKillPoints(MemberIndex, PartyStatus) + points, PartyStatus
            PlayerMsg MemberIndex, GetTranslation("Has ganado:", , UnTrimBack) & points & GetTranslation("puntos de", , UnTrimBoth) & JusticeToStr(PartyStatus) & GetTranslation("ahora tienes:", , UnTrimBoth) & GetPlayerKillPoints(MemberIndex, PartyStatus), GetColorByJustice(PartyStatus), , False
            SendKillPoints MemberIndex
        End If
    Next
    
    End With

End Sub

Sub ComputeGivenKillPoints(ByVal index As Long, ByVal points As Single, ByVal Status As Byte)
    If index = 0 Then Exit Sub
    
    If TempPlayer(index).inParty > 0 Then
        PartyShareKillPoints index, TempPlayer(index).inParty, points
    Else
        Call SetPlayerKillPoints(index, GetPlayerKillPoints(index, Status) + points, Status)
        SendKillPoints index
        PlayerMsg index, GetTranslation("Has ganado:", , UnTrimBack) & points & GetTranslation("puntos de", , UnTrimBoth) & JusticeToStr(Status) & GetTranslation("ahora tienes:", , UnTrimBoth) & GetPlayerKillPoints(index, Status), GetColorByJustice(Status), , False
    End If
End Sub

Sub ComputeArmyPvP(ByVal attacker As Long, ByVal victim As Long)
    If attacker = 0 Or victim = 0 Then Exit Sub
    
    Dim AStatus As Byte, vStatus As Byte
    AStatus = GetPlayerPK(attacker)
    vStatus = GetPlayerPK(victim)
    
    If GetPlayerIP(attacker) = GetPlayerIP(victim) Then Exit Sub
    
    If AStatus = vStatus Then Exit Sub
    
    If GetPlayerPK(attacker) = NONE_PLAYER Or GetPlayerPK(victim) = NONE_PLAYER Then Exit Sub  'hero vs hero or normal vs normal
    
    
    
   
    
    'If GetPlayerPK(victim) = NONE_PLAYER Then
    
        'If Not IsNeutralEnabled(victim) Then Exit Sub
        
        
       ' Call SetNeutral(victim, False)
        'ComputeGivenKillPoints index, 1, AStatus
        
        'SetPlayerJusticePoints attacker, AStatus, GetPlayerJusticePoints(attacker, AStatus) + 1 'add point
    'Else
        Dim vpoints As Long
        vpoints = GetPlayerKillPoints(victim, vStatus)
        
        If vpoints < 1 Then Exit Sub 'out if player hadn't points, also we know that they are not doing spawnkill
        
        SetPlayerJusticePoints attacker, AStatus, GetPlayerJusticePoints(attacker, AStatus) + 1 'add point
          
        Dim factor As Single
        factor = 0
        
        factor = GetLostPointsPercentByRange(victim)
        factor = factor + GetPointPercentByDifferences(attacker, victim)
        
        If factor < 1 Then factor = 1
        If factor > 100 Then factor = 100
    '    Exit Sub
        Dim points As Long
        points = vpoints * (factor / 100)
        
        If points < 1 Then points = 1
        
        ComputeGivenKillPoints attacker, points, AStatus
        
        Call SetPlayerKillPoints(victim, vpoints - points, vStatus)
        SendKillPoints victim
        PlayerMsg victim, GetTranslation("Has perdido:", , UnTrimBack) & points & GetTranslation("puntos de", , UnTrimBoth) & JusticeToStr(vStatus) & GetTranslation("ahora tienes:", , UnTrimBoth) & GetPlayerKillPoints(victim, vStatus), GetColorByJustice(vStatus), , False
    'End If
End Sub

Function GetLostPointsPercentByRange(ByVal index As Long) As Single
    Dim MinFactor As Single, MaxFactor As Single
    MinFactor = 0
    MaxFactor = 5
    If GetPlayerPK(index) = NONE_PLAYER Then Exit Function
    
    GetLostPointsPercentByRange = (MaxFactor - MinFactor) / GetArmyMaxRange(GetPlayerPK(index)) * GetPlayerArmyRange(index) + MinFactor
    
End Function

Function GetPointPercentByDifferences(ByVal attacker As Long, ByVal victim As Long) As Single
    Dim MinFactor As Single, MaxFactor As Single
    MinFactor = 0
    MaxFactor = 5
    Dim dif As Long
    dif = GetLevelDifference(victim, attacker)
    If dif > 0 Then
        GetPointPercentByDifferences = (MaxFactor - MinFactor) / dif / (MAX_LEVELS \ 2) + MinFactor
    Else
        GetPointPercentByDifferences = MinFactor
    End If
    If GetPointPercentByDifferences < MinFactor Then GetPointPercentByDifferences = MinFactor
    If GetPointPercentByDifferences > MaxFactor Then GetPointPercentByDifferences = MaxFactor
End Function

Function GetArmyMaxRange(ByVal justice As Byte) As Byte
    Select Case justice
    Case HERO_PLAYER
        GetArmyMaxRange = HeroRangesTypeCount - 1
    Case PK_PLAYER
        GetArmyMaxRange = PkRangesTypeCount - 1
    Case Else
        GetArmyMaxRange = 0
    End Select
    
End Function

Function JusticeToStr(ByVal justice As Byte) As String
    If justice = HERO_PLAYER Then
        JusticeToStr = "Heroe"
    ElseIf justice = PK_PLAYER Then
        JusticeToStr = "Asesino"
    Else
        JusticeToStr = "Neutral"
    End If
End Function

Function GetColorByJustice(ByVal justice As Byte) As Byte
    If justice = HERO_PLAYER Then
        GetColorByJustice = Yellow
    ElseIf justice = PK_PLAYER Then
        GetColorByJustice = BrightRed
    Else
        GetColorByJustice = White
    End If
End Function



Public Function IsPlayerNeutral(ByVal index As Long) As Boolean
IsPlayerNeutral = True
If index > 0 And index <= Player_HighIndex Then
    If player(index).PK = YES Then
        IsPlayerNeutral = False
    End If
End If
End Function

Public Sub SetPlayerJustice(ByVal Killer As Long, ByVal Killed As Long)
If Not (Killer > 0 And Killer <= Player_HighIndex And Killed > 0 And Killed <= Player_HighIndex) Then Exit Sub

Dim SendJust As Boolean
SendJust = False

Select Case IsPlayerNeutral(Killer)
    Case True
        If IsPlayerNeutral(Killed) Then 'Player Killed Hero or Normal
            player(Killer).PK = PK_PLAYER
            Call GlobalMsg(GetPlayerName(Killer) & GetTranslation(" se ha convertido en un asesino!", , UnTrimFront), BrightRed, False, True)
            SendJust = True
            
            ResetPlayerArmy Killer, HERO_PLAYER
            SendKillPoints Killer
        Else 'Player Killed PK
            If GetPlayerPK(Killer) = HERO_PLAYER Then
                Call GlobalMsg(GetPlayerName(Killer) & GetTranslation(" ha hecho justicia!", , UnTrimFront), Yellow, False, True)
            Else
                player(Killer).PK = HERO_PLAYER
                Call GlobalMsg(GetPlayerName(Killer) & GetTranslation(" se ha convertido en un héroe!", , UnTrimFront), Yellow, False, True)
                SendJust = True
            End If
        End If
    Case False 'Killer Player is PK
        If IsPlayerNeutral(Killed) Then 'Add points in case of killed is neutral
            Call GlobalMsg(GetPlayerName(Killer) & GetTranslation(" ha cometido un crimen!", , UnTrimFront), BrightRed, False, True)
        End If
End Select



If SendJust Then
    SendJusticeToMap Killer
End If
End Sub

Public Sub SetPlayerHitJustice(ByVal Killer As Long, ByVal Killed As Long)
If Not (Killer > 0 And Killer <= Player_HighIndex And Killed > 0 And Killed <= Player_HighIndex) Then Exit Sub

If GetPlayerPK(Killer) = NONE_PLAYER And GetPlayerPK(Killed) = HERO_PLAYER And GetMapMoral(GetPlayerMap(Killer)) <> MAP_MORAL_ARENA Then
    Call SetPlayerPK(Killer, PK_PLAYER)
    PlayerMsg Killer, "Atacas a un héroe y te conviertes en asesino!", BrightRed
    SendJusticeToMap Killer
End If

End Sub

Sub SetPlayerJusticePoints(ByVal index As Long, ByVal justice As Byte, ByVal points As Long)
    Select Case justice
    Case HERO_PLAYER
        player(index).HeroPoints = points
    Case PK_PLAYER
        player(index).PKPoints = points
    End Select
       
End Sub

Function GetPlayerJusticePoints(ByVal index As Long, ByVal justice As Byte) As Long
    Select Case justice
    Case HERO_PLAYER
        GetPlayerJusticePoints = player(index).HeroPoints
    Case PK_PLAYER
        GetPlayerJusticePoints = player(index).PKPoints
    End Select
End Function



Public Sub PlayerRedemption(ByVal index As Long)
    'Special Punishments to the player
    Dim Status As Byte
    Status = GetPlayerPK(index)
    Call SetPlayerPK(index, NONE_PLAYER)
    ResetPlayerArmy index, Status
    SendKillPoints index
End Sub

Public Sub SetNeutral(ByVal index As Long, ByVal neutral As Boolean)
    player(index).NeutralEnabled = neutral
End Sub

Public Function IsNeutralEnabled(ByVal index As Long) As Boolean
    IsNeutralEnabled = player(index).NeutralEnabled
End Function

Function GetPlayerArmyRangeName(ByVal index As Long) As String
    If GetPlayerPK(index) = NONE_PLAYER Then
        GetPlayerArmyRangeName = vbNullString
    Else
        GetPlayerArmyRangeName = "<" & RangeToStr(GetPlayerArmyRange(index), GetPlayerPK(index)) & ">"
    End If
End Function

Function RangeToStr(ByVal range As Byte, ByVal army As Byte) As String
Select Case army
Case HERO_PLAYER
    Select Case range
    Case Soldado
        RangeToStr = GetTranslation("Soldado")
    Case Escolta
        RangeToStr = GetTranslation("Escolta")
    Case Teniente
        RangeToStr = GetTranslation("Teniente")
    Case Capitan
        RangeToStr = GetTranslation("Capitan")
    Case Protector
        RangeToStr = GetTranslation("Protector")
    Case Caballero
        RangeToStr = GetTranslation("Caballero")
    End Select
Case PK_PLAYER
    Select Case range
    Case Mercenario
        RangeToStr = GetTranslation("Mercenario")
    Case Aniquilador
        RangeToStr = GetTranslation("Aniquilador")
    Case Devastador
        RangeToStr = GetTranslation("Devastador")
    Case Asolador
        RangeToStr = GetTranslation("Asolador")
    Case Comandante
        RangeToStr = GetTranslation("Comandante")
    Case Elite
        RangeToStr = GetTranslation("Elite")
    End Select
Case NONE_PLAYER
    RangeToStr = "None"
End Select
End Function



Sub SendKillPoints(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SKillPoints
    
    buffer.WriteByte GetPlayerPK(index)
    buffer.WriteLong CLng(Round((GetPlayerKillPoints(index, GetPlayerPK(index)))))
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Function GetJusticeSpawnSite(ByVal justice As Byte, ByRef mapnum As Long, ByRef X As Long, ByRef Y As Long) As Boolean
    Dim header As String
    Select Case justice
    Case NONE_PLAYER
        header = "NEUTRAL"
    Case HERO_PLAYER
        header = "HERO"
    Case PK_PLAYER
        header = "PK"
    End Select
    
    Dim vmap As String, VX As String, VY As String
    
    vmap = GetVar(App.Path & "\data\army.ini", header, "SpawnMap")
    VX = GetVar(App.Path & "\data\army.ini", header, "SpawnX")
    VY = GetVar(App.Path & "\data\army.ini", header, "SpawnY")
    
    If IsNumeric(vmap) And IsNumeric(VX) And IsNumeric(VY) Then
        mapnum = vmap
        X = VX
        Y = VY
        If Not OutOfBoundries(X, Y, mapnum) Then
            GetJusticeSpawnSite = True
        End If
    End If
End Function


Function CanPlayerAttackByJustice(ByVal attacker As Long, ByVal victim As Long, Optional ByVal sendmsg As Boolean = True) As Boolean
    CanPlayerAttackByJustice = True
    If (GetPlayerPK(attacker) = PK_PLAYER And GetPlayerPK(victim) = NONE_PLAYER) Or (GetPlayerPK(attacker) = NONE_PLAYER And GetPlayerPK(victim) = PK_PLAYER) Then
        If Abs(GetLevelDifference(attacker, victim)) > 20 Then
             CanPlayerAttackByJustice = False
             If sendmsg Then
                 PlayerMsg attacker, "Neutrales y Asesinos sólo pueden pelearse cuando se llevan menos de 20 lvl's de diferencia", BrightRed
             End If
        End If
    End If
End Function

