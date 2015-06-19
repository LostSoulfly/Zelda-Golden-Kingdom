Attribute VB_Name = "modRides"



Public MAX_GROWTH_BY_TIME As Single
Public MIN_GROWTH_BY_TIME As Single
Public MIN_BURN_BY_TIME As Single
Public MAX_BURN_BY_TIME As Single
Public MAX_STAMINA As Single
Public SPECIAL_BONUS_MS As Long


Public Const CHECK_GROWTH_TIME As Single = 0.05
Public GrowthTimer As Long





Public Type PlayerRideRec
    StaminaBurn As Single
    Stamina As Single
    Growth As Single
    LastStamina As Single
    Switch As Boolean
    SpecialBonusTimer As Long
    SpecialBonusActive As Boolean
    RandomBonusVect() As Boolean
End Type

Sub InitRideInfo(ByVal index As Long)
    InitRandomBonus index
    SetRideStamina MyIndex, MAX_STAMINA
    Player(MyIndex).RideInfo.StaminaBurn = MIN_BURN_BY_TIME
End Sub

Sub InitRandomBonus(ByVal index As Long)
    ReDim Player(index).RideInfo.RandomBonusVect(0 To MAX_STAMINA \ 2)
    SelectRandomly index
End Sub

Sub SelectRandomly(ByVal index As Long)
    Dim i As Long
    With Player(index).RideInfo
        i = Rand(UBound(.RandomBonusVect), LBound(.RandomBonusVect))
        .RandomBonusVect(i) = True
    End With
End Sub

Sub CheckSpecialBonusAppear(ByVal index As Long)
    If Player(index).RideInfo.SpecialBonusTimer = 0 Then
        If CanCalculateGrowth(index) Then
            With Player(index).RideInfo
            
            If GetRideStamina(index) <= MAX_STAMINA \ 2 And GetRideStamina(index) > 0 Then
                If .RandomBonusVect(CLng(GetRideStamina(index))) Then
                        .RandomBonusVect(GetRideStamina(index)) = False
                        StartSpecialBonus index
                End If
            End If
            
            
            End With
        End If
    End If
End Sub

Sub SetSpecialBonus(ByVal index As Long, ByVal value As Boolean)
    Player(index).RideInfo.SpecialBonusActive = value
End Sub

Function IsRecoveryingSpecialBonus(ByVal index As Long) As Boolean
    IsRecoveryingSpecialBonus = Player(index).RideInfo.SpecialBonusActive
End Function
Sub EndSpecialBonus(ByVal index As Long)
    Player(index).RideInfo.SpecialBonusTimer = 0
End Sub

Function IsSpecialBonusActive(ByVal index As Long) As Boolean

    If GetSpecialBonusTimer(index) <> 0 Then
        If GetSpecialBonusTimer(index) > GetTickCount Then
            IsSpecialBonusActive = True
        Else
            Player(index).RideInfo.SpecialBonusTimer = 0
        End If
    End If
End Function

Function GetSpecialBonusTimer(ByVal index As Long) As Long
    GetSpecialBonusTimer = Player(index).RideInfo.SpecialBonusTimer
End Function

Sub StartSpecialBonus(ByVal index As Long)
    Player(index).RideInfo.SpecialBonusTimer = GetTickCount + SPECIAL_BONUS_MS
End Sub

Sub CheckSpecialBonus(ByVal index As Long, ByVal tick As Long)
    If IsSpecialBonusActive(index) Then
        Call SetRideGrowth(index, MAX_GROWTH_BY_TIME * 2)
        Player(index).RideInfo.SpecialBonusTimer = 0
        SetSpecialBonus index, True
    End If
End Sub

Function CanCalculateGrowth(ByVal index As Long) As Boolean
    CanCalculateGrowth = (Not Player(index).RideInfo.Switch Or GetRideStamina(index) < GetRideLastStamina(index)) And Not IsRecoveryingSpecialBonus(index)
End Function

Sub SetRideLastStamina(ByVal index As Long, ByVal SetLastStamina As Single)
    Player(index).RideInfo.LastStamina = SetLastStamina
End Sub
 
Function GetRideLastStamina(ByVal index As Long) As Single
    GetRideLastStamina = Player(index).RideInfo.LastStamina
End Function

Sub SetRideGrowth(ByVal index As Long, ByVal SetGrowth As Single)
    Player(index).RideInfo.Growth = SetGrowth
End Sub
 
Function GetRideGrowth(ByVal index As Long) As Single
    GetRideGrowth = Player(index).RideInfo.Growth
End Function

Sub SetRideStamina(ByVal index As Long, ByVal SetStamina As Single)
    Player(index).RideInfo.Stamina = SetStamina
End Sub

Sub CheckStaminaGrowth(ByVal index As Long)
    If CanCalculateGrowth(index) Then
        'If IsPlayerRiding(index) And Not IsPlayerRunning(index) Then
        If Not IsPlayerRunning(index) Then
            If GetRideStamina(index) < MAX_STAMINA Then
                CalculateStaminaGrowth index
                SetRideLastStamina index, GetRideStamina(index)
                CalculateStaminaBurn index
                CheckSpecialBonus index, GetTickCount
            End If
        End If
    End If
End Sub

Sub CalculateStaminaBurn(ByVal index As Long)
    Dim X As Single
    X = GetRideLastStamina(index)
    Dim y As Single
    If X = 0 Then
        y = MAX_BURN_BY_TIME
    Else
        y = Line(MAX_STAMINA, 0, MIN_BURN_BY_TIME, MAX_BURN_BY_TIME, MAX_BURN_BY_TIME, X)
    End If
    
    Player(index).RideInfo.StaminaBurn = y
End Sub

Public Function Line(ByVal MaxX As Variant, ByVal MinX As Variant, ByVal MaxY As Variant, ByVal MinY As Variant, ByVal n As Variant, ByVal X As Variant, Optional ByVal AjustLimits As Boolean = True) As Variant
    Dim VX As Variant
    Dim VY As Variant
    VX = MaxX - MinX
    VY = MaxY - MinY
    
    If VX = 0 Then Exit Function
    
    Dim m As Variant
    m = VY / VX
    Line = m * X + n
    
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

Sub CalculateStaminaGrowth(ByVal index As Long)
    Dim X As Single, y As Single
    X = CSng(GetRideStamina(index)) / MAX_STAMINA
    y = X * MAX_GROWTH_BY_TIME
    If y <= MIN_GROWTH_BY_TIME Then
        y = MIN_GROWTH_BY_TIME
    End If
    Call SetRideGrowth(index, y)
    Player(index).RideInfo.Switch = True
End Sub

Function GetStaminaGrowth(ByVal Stamina As Single) As Single
    Dim X As Single, y As Single
    X = Stamina / MAX_STAMINA
    y = X * MAX_GROWTH_BY_TIME
    If y <= MIN_GROWTH_BY_TIME Then
        y = MIN_GROWTH_BY_TIME
    End If
    GetStaminaGrowth = y
End Function

Sub CheckIncreaseRideStamina(ByVal index As Long)
    'If IsPlayerRiding(index) And Not CanRideRun(index) Then
    If Not CanRideRun(index) Then
        If GetRideStamina(index) < MAX_STAMINA Then
            Dim X As Single
            X = GetRideStamina(index) + GetRideGrowth(index)
            If X >= MAX_STAMINA Then
                X = MAX_STAMINA
                SetSpecialBonus index, False
                Player(index).RideInfo.Switch = False
                Player(index).RideInfo.StaminaBurn = MIN_BURN_BY_TIME
                InitRandomBonus index
            End If
            Call SetRideStamina(index, X)
        End If
    End If
End Sub

Function CanRideRun(ByVal index As Long)
    CanRideRun = IsPlayerRunning(index) And Not IsRecoveryingSpecialBonus(index)
End Function

Sub DecreaseRideStamina(ByVal index As Long)
    If Player(index).RideInfo.Stamina > 0 Then
        Player(index).RideInfo.Stamina = Player(index).RideInfo.Stamina - Player(index).RideInfo.StaminaBurn
        CheckSpecialBonusAppear index
    End If
End Sub

Function GetRideStamina(ByVal index As Long) As Single
    GetRideStamina = Player(index).RideInfo.Stamina
End Function


Function IsPlayerRiding(ByVal index As Long) As Long
    If Player(index).State = StateRiding Then
        IsPlayerRiding = True
    End If
End Function

Sub SetPlayerState(ByVal index As Long, ByVal State As PlayerStateType)
    Player(index).State = State
End Sub

Function IsPlayerRolling(ByVal index As Long) As Boolean
    IsPlayerRolling = Player(index).MovementSprite
End Function



Public Sub HandleStaminaInfo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    MAX_STAMINA = Buffer.ReadLong
    MIN_BURN_BY_TIME = Buffer.ReadLong
    MAX_BURN_BY_TIME = Buffer.ReadLong
    MIN_GROWTH_BY_TIME = Buffer.ReadLong
    MAX_GROWTH_BY_TIME = Buffer.ReadLong
    SPECIAL_BONUS_MS = Buffer.ReadLong
    
    
    InitRideInfo MyIndex
    
    Set Buffer = Nothing
   
End Sub



