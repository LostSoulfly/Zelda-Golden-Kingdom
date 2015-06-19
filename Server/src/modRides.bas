Attribute VB_Name = "modRides"

Public Const RIDE_WALK_SPEED As Long = 14
Public Const RIDE_RUN_SPEED As Long = 22
Private Const RIDING_CUSTOM_SPRITE As Byte = 2

Private Const HORSE_1_CUSTOM_SPRITE As Byte = 1
Private Const HORSE_1_SPELL As Long = 200
Private Const HORSE_2_CUSTOM_SPRITE As Byte = 2
Private Const HORSE_2_SPELL As Long = 201
Private Const HORSE_3_CUSTOM_SPRITE As Byte = 3
Private Const HORSE_3_SPELL As Long = 202
Private Const HORSE_4_CUSTOM_SPRITE As Byte = 4
Private Const HORSE_4_SPELL As Long = 203
Private Const HORSE_5_CUSTOM_SPRITE As Byte = 5
Private Const HORSE_5_SPELL As Long = 204
Private Const HORSE_6_CUSTOM_SPRITE As Byte = 6
Private Const HORSE_6_SPELL As Long = 205

Private Const HORSE_7_CUSTOM_SPRITE As Byte = 8
Private Const HORSE_7_SPELL As Long = 220
Private Const HORSE_8_CUSTOM_SPRITE As Byte = 9
Private Const HORSE_8_SPELL As Long = 221
Private Const HORSE_9_CUSTOM_SPRITE As Byte = 10
Private Const HORSE_9_SPELL As Long = 222
Private Const HORSE_10_CUSTOM_SPRITE As Byte = 11
Private Const HORSE_10_SPELL As Long = 223
Private Const HORSE_11_CUSTOM_SPRITE As Byte = 12
Private Const HORSE_11_SPELL As Long = 224
Private Const HORSE_12_CUSTOM_SPRITE As Byte = 13
Private Const HORSE_12_SPELL As Long = 225

Private Const ZORA_1_SPELL As Long = 230
Private Const ZORA_2_SPELL As Long = 231

Public Const MAX_GROWTH_BY_TIME As Single = 6
Public Const MIN_GROWTH_BY_TIME As Single = 2
Public Const MIN_BURN_BY_TIME As Single = 1
Public Const MAX_BURN_BY_TIME As Single = 4
Public Const MAX_STAMINA As Single = 100
Public Const SPECIAL_BONUS_MS As Long = 400
Public Const MAX_ROLL_STAMINA As Single = 300

Public Const ROLL_SPEED As Long = 26

Private Const RollClass1 As Long = 5
Private Const RollClass2 As Long = 6

Public Function GetPlayerRideSpeed(ByVal index As Long, ByVal Movement As Byte) As Long
    If Movement = 1 Then
        GetPlayerRideSpeed = RIDE_WALK_SPEED
    ElseIf Movement = 2 Then
        GetPlayerRideSpeed = RIDE_RUN_SPEED
    End If
End Function

Public Function GetPlayerRideSprite(ByVal index As Long) As Long
    Dim spellnum As Long
    Dim X As Long
    spellnum = GetLastUsedSpell(index)
    
    If spellnum = HORSE_1_SPELL Then
        X = HORSE_1_CUSTOM_SPRITE
    ElseIf spellnum = HORSE_2_SPELL Then
        X = HORSE_2_CUSTOM_SPRITE
    ElseIf spellnum = HORSE_3_SPELL Then
        X = HORSE_3_CUSTOM_SPRITE
    ElseIf spellnum = HORSE_4_SPELL Then
        X = HORSE_4_CUSTOM_SPRITE
    ElseIf spellnum = HORSE_5_SPELL Then
        X = HORSE_5_CUSTOM_SPRITE
    ElseIf spellnum = HORSE_6_SPELL Then
        X = HORSE_6_CUSTOM_SPRITE
    ElseIf spellnum = HORSE_7_SPELL Then
        X = HORSE_7_CUSTOM_SPRITE
    ElseIf spellnum = HORSE_8_SPELL Then
        X = HORSE_8_CUSTOM_SPRITE
    ElseIf spellnum = HORSE_9_SPELL Then
        X = HORSE_9_CUSTOM_SPRITE
    ElseIf spellnum = HORSE_10_SPELL Then
        X = HORSE_10_CUSTOM_SPRITE
    ElseIf spellnum = HORSE_11_SPELL Then
        X = HORSE_11_CUSTOM_SPRITE
    ElseIf spellnum = HORSE_12_SPELL Then
        X = HORSE_12_CUSTOM_SPRITE
    Else
        X = HORSE_1_CUSTOM_SPRITE
    End If
    GetPlayerRideSprite = X
End Function

Sub ClearRiding(ByVal index As Long)
    SetPlayerCustomSprite index, 0
    SendPlayerSprite index
    SetPlayerSpeeds index, Walk_Speed, Run_Speed
    SendPlayerSpeedsToMap index
    SetPlayerState index, StateNone
    SendPlayerStateToMap index, GetPlayerState(index)
End Sub

Sub StartRiding(ByVal index As Long)
    SetPlayerCustomSprite index, GetPlayerRideSprite(index)
    SendPlayerSpriteToMap index
    SetPlayerSpeeds index, GetPlayerRideSpeed(index, MOVING_WALKING), GetPlayerRideSpeed(index, MOVING_RUNNING)
    SendPlayerSpeedsToMap index
    SetPlayerState index, StateRiding
    SendPlayerStateToMap index, GetPlayerState(index)
End Sub

Public Sub PlayerRide(ByVal index As Long)
    'If CanRide(index) Then
        StartRiding index
    'End If
End Sub

Function CanClassRide(ByVal ClassNum As Long) As Boolean
    If ClassNum >= 1 And ClassNum <= 4 Or ClassNum >= 9 And ClassNum <= 10 Then
        CanClassRide = True
    End If
End Function

Function CanClassRoll(ByVal ClassNum As Long) As Boolean
    If ClassNum = RollClass1 Or ClassNum = RollClass2 Then
        CanClassRoll = True
    End If
End Function


Function CanRide(ByVal index As Long) As Boolean
    If IsStateAllowedOnMap(GetPlayerMap(index), StateRiding) Then
        If CanClassRide(GetPlayerClass(index)) Then
            CanRide = True
        End If
    End If
End Function

Sub SendStaminaInfo(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SStaminaInfo
    If CanClassRoll(GetPlayerClass(index)) Then
        buffer.WriteLong MAX_ROLL_STAMINA
    Else
        buffer.WriteLong MAX_STAMINA
    End If
    buffer.WriteLong MIN_BURN_BY_TIME
    buffer.WriteLong MAX_BURN_BY_TIME
    buffer.WriteLong MIN_GROWTH_BY_TIME
    buffer.WriteLong MAX_GROWTH_BY_TIME
    buffer.WriteLong SPECIAL_BONUS_MS
    
    SendDataTo index, buffer.ToArray
    
    Set buffer = Nothing
End Sub
