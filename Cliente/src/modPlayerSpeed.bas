Attribute VB_Name = "modPlayerSpeed"

Public SpeedHack_Lapse As Long
Public SpeedHack_Timer As Long

Public Sub HandleSpeedReq(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SendSpeedAck GetTickCount
    SpeedHack_Lapse = Buffer.ReadByte
    SpeedHack_Timer = GetTickCount + SpeedHack_Lapse * 1000
    
    Set Buffer = Nothing
End Sub

Public Sub SendSpeedAck(ByVal tick As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CSpeedAck
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub


Public Sub HandlePlayerSpeed(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Call SetPlayerSpeed(Buffer.ReadLong, Buffer.ReadLong, Buffer.ReadLong)
    
    Set Buffer = Nothing
End Sub


Sub SetPlayerSpeed(ByVal index As Long, ByVal Movement As Long, ByVal Speed As Long)
    If Movement = MOVING_WALKING Then
        Player(index).WalkSpeed = Speed
    ElseIf Movement = MOVING_RUNNING Then
        Player(index).RunSpeed = Speed
    End If
End Sub

Function GetPlayerSpeed(ByVal index As Long, ByVal Movement As Long) As Long
    If Movement = MOVING_WALKING Then
        GetPlayerSpeed = Player(index).WalkSpeed
    ElseIf Movement = MOVING_RUNNING Then
        GetPlayerSpeed = Player(index).RunSpeed
    End If
End Function




Sub GetPlayerRunSpeed(ByVal index As Long, ByRef Movement As Byte, ByRef Speed As Long)
    If IsPlayerRiding(index) Then
        If GetRideStamina(index) > 0 And (CanRideRun(index)) Then
            DecreaseRideStamina index
            Speed = RUN_SPEED
            Movement = MOVING_RUNNING
        Else
            Speed = WALK_SPEED
            Movement = MOVING_WALKING
        End If
    ElseIf IsPlayerRolling(index) Then
        If GetRideStamina(index) > 0 And (CanRideRun(index)) Then
            DecreaseRideStamina index
            Speed = RUN_SPEED
            Movement = MOVING_RUNNING
        Else
            Speed = WALK_SPEED
            Movement = MOVING_WALKING
        End If
    
    Else
        Speed = RUN_SPEED
        Movement = MOVING_RUNNING
    End If
End Sub


Function IsPlayerRunning(ByVal index As Long) As Boolean
    If Player(index).Moving = MOVING_RUNNING Or ((DirUp Or DirDown Or DirLeft Or DirRight) And ShiftDown) Then
        IsPlayerRunning = True
    End If
End Function

