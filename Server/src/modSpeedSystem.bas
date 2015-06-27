Attribute VB_Name = "modSpeedSystem"
Option Explicit

Public Const Walk_Speed As Byte = 6
Public Const Run_Speed As Byte = 8

Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2
Public Const SPEEDHACK_LAPSE As Long = 10 'seconds

Public SpeedHack_Timer As Long

Sub CheckSpeedHack(ByVal index As Long)
    If GetPlayerAccess_Mode(index) = NONE_PLAYER Then
        SendSpeedReqTo (index)
    End If
End Sub

Sub SendSpeedReqTo(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpeedReq
    Buffer.WriteByte SPEEDHACK_LAPSE
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleSpeedAck(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    Set Buffer = Nothing
    
    Call CheckPlayerSpeedHack(index, GetTickCount)
    
End Sub

Sub CheckPlayerSpeedHack(ByVal index As Long, ByVal Tick As Long)
    If GetPlayerSpeedHackChecker(index) = 0 Then 'First Send
        SetPlayerSpeedHackChecker index, GetTickCount
    Else
        Dim T As Long
        T = GetPlayerSpeedHackChecker(index)
        SetPlayerSpeedHackChecker index, Tick
        If (Tick - T) / 1000 < CSng(SPEEDHACK_LAPSE - 2) Then
            KickPlayer index, "SpeedHack"
        End If
        GlobalMsg (Tick - T) / 1000, BrightRed, False
    End If
    
End Sub

Sub ComputePlayerSpeed(ByVal index As Long)
    SetPlayerSpeed index, MOVING_WALKING, Walk_Speed
    
    If CanClassRoll(GetPlayerClass(index)) Then
        SetPlayerSpeed index, MOVING_RUNNING, ROLL_SPEED
    Else
        SetPlayerSpeed index, MOVING_RUNNING, Run_Speed
    End If
End Sub

Sub SetPlayerSpeedHackChecker(ByVal index As Long, ByVal val As Long)
    TempPlayer(index).SpeedHackChecker = val
End Sub
Function GetPlayerSpeedHackChecker(ByVal index As Long) As Long
    GetPlayerSpeedHackChecker = TempPlayer(index).SpeedHackChecker
End Function

Sub SendPlayerSpeedToMap(ByVal index As Long, ByVal Movement As Long, ByVal Speed As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerSpeed
    Buffer.WriteLong index
    Buffer.WriteLong Movement
    Buffer.WriteLong Speed
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpeedTo(ByVal index As Long, ByVal SendIndex As Long, ByVal Movement As Long, ByVal Speed As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerSpeed
    Buffer.WriteLong SendIndex
    Buffer.WriteLong Movement
    Buffer.WriteLong Speed
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpeedToAll(ByVal index As Long, ByVal Movement As Long, ByVal Speed As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerSpeed
    Buffer.WriteLong index
    Buffer.WriteLong Movement
    Buffer.WriteLong Speed
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub


Sub SetPlayerSpeed(ByVal index As Long, ByVal Movement As Long, ByVal Speed As Long)
    If Movement = MOVING_WALKING Then
        TempPlayer(index).WalkSpeed = Speed
    ElseIf Movement = MOVING_RUNNING Then
        TempPlayer(index).RunSpeed = Speed
    End If
End Sub

Sub SetPlayerSpeeds(ByVal index As Long, ByVal Walk_Speed As Long, ByVal Run_Speed As Long)
    SetPlayerSpeed index, MOVING_WALKING, Walk_Speed
    SetPlayerSpeed index, MOVING_RUNNING, Run_Speed
End Sub

Sub SendPlayerSpeedsToMap(ByVal index As Long)
    SendPlayerSpeedToMap index, MOVING_WALKING, GetPlayerSpeed(index, MOVING_WALKING)
    SendPlayerSpeedToMap index, MOVING_RUNNING, GetPlayerSpeed(index, MOVING_RUNNING)
End Sub

Sub SendPlayerSpeeds(ByVal index As Long, ByVal SendIndex As Long)
    SendPlayerSpeedTo index, SendIndex, MOVING_WALKING, GetPlayerSpeed(index, MOVING_WALKING)
    SendPlayerSpeedTo index, SendIndex, MOVING_RUNNING, GetPlayerSpeed(index, MOVING_RUNNING)
End Sub

Function GetPlayerSpeed(ByVal index As Long, ByVal Movement As Long) As Long
    If Movement = MOVING_WALKING Then
        GetPlayerSpeed = TempPlayer(index).WalkSpeed
    ElseIf Movement = MOVING_RUNNING Then
        GetPlayerSpeed = TempPlayer(index).RunSpeed
    End If
End Function


Sub HandleFSpellActivacion(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    Set Buffer = Nothing
    
    Call Impactar(index, index, 1, 1)
    
End Sub
