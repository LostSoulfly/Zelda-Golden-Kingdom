Attribute VB_Name = "modPlayerState"



Public Enum PlayerStateType
    StateNone = 0
    StateSailing = 1
    StateRiding = 2
    Max_States
End Enum


Public Sub SetPlayerState(ByVal index As Long, ByVal state As PlayerStateType)
    player(index).state = state
End Sub

Public Function GetPlayerState(ByVal index As Long) As PlayerStateType
    GetPlayerState = player(index).state
End Function

Sub CheckPlayerStateAtWarp(ByVal index As Long, ByVal mapnum As Long)
    Select Case GetPlayerState(index)
    Case StateSailing
    Case StateRiding
        If Not IsStateAllowedOnMap(mapnum, StateRiding) Then
            ResetPlayerState index
        End If
    End Select
End Sub

Sub CheckPlayerStateAtJoin(ByVal index As Long)
    Select Case GetPlayerState(index)
    Case StateSailing
    Case StateRiding
        ResetPlayerState index
    End Select
End Sub


Sub ResetPlayerState(ByVal index As Long)
    Select Case GetPlayerState(index)
    Case StateSailing
        ClearSailing index
    Case StateRiding
        ClearRiding index
    End Select
End Sub



Public Function CanPlayerChangeState(ByVal index As Long, ByVal state As PlayerStateType) As Boolean
    Select Case state
    Case StateRiding
        If CanRide(index) Then
            CanPlayerChangeState = True
        End If
    Case StateSailing
        CanPlayerChangeState = True
    
    End Select
End Function

Public Sub ChangePlayerState(ByVal index As Long, ByVal state As PlayerStateType)
    Select Case state
    Case StateRiding
        PlayerRide index
    Case StateSailing
        PlayerNavigation index
    End Select
End Sub



Sub CheckPlayerStateChange(ByVal index As Long, ByVal state As PlayerStateType)
    Select Case GetPlayerState(index)
    Case StateNone
        If CanPlayerChangeState(index, state) Then
            ChangePlayerState index, state
        End If
    Case StateRiding
        If state = StateRiding Then 'return normal
            ClearRiding index
        End If
    Case StateSailing
        
    
    End Select
End Sub



Function IsStateAllowedOnMap(ByVal mapnum As Long, ByVal state As PlayerStateType) As Boolean
    IsStateAllowedOnMap = GetMapState(mapnum, state)
End Function

Sub SendPlayerStateToMap(ByVal index As Long, ByVal state As PlayerStateType)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerState
    buffer.WriteLong index
    buffer.WriteByte state
    
    SendDataToMap GetPlayerMap(index), buffer.ToArray
End Sub



