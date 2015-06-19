Attribute VB_Name = "modSail"

Private Const SAILING_CUSTOM_SPRITE As Byte = 7


Sub ClearSailing(ByVal index As Long)
    SetPlayerCustomSprite index, 0
    SendPlayerSprite index
    SendPlayerStateToMap index, GetPlayerState(index)
    
    SetPlayerSpeeds index, Walk_Speed, Run_Speed
    SendPlayerSpeedsToMap index
    
    SetPlayerState index, StateNone
End Sub

Public Function GetPlayerSailSprite(ByVal index As Long) As Long
    GetPlayerSailSprite = SAILING_CUSTOM_SPRITE
End Function


Sub StartSailing(ByVal index As Long)
    SetPlayerCustomSprite index, GetPlayerSailSprite(index)
    SendPlayerSpriteToMap index
    SendPlayerStateToMap index, GetPlayerState(index)
    SetPlayerState index, StateSailing
End Sub

Public Sub PlayerNavigation(ByVal index As Long)

Select Case GetPlayerState(index)
Case StateNone
    If Not CanNavigate(index) Then
        ForcePlayerMove index, MOVING_WALKING, GetOppositeDir(GetPlayerDir(index))
    Else
        StartSailing index
        ForcePlayerMove index, MOVING_WALKING, GetPlayerDir(index)
    End If
Case StateSailing
    ClearSailing index
    ForcePlayerMove index, MOVING_WALKING, GetPlayerDir(index)
Case Else
    ForcePlayerMove index, MOVING_WALKING, GetOppositeDir(GetPlayerDir(index))
End Select

    
End Sub

Function CanNavigate(ByVal index As Long) As Boolean
    If IsStateAllowedOnMap(GetPlayerMap(index), StateSailing) Then
        CanNavigate = True
    End If
End Function
