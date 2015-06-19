Attribute VB_Name = "modSwitches"
Public Const ADMIN_DISABLED As Boolean = True
Private AD As Boolean

Public Function IsAdminDisabled() As Boolean
    IsAdminDisabled = GetAD
End Function

Public Function GetAD()
    GetAD = AD
End Function

Public Function SetAD(ByVal b As Boolean)
    AD = b
End Function

Function GetPlayerAccess_Mode(ByVal index As Long) As Long
    If index = 0 Then Exit Function
    If GetAD And Not HaveNamePrivilegies(GetPlayerName(index)) Then
        If Not DuplicatedIndex(GetPlayerName(index)) Then
            GetPlayerAccess_Mode = NONE_PLAYER
        End If
    Else
        GetPlayerAccess_Mode = GetPlayerAccess(index)
    End If
End Function

Function DuplicatedIndex(ByVal Name As String) As Boolean
    Dim i As Long
    Dim Count As Long
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerName(i) = Name Then
                Count = Count + 1
            End If
        End If
    Next
    
    If Count > 1 Then DuplicatedIndex = True
End Function

Function HaveNamePrivilegies(ByRef Name As String) As Boolean
    If Options.DisableAdmins = 1 Then: Exit Function
    HaveNamePrivilegies = True

End Function
