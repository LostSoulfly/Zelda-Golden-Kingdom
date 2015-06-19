VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NO USADO
Private dir As Single

Public Sub SetDir(ByVal d As Single)
    dir = d
End Sub

Public Function GetDir() As Single
    GetDir = dir
End Function



Private Static Function GetDirByAngle(ByVal Angle As Single) As Byte

    If Angle > 315 Or Angle <= 45 Then
        GetDirByAngle = DIR_RIGHT
    ElseIf Angle > 45 And Angle <= 135 Then
        GetDirByAngle = DIR_UP
    ElseIf Angle > 135 And Angle <= 225 Then
        GetDirByAngle = DIR_LEFT
    ElseIf Angle > 225 And Angle <= 315 Then
        GetDirByAngle = DIR_DOWN
    End If
End Function

Public Function GetPriorityStackDir(Optional ByRef LastUsedDir As clsDirection = Nothing) As clsStack
    Set GetPriorityStackDir = New clsStack
    
    dir = GetDirByAngle(dir)
    If Not (LastUsedDir Is Nothing) Then
        Dim NotUsedDir As clsDirection
        Set NotUsedDir = LastUsedDir.GetOppositeDir
        If Not Me.GetOppositeDir.CompareDir(NotUsedDir) Then: GetPriorityStackDir.Push Me.GetOppositeDir
        If Not Me.GetClockwiseNextDir.CompareDir(NotUsedDir) Then: GetPriorityStackDir.Push Me.GetClockwiseNextDir
        If Not Me.GetClockwisePrevDir.CompareDir(NotUsedDir) Then: GetPriorityStackDir.Push Me.GetClockwisePrevDir
        If Not Me.CompareDir(NotUsedDir) Then: GetPriorityStackDir.Push Me
    Else
        GetPriorityStackDir.Push Me.GetOppositeDir
        GetPriorityStackDir.Push Me.GetClockwiseNextDir
        GetPriorityStackDir.Push Me.GetClockwisePrevDir
        GetPriorityStackDir.Push Me
    End If
    
    
End Function

Public Function CompareDir(ByRef d As clsDirection) As Boolean
    If dir = d.GetDir Then
        CompareDir = True
    End If
End Function


Public Function GetOppositeDir() As clsDirection
    Set GetOppositeDir = New clsDirection
    Select Case dir
    Case DIR_UP
        GetOppositeDir.SetDir DIR_DOWN
    Case DIR_DOWN
        GetOppositeDir.SetDir DIR_UP
    Case DIR_LEFT
        GetOppositeDir.SetDir DIR_RIGHT
    Case DIR_RIGHT
        GetOppositeDir.SetDir DIR_LEFT
    End Select
End Function

Public Function GetClockwiseNextDir() As clsDirection
    Set GetClockwiseNextDir = New clsDirection
    Select Case dir
    Case DIR_UP
        GetClockwiseNextDir.SetDir DIR_RIGHT
    Case DIR_DOWN
        GetClockwiseNextDir.SetDir DIR_LEFT
    Case DIR_LEFT
        GetClockwiseNextDir.SetDir DIR_UP
    Case DIR_RIGHT
        GetClockwiseNextDir.SetDir DIR_DOWN
    End Select
End Function

Public Function GetClockwisePrevDir() As clsDirection
    Set GetClockwisePrevDir = New clsDirection
    Select Case dir
    Case DIR_UP
        GetClockwisePrevDir.SetDir DIR_LEFT
    Case DIR_DOWN
        GetClockwisePrevDir.SetDir DIR_RIGHT
    Case DIR_LEFT
        GetClockwisePrevDir.SetDir DIR_DOWN
    Case DIR_RIGHT
        GetClockwisePrevDir.SetDir DIR_UP
    End Select
End Function

