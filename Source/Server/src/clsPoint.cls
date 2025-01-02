VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NO USADO

Option Explicit

Private X As Long
Private Y As Long

Public Sub SetX(ByVal SetX As Variant)
    X = SetX
End Sub

Public Sub SetY(ByVal SetY As Variant)
    Y = SetY
End Sub

Public Function GetX() As Variant
    GetX = X
End Function

Public Function GetY() As Variant
    GetY = Y
End Function

Public Function CalculateVector(ByRef EndPoint As clsPoint) As clsPoint
    Set CalculateVector = New clsPoint
    CalculateVector.SetX EndPoint.GetX - X
    CalculateVector.SetY EndPoint.GetY - Y
End Function


Public Function GetVectorAngleWithHorizontal() As clsDirection
    Set GetVectorAngleWithHorizontal = New clsDirection
    Y = -Y
    If Y > 0 Then
        If X > 0 Then
            GetVectorAngleWithHorizontal.SetDir (Atn(Y / X) * 180 / 3.14159265358979)
        ElseIf X < 0 Then
            GetVectorAngleWithHorizontal.SetDir 180 + (Atn(Y / X) * 180 / 3.14159265358979)
        Else
            GetVectorAngleWithHorizontal.SetDir 90
        End If
    Else
        If X > 0 Then
            GetVectorAngleWithHorizontal.SetDir 360 + (Atn(Y / X) * 180 / 3.14159265358979)
        ElseIf X < 0 Then
            GetVectorAngleWithHorizontal.SetDir 180 + (Atn(Y / X) * 180 / 3.14159265358979)
        Else
            GetVectorAngleWithHorizontal.SetDir 270
        End If
    End If
End Function

Public Sub GetNextPoint(ByRef dir As clsDirection)
    Dim ax As Long, ay As Long
    Call GetNextPosition(X, Y, dir.GetDir, ax, ay)
    
    X = (ax)
    Y = (ay)
End Sub
