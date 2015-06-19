VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NO USADO
Private StartPoint As clsPoint
Private EndPoint As clsPoint

Private DirVector As clsPoint

Public Sub SetVector(ByRef p1 As clsPoint, ByRef p2 As clsPoint)
    Set StartPoint = p1
    Set EndPoint = p2
    
    Set DirVector = New clsPoint
    Set DirVector = p1.CalculateVector(p2)
End Sub

Public Function GetVector() As clsPoint
    Set GetVector = DirVector
End Function

Public Function GetVectorAngle() As clsDirection
    Set GetVectorAngle = DirVector.GetVectorAngleWithHorizontal
End Function
