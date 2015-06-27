VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m As Long
Private p As clsPoint

Public Sub Class_initialize()
    Set p = New clsPoint
    m = 0
End Sub

Public Sub SetPoint(ByRef Point As clsPoint)
    Set p = Point
End Sub

Public Sub SetMap(ByVal map As Long)
    m = map
End Sub

Public Function GetMap() As Long
    GetMap = m
End Function

Public Function GetPoint() As clsPoint
    Set GetPoint = p
End Function
'NO USADO
Public Function GetNextPosition(ByRef dir As clsDirection, ByRef NextPos As clsPosition) As Boolean
    Dim X As Long, Y As Long
    X = p.GetX
    Y = p.GetY
    GetNextPosition = Not GetNextPositionByRef(dir.GetDir, m, X, Y)
    
    Set NextPos = New clsPosition
    Dim po As clsPoint
    Set po = New clsPoint
    po.SetX X
    po.SetY Y
    
    NextPos.SetPoint po
    NextPos.SetMap m
End Function

Public Function IsBlocked() As Boolean
    'If Not NpcTileIsOpen(m, p.GetX, p.GetY) Then
        'sBlocked = True
    'End If
    
    If Not IsTileWalkableByPlayer(GetTileType(m, p.GetX, p.GetY)) Then
        IsBlocked = True
    End If
End Function


Public Function Equals(ByRef p2 As clsPosition) As Boolean
    If m = p2.GetMap Then
        If p.GetX = p2.GetPoint.GetX And p.GetY = p2.GetPoint.GetY Then
            Equals = True
        End If
    End If
End Function

