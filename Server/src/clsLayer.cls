VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'NO USADO
Private X As Long
Private Y As Long
Private Tileset As Long

Public Sub SetX(ByVal SetX As Long)
    X = SetX
End Sub

Public Sub SetY(ByVal SetY As Long)
    Y = SetY
End Sub

Public Sub SetTileSet(ByVal SetTileSet As Long)
    Tileset = SetTileSet
End Sub

Public Function GetX() As Long
    GetX = X
End Function

Public Function GetY() As Long
    GetY = Y
End Function

Public Function GetTileSet() As Long
    GetTileSet = Tileset
End Function

