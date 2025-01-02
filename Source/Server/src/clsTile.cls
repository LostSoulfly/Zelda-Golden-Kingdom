VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'NO USADO
Private map_ref As clsMap 'External key

Private Layer As clsVector
Private TileType As clsTileType
Private DirBlock As Byte

Public Sub Class_initialize()
    Set Layer = New clsVector
    Layer.SetSize 1
End Sub

Public Sub Tile_initialize(ByRef map As clsMap)
    Set map_ref = map
End Sub

Public Sub SetNLayers(ByVal NLayers As Long)
    Layer.Redimensionate NLayers
End Sub

Public Sub AddLayer()
'Adds a layer on the last position
    Dim NLayers As Long
    NLayers = Layer.GetSize
    Layer.Redimensionate (NLayers + 1)
End Sub

Public Sub setlayer(ByRef setlayer As clsLayer, ByVal index As Long)
    Layer.SetElem setlayer, index
End Sub
Public Function GetLayer(ByVal index As Long) As clsLayer
    Set GetLayer = Layer.GetElem(index)
End Function
