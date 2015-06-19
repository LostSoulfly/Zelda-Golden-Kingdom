VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTileAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NO USADO

Public Enum AttributeType
 TILE_TYPE_WALKABLE = 0
 TILE_TYPE_BLOCKED = 1
 TILE_TYPE_WARP = 2
 TILE_TYPE_ITEM = 3
 TILE_TYPE_NPCAVOID = 4
 TILE_TYPE_key = 5
 TILE_TYPE_KEYOPEN = 6
 TILE_TYPE_RESOURCE = 7
 TILE_TYPE_DOOR = 8
 TILE_TYPE_NPCSPAWN = 9
 TILE_TYPE_SHOP = 10
 TILE_TYPE_BANK = 11
 TILE_TYPE_HEAL = 12
 TILE_TYPE_TRAP = 13
 TILE_TYPE_SLIDE = 14
 TILE_TYPE_SCRIPT = 15
 TILE_TYPE_ICE = 16
End Enum

Private TileType As AttributeType


Public Property Let Tile(ByVal Attr As AttributeType)
    TileType = Attr
End Property

Public Property Get Tile() As AttributeType
    Tile = TileType
End Property

Public Sub StepTile(ByRef player As Long) ' clsplayer

End Sub


