VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NO USADO
Private Name As String * NAME_LENGTH
Private Music As String * NAME_LENGTH
    
Private Revision As Long
Private moral As Byte

Private Weather As Long ' 0 = None, 1 = Rain 2, = Snow, 3 = Sandstorm
    
Private Boundries As clsVector
Private Boot As clsPosition
Private Tile As clsMatrix
Private NPC As clsVector
Private NPCSProperties As clsVector


Public Sub Class_initialize()

End Sub




