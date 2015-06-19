VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NO USADO
Private first As Variant
Private Second As Variant

Public Sub SetFirst(ByVal v As Variant)
    first = v
End Sub

Public Function GetFirst() As Variant
    GetFirst = first
End Function

Public Sub SetSecond(ByVal v As Variant)
    Second = v
End Sub

Public Function GetSecond() As Variant
    GetSecond = Second
End Function
