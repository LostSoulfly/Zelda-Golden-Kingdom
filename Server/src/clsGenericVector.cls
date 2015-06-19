VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenericVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NO USADO
Private v() As Variant
Private vsize As Long


Public Sub SetSize(ByVal size As Long)
    ReDim v(size)
    vsize = size
End Sub

Public Sub SetElem(ByRef E As Variant, ByVal i As Long)
    'Values X,Y must be in range
    v(i - 1) = E
End Sub

Public Function GetElem(ByVal i As Long) As Variant
    Set GetElem = v(i - 1)
End Function

Public Sub Redimensionate(ByVal size As Long)
    ReDim Preserve v(size)
    vsize = size
End Sub

Public Function GetSize() As Long
    GetSize = vsize
End Function

