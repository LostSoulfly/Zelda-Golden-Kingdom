VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NO USADO
Private m() As Object
Private NRows As Long, NColumns As Long


Public Sub SetSize(ByVal Rows As Long, ByVal Columns As Long)
    ReDim m(Rows, Columns)
    NRows = Rows
    NColumns = Columns
End Sub

Public Sub SetElem(ByRef E As Object, ByVal X As Long, ByVal Y As Long)
    'Values X,Y must be in range
     Set m(X, Y) = E
End Sub

Public Function GetElem(ByVal X As Long, ByVal Y As Long) As Object
    Set GetElem = m(X, Y)
End Function

Public Sub Redimensionate(ByVal Rows As Long, ByVal Columns As Long)
    ReDim Preserve m(X, Y)
    NRows = Rows
    NColumns = Columns
End Sub

Public Sub GetSize(ByRef Rows As Long, ByRef Columns As Long)
    Rows = NRows
    Columns = NColumns
End Sub
