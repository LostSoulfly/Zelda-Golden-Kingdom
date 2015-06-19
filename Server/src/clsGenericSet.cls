VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenericSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NO USADO

Private col As Collection


Private Sub Class_initialize()
    Set col = New Collection
End Sub

Public Sub Add(ByRef item As Variant)
    col.Add item
End Sub

Public Sub Delete(ByRef item As Variant)
    Dim i As Long
    i = FindIndex(item)
    If i > 0 And i <= col.Count Then
        col.Remove (i)
    End If
End Sub

Public Function Exists(ByRef item As Variant) As Boolean
    If FindIndex(item) > 0 Then
        Exists = True
    End If
End Function

Private Function FindIndex(ByRef item As Variant) As Long
    Dim i As Long
    For i = 1 To col.Count
        If col.item(i) = item Then
            FindIndex = i
            Exit Function
        End If
    Next
End Function
