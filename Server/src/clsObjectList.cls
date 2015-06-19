VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NO USADO
Private Items As Collection

Public Sub Class_initialize()
    Set Items = New Collection
End Sub

Public Function IsEmpty() As Boolean
    If Items.Count = 0 Then
        IsEmpty = True
    End If
End Function


Public Sub IteratorBegin(ByRef It As clsListIterator)
    It.SetIndex 1
End Sub

Public Sub IteratorEnd(ByRef It As clsListIterator)
    It.SetIndex Items.Count
End Sub

Public Sub IteratorNext(ByRef It As clsListIterator)
    It.SetIndex It.GetIndex + 1
End Sub

Public Sub IteratorPrev(ByRef It As clsListIterator)
    It.SetIndex It.GetIndex - 1
End Sub

Public Function GetItem(ByRef It As clsListIterator) As Object
    GetItem = Items.item(It.GetIndex)
End Function

Public Sub DeleteItem(ByRef It As clsListIterator)
    Items.Remove (It.GetIndex)
End Sub

Public Function IsIteratorEnd(ByRef It As clsListIterator)
    If It.GetIndex = Items.Count Then
        IsIteratorEnd = True
    End If
End Function

Public Function IsIteratorBegin(ByRef It As clsListIterator)
    If It.GetIndex = 1 Then
        IsIteratorBegin = True
    End If
End Function

