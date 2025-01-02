VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'NO USADO
Private index As Long
Private List As clsObjectList

Public Sub CreateIterator(ByRef l As clsObjectList)
    Set List = l
End Sub

Public Sub SetIndex(ByVal i As Long)
    index = i
End Sub

Public Function GetIndex() As Long
    GetIndex = index
End Function

Public Sub NextIndex()
    List.IteratorNext Me
End Sub

Public Sub PrevIndex()
    List.IteratorPrev Me
End Sub

Public Sub SetToBegin()
    List.IteratorBegin Me
End Sub

Public Sub SetToEnd()
    List.IteratorEnd Me
End Sub

Public Function IsOnEnd() As Boolean
    IsOnEnd = List.IsIteratorEnd(Me)
End Function

Public Function IsOnBegin() As Boolean
    IsOnBegin = List.IsIteratorBegin(Me)
End Function

Public Function GetMyItem() As Object
    Set GetMyItem = List.GetItem(Me)
End Function


