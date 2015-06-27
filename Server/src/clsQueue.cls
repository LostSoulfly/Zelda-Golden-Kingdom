VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'NO USADO
Private qFront As clsQueueItem
Private qRear As clsQueueItem
Public Sub Class_initialize()

End Sub

Public Sub Push(ByRef E As Variant)
    Dim qNew As clsQueueItem
    Set qNew = New clsQueueItem
    Set qNew.E = E
    ' What if the queue is empty? Better point
    ' both the front and rear pointers at the
    ' new item.
    If IsEmpty Then
        Set qFront = qNew
        Set qRear = qNew
    Else
        Set qRear.NextItem = qNew
        Set qRear = qNew
    End If
End Sub

Property Get IsEmpty() As Boolean
    ' Return True if the queue contains
    ' no items.
    IsEmpty = ((qFront Is Nothing) And (qRear Is Nothing))
End Property


Public Sub Remove()
    ' Remove an item from the head of the
    ' list, and return its value.
    If Not IsEmpty Then
        ' If there's only one item
        ' in the queue, qFront and qRear
        ' will be pointing to the same node.
        ' Use the Is operator to test for that.
        If qFront Is qRear Then
            Set qFront = Nothing
            Set qRear = Nothing
        Else
            Set qFront = qFront.NextItem
        End If
    End If
End Sub

Public Function Front() As Variant
    If Not IsEmpty Then
        Set Front = qFront.E
    
    End If
End Function
