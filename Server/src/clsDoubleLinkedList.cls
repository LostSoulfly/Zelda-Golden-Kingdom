VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDoubleLinkedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'NO USADO
Private MAX_LIST_ITEMS As Long

Private Type listpair
    E As Variant
    next As Integer
    previous As Integer
End Type


Private a() As listpair
Private prev As Integer
Private free As Integer




Public Sub Create(ByVal max As Long)
    MAX_LIST_ITEMS = max
    prev = 1
    ReDim a(1 To MAX_LIST_ITEMS + 1)
    MAX_LIST_ITEMS = max
    a(prev).next = -1
    a(prev).previous = -1
    Dim i As Integer
    For i = 2 To MAX_LIST_ITEMS
        a(i).next = i + 1
        a(i).previous = i - 1
    Next
    a(MAX_LIST_ITEMS + 1).previous = MAX_LIST_ITEMS
    a(MAX_LIST_ITEMS + 1).next = -1
    free = 2
End Sub

Public Sub Insert(ByVal E As Object)
    Dim tmp As Integer
    If free <> -1 Then
        tmp = free
        free = a(free).next
        a(tmp).E = E
        a(tmp).next = a(prev).next
        a(prev).next = tmp
        prev = tmp
    End If
End Sub

Public Sub Delete()
    Dim tmp As Integer
    If a(prev).next <> -1 Then
        tmp = a(prev).next
        a(prev).next = a(tmp).next
        a(tmp).next = free
        free = tmp
    End If
End Sub

Public Sub Begin()
    prev = 1
End Sub

Public Sub LNext()
    If a(prev).next <> -1 Then
        prev = a(prev).next
    End If
End Sub

Public Sub LPrev()
    
End Sub

Public Function Actual() As Object
    If a(prev).next <> -1 Then
        Actual = a(a(prev).next).E
    End If
End Function

Public Function LEnd() As Boolean
    LEnd = (a(prev).next = -1)
End Function

Public Function LEmpty() As Boolean
    LEmpty = a(1).next = -1
End Function

Public Function Full() As Boolean
    Full = (free = -1)
End Function




