Attribute VB_Name = "modList"
Option Explicit

Private Const MAX_LIST_ITEMS As Long = MAX_PLAYERS

Private Type listpair
    E As Variant
    next As Integer
End Type

Public Type List
    a(1 To MAX_LIST_ITEMS + 1) As listpair
    prev As Integer
    free As Integer
End Type


Public Sub ListCreate(ByRef l As List)
    l.prev = 1
    l.a(l.prev).next = -1
    Dim i As Integer
    For i = 2 To MAX_LIST_ITEMS
        l.a(i).next = i + 1
    Next
    l.a(MAX_LIST_ITEMS + 1).next = -1
    l.free = 2
End Sub

Public Sub ListInsert(ByRef l As List, ByVal E As Variant)
    Dim tmp As Integer
    If l.free <> -1 Then
        tmp = l.free
        l.free = l.a(l.free).next
        l.a(tmp).E = E
        l.a(tmp).next = l.a(l.prev).next
        l.a(l.prev).next = tmp
        l.prev = tmp
    End If
End Sub

Public Sub ListDelete(ByRef l As List)
    Dim tmp As Integer
    If l.a(l.prev).next <> -1 Then
        tmp = l.a(l.prev).next
        l.a(l.prev).next = l.a(tmp).next
        l.a(tmp).next = l.free
        l.free = tmp
    End If
End Sub

Public Sub ListBegin(ByRef l As List)
    l.prev = 1
End Sub

Public Sub ListNext(ByRef l As List)
    If l.a(l.prev).next <> -1 Then
        l.prev = l.a(l.prev).next
    End If
End Sub

Public Function ListActual(ByRef l As List) As Variant
    If l.a(l.prev).next <> -1 Then
        ListActual = l.a(l.a(l.prev).next).E
    End If
End Function

Public Function ListEnd(ByRef l As List) As Boolean
    ListEnd = (l.a(l.prev).next = -1)
End Function

Public Function ListEmpty(ByRef l As List) As Boolean
    ListEmpty = l.a(1).next = -1
End Function

Public Function ListFull(ByRef l As List) As Boolean
    ListFull = (l.free = -1)
End Function

Public Sub SetListEnd(ByRef l As List)
    If Not ListEmpty(l) Then
        Call ListBegin(l)
        While Not (ListEnd(l))
            Call ListNext(l)
        Wend
    End If
End Sub

Public Sub ListPush(ByRef l As List, ByRef E As Variant)
    If ListFull(l) Then
        ListBegin l
        ListDelete l
    End If
    
    SetListEnd l
    ListInsert l, E
End Sub

Public Sub ListPop(ByRef l As List)
    If Not ListEmpty(l) Then
        ListBegin l
        ListDelete l
    End If
End Sub


Function SearchItem(ByRef l As List, ByVal E As Variant) As Boolean
    ListBegin l
    While Not ListEnd(l) And ListActual(l) <> E
        ListNext l
    Wend
    
    If Not ListEnd(l) Then
        SearchItem = True
    End If
End Function
