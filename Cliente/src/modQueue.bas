Attribute VB_Name = "modMsgList"
Public Type ChatMsgRec
    header As String
    colour As Long
    text As String
    saycolour As Long
    ArrivedAt As Long
End Type

Private Type listpair
    e As ChatMsgRec
    next As Integer
End Type

Public Type list
    A(1 To MAX_CHAT_LINES + 1) As listpair
    prev As Integer
    free As Integer
End Type


Public Sub ListCreate(ByRef l As list)
    l.prev = 1
    l.A(l.prev).next = -1
    Dim i As Integer
    For i = 2 To MAX_CHAT_LINES
        l.A(i).next = i + 1
    Next
    l.A(MAX_CHAT_LINES + 1).next = -1
    l.free = 2
End Sub

Public Sub ListInsert(ByRef l As list, ByRef e As ChatMsgRec)
    Dim tmp As Integer
    If l.free <> -1 Then
        tmp = l.free
        l.free = l.A(l.free).next
        l.A(tmp).e = e
        l.A(tmp).next = l.A(l.prev).next
        l.A(l.prev).next = tmp
        l.prev = tmp
    End If
End Sub

Public Sub ListDelete(ByRef l As list)
    Dim tmp As Integer
    If l.A(l.prev).next <> -1 Then
        tmp = l.A(l.prev).next
        l.A(l.prev).next = l.A(tmp).next
        l.A(tmp).next = l.free
        l.free = tmp
    End If
End Sub

Public Sub ListBegin(ByRef l As list)
    l.prev = 1
End Sub

Public Sub ListNext(ByRef l As list)
    If l.A(l.prev).next <> -1 Then
        l.prev = l.A(l.prev).next
    End If
End Sub

Public Function ListActual(ByRef l As list) As ChatMsgRec
    If l.A(l.prev).next <> -1 Then
        ListActual = l.A(l.A(l.prev).next).e
    End If
End Function

Public Function ListEnd(ByRef l As list) As Boolean
    ListEnd = (l.A(l.prev).next = -1)
End Function

Public Function ListEmpty(ByRef l As list) As Boolean
    ListEmpty = l.A(1).next = -1
End Function

Public Function ListFull(ByRef l As list) As Boolean
    ListFull = (l.free = -1)
End Function

Public Sub SetListEnd(ByRef l As list)
    If Not ListEmpty(l) Then
        Call ListBegin(l)
        While Not (ListEnd(l))
            Call ListNext(l)
        Wend
    End If
End Sub

Public Sub ListPush(ByRef l As list, ByRef e As ChatMsgRec)
    If ListFull(l) Then
        ListBegin l
        ListDelete l
    End If
    
    SetListEnd l
    ListInsert l, e
End Sub

