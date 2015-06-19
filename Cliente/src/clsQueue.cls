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

Private A(1 To MAX_BYTE) As Variant
Private first As Integer
Private last As Integer
Private N As Integer

Private Sub Class_Initialize()
    first = 1
    last = 0
    N = 0
End Sub

Public Sub Push(ByRef s As Variant)
    If N <> MAX_BYTE Then
        last = (last Mod MAX_BYTE) + 1
        A(last) = s
        N = N + 1
    End If
End Sub

Public Sub Pop()
    If N > 0 Then
        first = (first Mod MAX_BYTE) + 1
        N = N - 1
    End If
End Sub

Public Function Front() As Variant
    If N > 0 Then
        Front = A(first)
    End If
End Function

Public Function IsEmpty() As Boolean
    IsEmpty = (N = 0)
End Function

Public Function IsFull() As Boolean
    IsFull = (N = MAX_BYTE)
End Function



