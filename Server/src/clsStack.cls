VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NO USADO
Private Stack() As Collection

Private Top As Long
Private ArrayTop As Long

Public Sub Class_initialize()
    ArrayTop = 0
    ReDim Stack(ArrayTop) As Collection
    Set Stack(ArrayTop) = New Collection
    Top = 0
End Sub

Public Sub Push(ByVal E As Object)
    If ArrayFull Then
        AddArrayStack
        Top = 0
    End If
    Stack(ArrayTop).Add E
    Top = Top + 1
End Sub

Private Function ArrayFull() As Boolean
    If Stack(ArrayTop).Count >= 256 Then
        ArrayFull = True
    End If
End Function

Private Function ArrayEmpty() As Boolean
    If Stack(ArrayTop).Count = 0 Then
        ArrayEmpty = True
    End If
End Function

Private Sub AddArrayStack()
    ArrayTop = ArrayTop + 1
    ReDim Preserve Stack(ArrayTop)
    Set Stack(ArrayTop) = New Collection
End Sub

Private Sub DeleteArrayStack()
    ArrayTop = ArrayTop - 1
    ReDim Preserve Stack(ArrayTop)
End Sub
Public Sub Pop()
    If Top = 0 Then Exit Sub
    
    Stack(ArrayTop).Remove Top
    Top = Top - 1
    
    If ArrayEmpty Then
        If ArrayTop > 0 Then
            DeleteArrayStack
            Top = 256
        End If
    End If
        
        
    
    
    
End Sub

Public Function Front() As Object
    If Top = 0 Then
        Set Front = Nothing
    Else
        Set Front = Stack(ArrayTop).item(Top)
    End If
End Function

Public Function IsEmpty() As Boolean
    IsEmpty = (Top = 0)
End Function

Public Sub Clear()
    Me.Class_initialize
End Sub

Public Function ToVector() As clsVector
    Dim Copy As clsStack
    Set Copy = New clsStack
    Set Copy = Me
    
    If Top = 0 Then
        Set ToVector = Nothing
        Exit Function
    End If
    
    Set ToVector = New clsVector
    
    
    ToVector.SetSize ArrayTop * 256 + Top - 1
    
    Dim i As Long
    i = 0
    
    While Not Copy.IsEmpty
        ToVector.SetElem Copy.Front, i
        i = i + 1
        Copy.Pop
    Wend
    
End Function






