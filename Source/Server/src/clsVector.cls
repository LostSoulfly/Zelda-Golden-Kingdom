VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'NO USADO
Private v() As Object
Private vsize As Long


Public Sub SetSize(ByVal size As Long)
    ReDim v(size)
    vsize = size
End Sub

Public Sub SetElem(ByVal E As Object, ByVal i As Long)
    'Values X,Y must be in range
    Set v(i) = E
End Sub

Public Sub DeleteElem(ByVal i As Long)
    Set v(i) = Nothing
End Sub

Public Function GetElem(ByVal i As Long) As Object
    Set GetElem = v(i)
End Function

Public Sub Redimensionate(ByVal size As Long)
    ReDim Preserve v(size)
    vsize = size
End Sub


Public Function GetSize() As Long
    GetSize = vsize
End Function

Public Function ToStack() As clsStack
    Set ToStack = New clsStack
    
    Dim i As Long
    For i = vsize To 0 Step -1
        ToStack.Push v(i)
    Next
End Function

Public Sub Unificate()
    Dim i As Long
    i = 0
    While i < vsize
        If v(i) Is Nothing Then
            Dim j As Long
            For j = i To vsize - 1
                Set v(j) = v(j + 1)
            Next
            vsize = vsize - 1
            Me.Redimensionate (vsize)
        Else
            i = i + 1
        End If
    Wend
    
    
End Sub
