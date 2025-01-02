VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDoubleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private X As Integer
Private Y As Integer
Private Z As Integer

Private ContainerFrame As Frame
Private Lists(1 To 2) As ListBox
Private AddButton As CommandButton
Private WithdrawButton As CommandButton

Private FHeight As Long
Private FWidth As Long

Public Sub Class_Initialize()

End Sub

Public Sub DrawExistentFrame(ByRef Frame As Frame, ByRef list1 As ListBox, ByRef List2 As ListBox, ByRef cmd1 As CommandButton, ByRef cmd2 As CommandButton)
    Set ContainerFrame = Frame
    Set Lists(1) = list1
    Set Lists(2) = List2
    Set AddButton = cmd1
    Set WithdrawButton = cmd2
    Me.Draw
End Sub

Public Sub AddButton_click()
    If Lists(1).ListIndex > 0 Then
        Call Add(Lists(1).list(Lists(1).ListIndex), 2)
        Call Delete(Lists(1).ListIndex, 1)
    End If
End Sub

Public Sub WithdrawButton_Click()
    If Lists(2).ListIndex > 0 Then
        Call Add(Lists(2).list(Lists(2).ListIndex), 1)
        Call Delete(Lists(2).ListIndex, 2)
    End If
End Sub
Public Sub SetSize(ByVal Height As Long, ByVal Width As Long)
    FHeight = Height
    FWidth = Width
    X = FWidth / 7
    Y = X
    Z = 2 * X
End Sub


Public Sub Draw()
    Dim i As Long
    For i = 1 To 2
        Lists(i).Width = FWidth - 2 * X
        Lists(i).Height = (FHeight - 3 * Y) / 2
    Next
    
    
    Lists(1).Top = Y
    Lists(1).Left = X
    
    Lists(2).Top = Lists(1).Top + Lists(1).Height + Y
    Lists(2).Left = Lists(1).Left
    
    AddButton.Width = (Lists(1).Width - Z) / 2
    AddButton.Height = Y
    AddButton.Caption = "V"
    
    WithdrawButton.Width = (Lists(1).Width - Z) / 2
    WithdrawButton.Height = Y
    WithdrawButton.Caption = "^"
    
    AddButton.Top = (Lists(2).Top - Y)
    AddButton.Left = Lists(1).Left
    
    WithdrawButton.Top = (Lists(2).Top - Y)
    WithdrawButton.Left = AddButton.Left + AddButton.Width + Z
      
End Sub

Public Sub Add(ByRef elem As String, ByVal list As Byte)
    Lists(list).AddItem elem
End Sub

Public Sub Delete(ByVal index As Long, ByVal list As Byte)
    Lists(list).RemoveItem (index)
End Sub

Public Sub Traspass(ByVal index As Long)
    Dim elem As String
    elem = Lists(1).list(index)
    Delete index, 1
    Add elem, 2
End Sub

Public Sub Withdraw(ByVal index As Long)
    Dim elem As String
    elem = Lists(2).list(index)
    Delete index, 2
    Add elem, 1
End Sub

Private Function Search(ByRef elem As String, ByVal list As Byte) As Long
    Dim i As Long
    Search = -1
    For i = 0 To Lists(list).ListCount
        If Lists(list).list(i) = elem Then
            Search = i
            Exit Function
        End If
    Next
End Function
