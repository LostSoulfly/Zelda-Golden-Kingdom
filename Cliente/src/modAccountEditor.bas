Attribute VB_Name = "modAccountEditor"
Option Explicit

Public EditUserIndex As Byte

Public Sub AddInfo(ByVal Text As String)

frmAccountEditor.lblInfo.Caption = Text

End Sub

Public Sub AccountEditorInit(ByVal index As Byte)
Dim i As Byte
Dim ItemName As String

With frmAccountEditor
    .FrameAccountDetails.Visible = True
    .txtUserName.Text = Trim$(Player(index).Name)
    .txtPassword.Text = Trim$(Player(index).Password)
    .txtAccess.Text = Trim$(Player(index).Access)
    .cmbClass.ListIndex = Player(index).Class - 1
    .txtSprite.Text = Player(index).Sprite

    
    'bank
    .frameBank.Visible = True
    For i = 1 To 99
        If Bank(EditUserIndex).Item(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(Item(Bank(EditUserIndex).Item(i).Num).Name)
        End If
        .lstBank.AddItem (i & ": " & ItemName & "  x  " & Bank(EditUserIndex).Item(i).Value)
    Next
    .lstBank.ListIndex = 0
    
    'inventory
    .frameInventory.Visible = True
    For i = 1 To 35
        If Player(index).Inv(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(Item(Player(index).Inv(i).Num).Name)
        End If
        .lstInventory.AddItem (i & ": " & ItemName & "  x  " & Player(index).Inv(i).Value)
    Next
    .lstInventory.ListIndex = 0
End With

End Sub

Public Sub BankEditorInit()
Dim i As Byte
Dim ItemName As String

With frmAccountEditor
    .lstBank.Clear
    For i = 1 To 99 '99 bank space
        If Bank(EditUserIndex).Item(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(Item(Bank(EditUserIndex).Item(i).Num).Name)
        End If
        .lstBank.AddItem (i & ": " & ItemName & "  x  " & Bank(EditUserIndex).Item(i).Value)
    Next
    .lstBank.ListIndex = 0
End With

End Sub

Public Sub SaveEditPlayer(ByVal index As Byte)

With Player(index)
    .Name = frmAccountEditor.txtUserName.Text
    .Password = frmAccountEditor.txtPassword.Text
    .Access = frmAccountEditor.txtAccess.Text
    .Class = frmAccountEditor.cmbClass.ListIndex + 1
    .Sprite = frmAccountEditor.txtSprite.Text
End With

Call CheckPlayerLevelUp(EditUserIndex)
Call SendPlayerData(index)

Call PlayerMsg(index, "Tu cuenta ha sido editada por el administrador.", Pink)

End Sub

Public Sub InvEditorInit()
Dim i As Byte
Dim ItemName As String

With frmAccountEditor
    'inventory
    .lstInventory.Clear
    .frameInventory.Visible = True
    For i = 1 To 35
        If Player(EditUserIndex).Inv(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(Item(Player(EditUserIndex).Inv(i).Num).Name)
        End If
        .lstInventory.AddItem (i & ": " & ItemName & "  x  " & Player(EditUserIndex).Inv(i).Value)
    Next
    .lstInventory.ListIndex = 0
End With

End Sub



