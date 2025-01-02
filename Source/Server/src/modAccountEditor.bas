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
    .txtUserName.Text = Trim$(player(index).Name)
    .txtPassword.Text = Trim$(DesEncriptatePassword((player(index).password)))
    .txtAccess.Text = Trim$(player(index).Access)
    .cmbClass.ListIndex = player(index).Class - 1
    .txtSprite.Text = player(index).Sprite

    
    'bank
    .frameBank.Visible = True
    For i = 1 To 99
        If Bank(EditUserIndex).item(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(item(Bank(EditUserIndex).item(i).Num).Name)
        End If
        .lstBank.AddItem (i & ": " & ItemName & "  x  " & Bank(EditUserIndex).item(i).Value)
    Next
    .lstBank.ListIndex = 0
    
    'inventory
    .frameInventory.Visible = True
    For i = 1 To 35
        If player(index).Inv(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(item(player(index).Inv(i).Num).Name)
        End If
        .lstInventory.AddItem (i & ": " & ItemName & "  x  " & player(index).Inv(i).Value)
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
        If Bank(EditUserIndex).item(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(item(Bank(EditUserIndex).item(i).Num).Name)
        End If
        .lstBank.AddItem (i & ": " & ItemName & "  x  " & Bank(EditUserIndex).item(i).Value)
    Next
    .lstBank.ListIndex = 0
End With

End Sub

Public Sub SaveEditPlayer(ByVal index As Byte)

With player(index)
    .Name = frmAccountEditor.txtUserName.Text
    .password = EncriptatePassword(frmAccountEditor.txtPassword.Text)
    .Access = frmAccountEditor.txtAccess.Text
    .Class = frmAccountEditor.cmbClass.ListIndex + 1
    .Sprite = frmAccountEditor.txtSprite.Text
End With

Call CheckPlayerLevelUp(EditUserIndex)
Call SendPlayerData(index)

Call PlayerMsg(index, "Your account has been edited by an Admin.", Pink)

End Sub

Public Sub InvEditorInit()
Dim i As Byte
Dim ItemName As String

With frmAccountEditor
    'inventory
    .lstInventory.Clear
    .frameInventory.Visible = True
    For i = 1 To 35
        If player(EditUserIndex).Inv(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(item(player(EditUserIndex).Inv(i).Num).Name)
        End If
        .lstInventory.AddItem (i & ": " & ItemName & "  x  " & player(EditUserIndex).Inv(i).Value)
    Next
    .lstInventory.ListIndex = 0
End With

End Sub



