VERSION 5.00
Begin VB.Form frmAccountEditor 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSaveAcc 
      Caption         =   "Save Acc"
      Height          =   255
      Left            =   1560
      TabIndex        =   46
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Acc"
      Height          =   255
      Left            =   1560
      TabIndex        =   45
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame FrameStats 
      Caption         =   "Char Data"
      Height          =   2415
      Left            =   240
      TabIndex        =   30
      Top             =   4320
      Width           =   2055
      Begin VB.TextBox txtEExp 
         Height          =   285
         Left            =   480
         TabIndex        =   37
         Text            =   "0"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtEStr 
         Height          =   285
         Left            =   480
         TabIndex        =   36
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtEEnd 
         Height          =   285
         Left            =   1440
         TabIndex        =   35
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtEInt 
         Height          =   285
         Left            =   480
         TabIndex        =   34
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtEAgi 
         Height          =   285
         Left            =   1440
         TabIndex        =   33
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtEWill 
         Height          =   285
         Left            =   1440
         TabIndex        =   32
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtEPts 
         Height          =   285
         Left            =   480
         TabIndex        =   31
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Exp:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Str:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "End:"
         Height          =   255
         Left            =   960
         TabIndex        =   42
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Int:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Agi:"
         Height          =   255
         Left            =   960
         TabIndex        =   40
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Will:"
         Height          =   255
         Left            =   960
         TabIndex        =   39
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Pts:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.CommandButton mnuLevel 
      Caption         =   "Change Level"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Frame frameInventory 
      Caption         =   "Inventory"
      Height          =   6735
      Left            =   6120
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton cmdSaveInventory 
         Caption         =   "Save"
         Height          =   375
         Left            =   840
         TabIndex        =   25
         Top             =   6240
         Width           =   1455
      End
      Begin VB.HScrollBar scrlInvItem 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   5520
         Width           =   3015
      End
      Begin VB.TextBox txtAmountInv 
         Height          =   285
         Left            =   960
         TabIndex        =   23
         Text            =   "0"
         Top             =   5880
         Width           =   2175
      End
      Begin VB.ListBox lstInventory 
         Height          =   4935
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label13 
         Caption         =   "Ammount:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   5880
         Width           =   3015
      End
      Begin VB.Label lblInvItem 
         Caption         =   "Inv item: None"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   5280
         Width           =   3015
      End
   End
   Begin VB.Frame frameBank 
      Caption         =   "Bank"
      Height          =   6735
      Left            =   2640
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton cmdSaveBank 
         Caption         =   "Save"
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   6240
         Width           =   1455
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Text            =   "0"
         Top             =   5880
         Width           =   2175
      End
      Begin VB.HScrollBar scrlBankItem 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   5520
         Width           =   3015
      End
      Begin VB.ListBox lstBank 
         Height          =   4935
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "Ammount:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   5880
         Width           =   3135
      End
      Begin VB.Label lblBankItem 
         Caption         =   "Bank item: None"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   5280
         Width           =   3015
      End
   End
   Begin VB.Frame FrameAccountDetails 
      Caption         =   "Account Details"
      Height          =   2655
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox txtSprite 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ComboBox cmbClass 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtAccess 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Sprite: "
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Class:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Access:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdSavePlayer 
      Caption         =   "Save Player"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdFindPlayer 
      Caption         =   "Find Player"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtUserNameLoad 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   600
      TabIndex        =   28
      Top             =   6720
      Width           =   8535
   End
End
Attribute VB_Name = "frmAccountEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFindPlayer_Click()
Dim Username As String
Dim i As Byte

Username = Trim$(txtUserNameLoad.Text)
lstBank.Clear
lstInventory.Clear
frameBank.Visible = False
FrameAccountDetails.Visible = False
frameInventory.Visible = False
FrameStats.Visible = True
mnuLevel.Visible = True

For i = 1 To Player_HighIndex
    If IsPlaying(i) = True Then
        If LCase$(Trim$(player(i).name)) = LCase$(Username) Then
            EditUserIndex = i
            Call AccountEditorInit(i)
        ElseIf AccountExist(Username) Then
            AddInfo ("User is offline.")
            cmdLoad_Click
            Else
            AddInfo "User does not exist!"
        End If
    End If
Next

End Sub

Private Sub cmdLoad_Click()
Username = Trim$(txtUserNameLoad.Text)
lstBank.Clear
lstInventory.Clear
frameBank.Visible = False
FrameAccountDetails.Visible = False
frameInventory.Visible = False
FrameStats.Visible = True
mnuLevel.Visible = True

If Not AccountExist(Username) Then Exit Sub
If IsPlaying(MAX_PLAYERS) Then Exit Sub

ClearPlayer MAX_PLAYERS
ClearBank MAX_PLAYERS
LoadPlayer MAX_PLAYERS, Username
LoadBank MAX_PLAYERS, Username
EditUserIndex = MAX_PLAYERS
Call AccountEditorInit(MAX_PLAYERS)


End Sub

Private Sub cmdSaveAcc_Click()
SavePlayer EditUserIndex
'SendPlayerData EditUserIndex
Call SaveEditPlayer(EditUserIndex)

With player(EditUserIndex)
    .name = frmAccountEditor.txtUserName.Text
    .password = EncriptatePassword(frmAccountEditor.txtPassword.Text)
    .Access = frmAccountEditor.txtAccess.Text
    .Class = frmAccountEditor.cmbClass.ListIndex + 1
    .Sprite = frmAccountEditor.txtSprite.Text
End With

Call CheckPlayerLevelUp(EditUserIndex)
'Call SendPlayerData(index)

SavePlayer EditUserIndex

ClearPlayer MAX_PLAYERS
ClearBank MAX_PLAYERS

End Sub

Private Sub cmdSaveBank_Click()

If IsPlaying(EditUserIndex) = False Then
    Call AddInfo("¡El jugador no está en línea!")
    Exit Sub
End If

Bank(EditUserIndex).item(lstBank.ListIndex + 1).Num = scrlBankItem.Value
Bank(EditUserIndex).item(lstBank.ListIndex + 1).Value = txtAmount.Text

Call SaveBank(EditUserIndex)
Call BankEditorInit

End Sub

Private Sub cmdSaveInventory_Click()

If IsPlaying(EditUserIndex) = False Then
    Call AddInfo("¡El jugador no está en línea!")
    Exit Sub
End If

player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Num = scrlInvItem.Value
player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Value = txtAmountInv.Text

Call SendInventoryUpdate(EditUserIndex, lstInventory.ListIndex + 1)

Call InvEditorInit

End Sub

Private Sub cmdSavePlayer_Click()

If IsPlaying(EditUserIndex) = False Then
    AddInfo ("¡El jugador no está en línea!")
    Exit Sub
End If
Dim i As Long
    i = (EditUserIndex)
    If Not (EditUserIndex) = 0 Then
        SavePlayer i
        SendPlayerData i
    Else
        Call MsgBox("Player not found!", vbOKOnly)
    End If
Call SaveEditPlayer(EditUserIndex)

End Sub

Private Sub Form_Load()
Dim i As Byte

scrlBankItem.max = MAX_ITEMS
scrlInvItem.max = MAX_ITEMS

cmbClass.Text = Trim$(Class(1).name)
For i = 1 To Max_Classes
    cmbClass.AddItem Trim$(Class(i).name)
Next

End Sub

Private Sub lstInventory_Click()
Dim ItemName As String

If player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Num = 0 Then
    ItemName = "None"
Else
    ItemName = Trim$(item(player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Num).name)
End If

lblInvItem.Caption = "Inv item: " & ItemName
txtAmountInv.Text = player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Value
scrlInvItem.Value = player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Num

End Sub

Private Sub lstBank_Click()
Dim ItemName As String

If Bank(EditUserIndex).item(lstBank.ListIndex + 1).Num = 0 Then
    ItemName = "None"
Else
    ItemName = Trim$(item(Bank(EditUserIndex).item(lstBank.ListIndex + 1).Num).name)
End If

lblBankItem.Caption = "Bank item: " & ItemName
txtAmount.Text = Bank(EditUserIndex).item(lstBank.ListIndex + 1).Value
scrlBankItem.Value = Bank(EditUserIndex).item(lstBank.ListIndex + 1).Num

End Sub
Private Sub mnuLevel_Click()
Dim level As Integer
Dim Player_Level As Integer
level = InputBox("Level 1-100:", "Level")
Dim name As String
name = frmServer.lvwInfo.SelectedItem.SubItems(3)
If Not name = "Not Playing" Then

Player_Level = GetPlayerLevel(FindPlayer(name))
' If you want to change points please pm me <img src='http://www.touchofdeathforums.com/community/public/style_emoticons/<#EMO_DIR#>/wink.png' class='bbc_emoticon' alt=';)' />
Call SetPlayerLevel(FindPlayer(name), level)
Call SendPlayerData(FindPlayer(name))
Call PlayerMsg(FindPlayer(name), GetTranslation("Te han cambiado tu nivel") & " " & Player_Level & " " & GetTranslation("al nivel") & " " & level, BrightCyan, , False)
End If
End Sub

Private Sub scrlBankItem_Change()

If scrlBankItem.Value = 0 Then
    lblBankItem.Caption = "Bank item: None"
Else
    lblBankItem.Caption = "Bank item: " & item(scrlBankItem.Value).name
End If

End Sub

Private Sub scrlInvItem_Change()

If scrlInvItem.Value = 0 Then
    lblInvItem.Caption = "Inv item: None"
Else
    lblInvItem.Caption = "Inv item: " & item(scrlInvItem.Value).name
End If

End Sub

Private Sub txtAccess_Change()

If IsNumeric(txtAccess.Text) = False Then txtAccess.Text = player(EditUserIndex).Access

End Sub

Private Sub txtAmountInv_Change()

If IsNumeric(txtAmountInv.Text) = False Then txtAmountInv.Text = player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Value
If txtAmountInv.Text > 2000000000 Then txtAmountInv.Text = 2000000000

End Sub

Private Sub txtPassword_Change()

If txtPassword.Text = vbNullString Then txtPassword.Text = DesEncriptatePassword(player(EditUserIndex).password)

End Sub

Private Sub txtSprite_Change()

If IsNumeric(txtSprite.Text) = False Then txtSprite.Text = player(edituseindex).Sprite

End Sub

Private Sub txtUserName_Change()

If txtUserName.Text = vbNullString Then txtUserName.Text = player(EditUserIndex).name

End Sub

Private Sub txtAmount_Change()

If IsNumeric(txtAmount.Text) = False Then txtAmount.Text = Bank(EditUserIndex).item(lstBank.ListIndex + 1).Value
If txtAmount.Text > 2000000000 Then txtAmount.Text = 2000000000

End Sub
