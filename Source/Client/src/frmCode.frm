VERSION 5.00
Begin VB.Form frmCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Redeem Code"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Enter the code here"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label lblCash 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Balance: X"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAccept_Click()

    If Len(txtCode.text) < 50 And Not txtCode.text = vbNullString Then
        SendCode txtCode.text
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    lblCash.Caption = "Current Balance: " & GetBonusPoints(MyIndex) & " " & CURRENCY_NAME
    
    Dim e As Control
    
    For Each e In Me.Controls
        If (TypeOf e Is Label) Then
            'e.caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is CheckBox) Then
            'e.caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is OptionButton) Then
            'e.caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is Frame) Then
            'e.caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is CommandButton) Then
            'e.caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is TextBox) Then
            'e.text = GetTranslation(e.text)
        End If
    Next
    
    'me.caption = GetTranslation(Me.Caption)
    
End Sub
