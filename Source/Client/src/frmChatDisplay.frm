VERSION 5.00
Begin VB.Form frmChatDisplay 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Display"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   2715
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkActivateChat 
      BackColor       =   &H0080FF80&
      Caption         =   "Anti-Scroll Down"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CheckBox chkChat 
      BackColor       =   &H0080FF80&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CheckBox chkChat 
      BackColor       =   &H0080FF80&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CheckBox chkChat 
      BackColor       =   &H0080FF80&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CheckBox chkChat 
      BackColor       =   &H0080FF80&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.CheckBox chkChat 
      BackColor       =   &H0080FF80&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.CheckBox chkChat 
      BackColor       =   &H0080FF80&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmChatDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkActivateChat_Click()
    HideIncomingMessages = ITB(chkActivateChat.value)
End Sub

Private Sub chkChat_Click(index As Integer)
    Options.ActivatedChats(index + 1) = chkChat(index).value
End Sub

Private Sub cmdAccept_Click()
    ClearChat
    SaveOptions
End Sub

Private Sub Form_Load()
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
