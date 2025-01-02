VERSION 5.00
Begin VB.Form frmTriForce 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picTriforce 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   120
      Picture         =   "frmTriForce.frx":0000
      ScaleHeight     =   4500
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.Label lblTriforceWisdom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wisdom"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lblTriforceCourage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Courage"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label lblTriforcePower 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Power"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblTriforceInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Triforce"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   4080
         Width           =   3015
      End
      Begin VB.Label lblTriforceAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   4080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblTriforceDecline 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   5040
         TabIndex        =   1
         Top             =   4080
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmTriForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim e As Control
        
    For Each e In Me.Controls
        If (TypeOf e Is Label) Then
            ''e.caption = GetTranslation(e.Caption)
            e.Visible = True
        End If
        If (TypeOf e Is CheckBox) Then
            ''e.caption = GetTranslation(e.Caption)
            e.Visible = True
        End If
        If (TypeOf e Is OptionButton) Then
            ''e.caption = GetTranslation(e.Caption)
            e.Visible = True
        End If
    Next
    
End Sub


Private Sub lblTriforceAccept_Click()
    If SelectedTriforce > 0 And SelectedTriforce < TriforceType.TriforceType_Count Then
        Dim message As String
        message = "Warning: This will reset the following: " & vbNewLine
        message = message & "Level" & vbNewLine
        message = message & "Stats Points" & vbNewLine
        message = message & "Experience" & vbNewLine
        message = message & "Inventory" & vbNewLine
        message = message & "Equipment" & vbNewLine
        message = message & "Quests" & vbNewLine
        message = message & "Spells" & vbNewLine
        message = message & "Rupee Bags" & vbNewLine
        message = message & "PK status" & vbNewLine
        
        If MsgBox(message, vbYesNoCancel, "Triforce") = vbYes Then
            Call SendResetPlayer(MyIndex, SelectedTriforce)
        End If
    End If
    
    ResetTriforcePicInfo True
    picTriforce.Visible = False
    'play sound
    PlaySound Sound_ButtonTriforce
    
    DoEvents
    
    Me.Visible = False
    Unload frmTriForce
    
End Sub

Private Sub lblTriforceClose_Click()
    ResetTriforcePicInfo True
    picTriforce.Visible = False
    'play sound
    PlaySound Sound_ButtonClose
End Sub

Private Sub ResetTriforcePicInfo(ByVal Control As Boolean)
If Control Then
    lblTriforceAccept.Visible = False
    lblTriforceDecline.Visible = False
    lblTriforceInfo.Caption = "Select a Triforce"
Else
    lblTriforceAccept.Visible = True
    lblTriforceDecline.Visible = True
    lblTriforceInfo.Caption = "Are you sure?"
End If
End Sub

Private Sub lblTriforceDecline_Click()
    ResetTriforcePicInfo True
    'play sound
    PlaySound Sound_ButtonCancel
End Sub

Private Sub lblTriforceCourage_Click()
    ResetTriforcePicInfo False
    SelectedTriforce = TriforceType.TRIFORCE_COURAGE
    'play sound
    PlaySound Sound_ButtonAccept
End Sub

Private Sub lblTriforcePower_Click()
    ResetTriforcePicInfo False
    SelectedTriforce = TriforceType.TRIFORCE_POWER
    'play sound
    PlaySound Sound_ButtonAccept
    End Sub

Private Sub lblTriforceWisdom_Click()
    ResetTriforcePicInfo False
    SelectedTriforce = TriforceType.TRIFORCE_WISDOM
    'play sound
    PlaySound Sound_ButtonAccept
End Sub

