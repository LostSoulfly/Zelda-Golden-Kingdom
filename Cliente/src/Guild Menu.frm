VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   3210
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGuild 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   120
      ScaleHeight     =   302
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2910
      Begin VB.ListBox lstGuildMembers 
         Appearance      =   0  'Flat
         BackColor       =   &H006A8092&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2190
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2445
      End
      Begin VB.Frame frmGuildC 
         BackColor       =   &H006A8092&
         Caption         =   "Guild Creation"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   2445
         Begin VB.TextBox txtGuildC 
            Height          =   285
            Left            =   240
            TabIndex        =   2
            Text            =   "Enter guild name..."
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblGuildCAccept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Accept"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   645
         End
         Begin VB.Label lblGuildCCancel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Decline"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1680
            TabIndex        =   3
            Top             =   840
            Width           =   720
         End
      End
      Begin VB.Label lblGuildLeave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Guild"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   1035
      End
      Begin VB.Label lblGuildKick 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expulsar del Clan"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   1515
      End
      Begin VB.Label lblGuildInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invite to Guild"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label lblGuild 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guild:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1095
         TabIndex        =   10
         Top             =   45
         Width           =   645
      End
      Begin VB.Label lblGuildC 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create Guild"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   780
         TabIndex        =   9
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label lblGuildDisband 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deshacer Clan"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   3600
         Width           =   1230
      End
      Begin VB.Label lblGuildYes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1800
         TabIndex        =   7
         Top             =   3600
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblGuildNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2280
         TabIndex        =   6
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
