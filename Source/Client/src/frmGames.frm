VERSION 5.00
Begin VB.Form frmGames 
   Caption         =   "Games"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraGames 
      Caption         =   "Game Rooms"
      Height          =   5055
      Left            =   5880
      TabIndex        =   2
      Top             =   360
      Width           =   4455
      Begin VB.CommandButton cmdStartGame 
         Caption         =   "Start the Game"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4560
         Width           =   1695
      End
      Begin VB.ComboBox cmdGameType 
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Text            =   "Type"
         Top             =   4200
         Width           =   855
      End
      Begin VB.CommandButton cmdCreateGame 
         Caption         =   "Create a Game"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton cmdJoin 
         Caption         =   "Join it"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   3840
         Width           =   1695
      End
      Begin VB.ListBox listGames 
         Height          =   3180
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame fraTeams 
      Caption         =   "Online Teams"
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      Begin VB.CommandButton cmdBet 
         Caption         =   "Bet"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   3480
         Width           =   1695
      End
      Begin VB.ListBox listTeamIntegrants 
         Height          =   2790
         Left            =   2880
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdCreateTeam 
         Caption         =   "Creating a Team"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton cmdInviteTeam 
         Caption         =   "Invite Team"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3840
         Width           =   1695
      End
      Begin VB.ListBox listTeams 
         Height          =   3180
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
