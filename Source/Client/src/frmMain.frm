VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   681
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ComboBox cmbChat 
      Height          =   300
      ItemData        =   "frmMain.frx":0A4E
      Left            =   5760
      List            =   "frmMain.frx":0A5B
      Style           =   2  'Dropdown List
      TabIndex        =   230
      Top             =   9240
      Width           =   1455
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8160
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   63
      Top             =   4920
      Visible         =   0   'False
      Width           =   3600
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   1455
         TabIndex        =   210
         Top             =   3240
         Width           =   1455
         Begin VB.OptionButton optMiniMapOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   212
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optMiniMapOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   211
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1455
         TabIndex        =   206
         Top             =   2760
         Width           =   1455
         Begin VB.OptionButton optSafeOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   208
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optSafeOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   600
            TabIndex        =   207
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1455
         TabIndex        =   203
         Top             =   2280
         Width           =   1455
         Begin VB.OptionButton optLvlOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   205
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optLvlOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   600
            TabIndex        =   204
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1455
         TabIndex        =   200
         Top             =   1800
         Width           =   1455
         Begin VB.OptionButton optNOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   600
            TabIndex        =   202
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optNOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   201
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.CommandButton CmdMap 
         BackColor       =   &H8000000E&
         Caption         =   "World Map"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton cmdCode 
         BackColor       =   &H8000000E&
         Caption         =   "Redeem"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.HScrollBar scrlVolume 
         Height          =   255
         LargeChange     =   10
         Left            =   1800
         Max             =   100
         TabIndex        =   193
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmdVideoOptions 
         BackColor       =   &H8000000E&
         Caption         =   "Video"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   192
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdChatDisplay 
         BackColor       =   &H8000000E&
         Caption         =   "Chat"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   191
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdOnline 
         BackColor       =   &H8000000E&
         Caption         =   "Online"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdGuildOpen 
         BackColor       =   &H8000000E&
         Caption         =   "Guild"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton HelpBoardButton 
         BackColor       =   &H8000000E&
         Caption         =   "Mini-Manual"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   3600
         Width           =   3375
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1455
         TabIndex        =   69
         Top             =   1320
         Width           =   1455
         Begin VB.OptionButton optSOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   71
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optSOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   600
            TabIndex        =   70
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1455
         TabIndex        =   66
         Top             =   840
         Width           =   1455
         Begin VB.OptionButton optMOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   600
            TabIndex        =   68
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optMOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Label lblPing 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ping:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1800
         TabIndex        =   214
         Top             =   3330
         Width           =   1650
      End
      Begin VB.Label MiniMap 
         BackStyle       =   0  'Transparent
         Caption         =   "Mini Map"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   213
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblVolume 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   194
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblSafeMode 
         BackStyle       =   0  'Transparent
         Caption         =   "Safe Mode"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   188
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   121
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   116
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   65
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   64
         Top             =   600
         Width           =   675
      End
   End
   Begin VB.PictureBox picPets 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8760
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   150
      Top             =   3960
      Visible         =   0   'False
      Width           =   2910
      Begin VB.HScrollBar scrlPet 
         Height          =   255
         Left            =   240
         Max             =   6
         Min             =   1
         TabIndex        =   151
         Top             =   600
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pet Mode:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   255
         Left            =   360
         TabIndex        =   229
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblPetPassiveActive 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Passive"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1560
         TabIndex        =   228
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblPetMPNum 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0/0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   810
         TabIndex        =   187
         Top             =   2760
         Width           =   1875
      End
      Begin VB.Label lblPetMP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MP:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   255
         Left            =   240
         TabIndex        =   186
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblPetStats 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pet Stats"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   165
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label lblPetDisband 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Disband"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   163
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblPetFollow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Follow"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   255
         Left            =   360
         TabIndex        =   162
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblPetDeambulate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Wander"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   161
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblPetAttack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Attack"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   255
         Left            =   360
         TabIndex        =   160
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblPetTame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Let's Tame The Target!"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   240
         TabIndex        =   159
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lblPetLvlNum 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2040
         TabIndex        =   158
         Top             =   3000
         Width           =   585
      End
      Begin VB.Label lbPetlLvl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   255
         Left            =   240
         TabIndex        =   157
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblPetExp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   255
         Left            =   240
         TabIndex        =   156
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label lblPetExpNum 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   720
         TabIndex        =   155
         Top             =   3240
         Width           =   1965
      End
      Begin VB.Label lblAcceptPet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Pet"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   154
         Top             =   900
         Width           =   2895
      End
      Begin VB.Label lblChoosePet 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   152
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.PictureBox picPetStats 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8760
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   166
      Top             =   4080
      Visible         =   0   'False
      Width           =   2910
      Begin VB.Frame frmPetExp 
         BackColor       =   &H80000007&
         Caption         =   "Exp"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   182
         Top             =   840
         Width           =   2655
         Begin VB.HScrollBar scrlPetExp 
            Height          =   255
            Left            =   120
            Max             =   5
            Min             =   1
            TabIndex        =   183
            Top             =   720
            Value           =   3
            Width           =   2415
         End
         Begin VB.Label lblPetExpText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exp: 50%"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   120
            TabIndex        =   184
            Top             =   360
            Width           =   1290
         End
      End
      Begin VB.Label lblClosepicPets 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   185
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label lblPetForsakeNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2490
         TabIndex        =   181
         Top             =   3360
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblPetForsakeYes 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   180
         Top             =   3360
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblPetForsake 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abandon Pet"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   600
         TabIndex        =   179
         Top             =   3360
         Width           =   1710
      End
      Begin VB.Label lblPetName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   600
         TabIndex        =   178
         Top             =   480
         Width           =   1680
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   6
         Left            =   1230
         TabIndex        =   177
         Top             =   2520
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   9
         Left            =   2400
         TabIndex        =   176
         Top             =   2550
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   7
         Left            =   1230
         TabIndex        =   175
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   10
         Left            =   2400
         TabIndex        =   174
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   8
         Left            =   1230
         TabIndex        =   173
         Top             =   2970
         Width           =   120
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   6
         Left            =   1380
         TabIndex        =   172
         Top             =   2505
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   9
         Left            =   2550
         TabIndex        =   171
         Top             =   2505
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   7
         Left            =   1380
         TabIndex        =   170
         Top             =   2730
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   10
         Left            =   2550
         TabIndex        =   169
         Top             =   2730
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   8
         Left            =   1380
         TabIndex        =   168
         Top             =   2955
         Width           =   105
      End
      Begin VB.Label lblPetPoints 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2400
         TabIndex        =   167
         Top             =   2970
         Width           =   120
      End
   End
   Begin VB.PictureBox picParty 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8280
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   97
      Top             =   4920
      Visible         =   0   'False
      Width           =   2910
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   4
         Left            =   90
         Top             =   3075
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   4
         Left            =   90
         Top             =   2940
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   3
         Left            =   90
         Top             =   2340
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   3
         Left            =   90
         Top             =   2205
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   2
         Left            =   90
         Top             =   1620
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   2
         Left            =   90
         Top             =   1485
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   1
         Left            =   90
         Top             =   870
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   1
         Left            =   90
         Top             =   735
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Label lblPartyLeave 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   103
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblPartyInvite 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   102
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   101
         Top             =   2670
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   100
         Top             =   1935
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   99
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   98
         Top             =   465
         Width           =   2415
      End
   End
   Begin VB.PictureBox picCharacter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4860
      Left            =   8640
      ScaleHeight     =   324
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   2910
      Begin VB.CommandButton cmdCounters 
         BackColor       =   &H80000009&
         Caption         =   "Counter"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   222
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdCounters 
         BackColor       =   &H80000009&
         Caption         =   "Death"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   221
         Top             =   3960
         Width           =   1095
      End
      Begin VB.PictureBox picFace 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   735
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   80
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label lblHP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   1680
         TabIndex        =   226
         Top             =   4440
         Width           =   1125
      End
      Begin VB.Label lblHP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   225
         Top             =   4440
         Width           =   465
      End
      Begin VB.Label lblInvWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weight: 100%"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   216
         Top             =   4200
         Width           =   1260
      End
      Begin VB.Label lblKillPoints 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Armada: 0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   215
         Top             =   3960
         Width           =   1380
      End
      Begin VB.Label lblPoints 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2400
         TabIndex        =   87
         Top             =   2970
         Width           =   120
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   1380
         TabIndex        =   51
         Top             =   2955
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   2550
         TabIndex        =   50
         Top             =   2730
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   1380
         TabIndex        =   49
         Top             =   2730
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   2550
         TabIndex        =   48
         Top             =   2505
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   1380
         TabIndex        =   47
         Top             =   2505
         Width           =   105
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   1230
         TabIndex        =   13
         Top             =   2970
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   2400
         TabIndex        =   12
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   1230
         TabIndex        =   11
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   2400
         TabIndex        =   10
         Top             =   2550
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   1230
         TabIndex        =   9
         Top             =   2520
         Width           =   120
      End
      Begin VB.Label lblCharName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   495
         Width           =   2640
      End
   End
   Begin VB.PictureBox picHotbar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   60
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   476
      TabIndex        =   83
      Top             =   9630
      Width           =   7140
   End
   Begin VB.TextBox txtMyChat 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   60
      MaxLength       =   200
      TabIndex        =   3
      Top             =   9300
      Width           =   5580
   End
   Begin VB.PictureBox picCurrency 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   2400
      ScaleHeight     =   1995
      ScaleWidth      =   7140
      TabIndex        =   55
      Top             =   4560
      Width           =   7140
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   57
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblCurrencyCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3150
         TabIndex        =   59
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label lblCurrencyOk 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Pull"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3300
         TabIndex        =   58
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblCurrency 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "How much do you want to throw?"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   56
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   12000
      ScaleHeight     =   583
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   2745
      Begin VB.CommandButton cmdClose 
         Caption         =   "x"
         Height          =   300
         Left            =   2400
         TabIndex        =   227
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdACustomSprite 
         Caption         =   "Sprites"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   189
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton cmdAPet 
         Caption         =   "Pet"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   164
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton cmdAAction 
         Caption         =   "Action"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   149
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMovement 
         Caption         =   "Movement"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   148
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmdAName 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   122
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdADoor 
         Caption         =   "Switch"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   115
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdAQuest 
         Caption         =   "Quest"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   114
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdSSMap 
         Caption         =   "Screenshot Map"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   8400
         Width           =   2295
      End
      Begin VB.CommandButton cmdLevel 
         Caption         =   "Level Up"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   8040
         Width           =   2295
      End
      Begin VB.CommandButton cmdAAnim 
         Caption         =   "Animation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   52
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdAAccess 
         Caption         =   "Set Access"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtAAccess 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   44
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtASprite 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   42
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton cmdARespawn 
         Caption         =   "Respawn"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   41
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdASprite 
         Caption         =   "Set Sprite"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   40
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Spawn Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   7560
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   255
         LargeChange     =   10
         Left            =   240
         Min             =   1
         TabIndex        =   38
         Top             =   7200
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   36
         Top             =   6720
         Value           =   1
         Width           =   2295
      End
      Begin VB.CommandButton cmdASpell 
         Caption         =   "Spell"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAShop 
         Caption         =   "Shop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdAResource 
         Caption         =   "Resource"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   32
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdANpc 
         Caption         =   "NPC"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMap 
         Caption         =   "Map"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdAItem 
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdADestroy 
         Caption         =   "Del Bans"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMapReport 
         Caption         =   "Map Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton cmdALoc 
         Caption         =   "Loc"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Warp To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtAMap 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "WarpMe2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Warp2Me"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kick"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtAName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000E&
         X1              =   16
         X2              =   168
         Y1              =   528
         Y2              =   528
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Access:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1440
         TabIndex        =   45
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite#:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1440
         TabIndex        =   43
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   6960
         Width           =   2295
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Item: None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   6480
         Width           =   2295
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000E&
         X1              =   16
         X2              =   168
         Y1              =   432
         Y2              =   432
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000E&
         X1              =   16
         X2              =   168
         Y1              =   376
         Y2              =   376
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Editors:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         X1              =   16
         X2              =   168
         Y1              =   192
         Y2              =   192
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Map#:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         X1              =   16
         X2              =   168
         Y1              =   136
         Y2              =   136
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Panel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   2865
      End
   End
   Begin VB.PictureBox picTrade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   240
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   74
      Top             =   1080
      Visible         =   0   'False
      Width           =   7200
      Begin VB.PictureBox picYourTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   435
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   75
         Top             =   645
         Width           =   2895
      End
      Begin VB.PictureBox picTheirTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   3855
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   76
         Top             =   645
         Width           =   2895
      End
      Begin VB.Image imgDeclineTrade 
         Height          =   435
         Left            =   3675
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Image imgAcceptTrade 
         Height          =   435
         Left            =   2475
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Label lblTradeStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   600
         TabIndex        =   79
         Top             =   5520
         Width           =   5895
      End
      Begin VB.Label lblTheirWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4920
         TabIndex        =   78
         Top             =   4650
         Width           =   1575
      End
      Begin VB.Label lblYourWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   77
         Top             =   4650
         Width           =   1575
      End
   End
   Begin VB.PictureBox picSpellDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   0
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   62
      Top             =   10200
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picSpellDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   86
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblSpellDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """Default Item Description! :D"""
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1530
         Left            =   240
         TabIndex        =   85
         Top             =   1800
         Width           =   2640
      End
      Begin VB.Label lblSpellName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   84
         Top             =   240
         Width           =   2805
      End
   End
   Begin VB.PictureBox picSpeech 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   2640
      ScaleHeight     =   2085
      ScaleWidth      =   7140
      TabIndex        =   117
      Top             =   4560
      Visible         =   0   'False
      Width           =   7140
      Begin VB.PictureBox picSpeechFace 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   300
         ScaleHeight     =   1500
         ScaleWidth      =   1500
         TabIndex        =   119
         Top             =   300
         Width           =   1500
      End
      Begin VB.PictureBox picSpeechClose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6840
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   118
         Top             =   0
         Width           =   300
      End
      Begin VB.Label lblSpeech 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1500
         Left            =   2040
         TabIndex        =   120
         Top             =   300
         Width           =   4755
      End
   End
   Begin VB.PictureBox picQuestDialogue 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   3840
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   104
      Top             =   2040
      Visible         =   0   'False
      Width           =   4905
      Begin VB.Label lblQuestExtra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extra"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   106
         Top             =   1920
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblQuestSay 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1125
         Left            =   240
         TabIndex        =   110
         Top             =   720
         Width           =   4425
      End
      Begin VB.Label lblQuestAccept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accept Quest"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   210
         Left            =   240
         TabIndex        =   109
         Top             =   1920
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lblQuestClose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Left            =   4080
         TabIndex        =   108
         Top             =   1920
         Width           =   660
      End
      Begin VB.Label lblQuestName 
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Quest"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   107
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label lblQuestSubtitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Subtitle"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   105
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.PictureBox picItemDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   3240
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   6
      Top             =   10200
      Visible         =   0   'False
      Width           =   3255
      Begin VB.PictureBox picItemDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   82
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblItemWeight 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   240
         TabIndex        =   190
         Top             =   3000
         Width           =   2640
      End
      Begin VB.Label lblItemDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1170
         Left            =   240
         TabIndex        =   81
         Top             =   1800
         Width           =   2640
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   7
         Top             =   210
         Width           =   2805
      End
   End
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
      Height          =   4290
      Left            =   9000
      ScaleHeight     =   286
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   125
      Top             =   1080
      Visible         =   0   'False
      Width           =   2910
      Begin VB.PictureBox picGuildInvitation 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         FillColor       =   &H00442501&
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
         Height          =   735
         Left            =   240
         ScaleHeight     =   705
         ScaleWidth      =   2385
         TabIndex        =   141
         Top             =   2640
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Label lblGuildDeclineInvitation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reject"
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
            Left            =   1320
            TabIndex        =   144
            Top             =   360
            Width           =   795
         End
         Begin VB.Label lblGuildAcceptInvitation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Accept"
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
            Left            =   360
            TabIndex        =   143
            Top             =   360
            Width           =   690
         End
         Begin VB.Label lblGuildInvitation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invitation to the Guild"
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
            Left            =   360
            TabIndex        =   142
            Top             =   0
            Width           =   1755
         End
      End
      Begin VB.Frame frmGuildC 
         BackColor       =   &H00004000&
         Caption         =   "Creation of the Guild"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1215
         Left            =   240
         TabIndex        =   127
         Top             =   1320
         Width           =   2445
         Begin VB.TextBox txtGuildC 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   128
            Text            =   "Write the name"
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblGuildCCancel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reject"
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
            Left            =   1440
            TabIndex        =   130
            Top             =   840
            Width           =   900
         End
         Begin VB.Label lblGuildCAccept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Create"
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
            Left            =   240
            TabIndex        =   129
            Top             =   840
            Width           =   555
         End
      End
      Begin VB.ListBox lstGuildMembers 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
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
         Height          =   2010
         Left            =   240
         TabIndex        =   126
         Top             =   600
         Width           =   2445
      End
      Begin VB.Label lblGuildTransfer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer Founder"
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
         TabIndex        =   145
         Top             =   3360
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblGuildFounder 
         BackStyle       =   0  'Transparent
         Caption         =   "Founder:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   140
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label lblGuildAdminPanel 
         BackStyle       =   0  'Transparent
         Caption         =   "Managing the Guild"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   139
         Top             =   3840
         Width           =   2415
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
         Left            =   2400
         TabIndex        =   138
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
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
         Left            =   1920
         TabIndex        =   137
         Top             =   3600
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblGuildDisband 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Undo Guild"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   136
         Top             =   3600
         Width           =   1230
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
         TabIndex        =   135
         Top             =   1200
         Width           =   1365
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
         Left            =   225
         TabIndex        =   134
         Top             =   150
         Width           =   2520
      End
      Begin VB.Label lblGuildInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invite to the Guild"
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
         TabIndex        =   133
         Top             =   2640
         Width           =   1290
      End
      Begin VB.Label lblGuildKick 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expel from the Guild"
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
         TabIndex        =   132
         Top             =   3120
         Width           =   1515
      End
      Begin VB.Label lblGuildLeave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leaving Guild"
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
         TabIndex        =   131
         Top             =   2880
         Width           =   1425
      End
   End
   Begin VB.PictureBox picDialogue 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   2640
      ScaleHeight     =   2085
      ScaleWidth      =   7140
      TabIndex        =   91
      Top             =   4560
      Visible         =   0   'False
      Width           =   7140
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   3405
         TabIndex        =   96
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label lblDialogue_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Robin has requested a trade. Would you like to accept?"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   95
         Top             =   720
         Width           =   6615
      End
      Begin VB.Label lblDialogue_Title 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Request"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   3450
         TabIndex        =   93
         Top             =   1320
         Width           =   195
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   3405
         TabIndex        =   92
         Top             =   1560
         Width           =   285
      End
   End
   Begin VB.OptionButton ChatOpts 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Guild"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   2640
      TabIndex        =   220
      Top             =   6660
      Width           =   750
   End
   Begin VB.OptionButton ChatOpts 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Team"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   1710
      TabIndex        =   219
      Top             =   6660
      Width           =   930
   End
   Begin VB.OptionButton ChatOpts 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Global"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   810
      TabIndex        =   218
      Top             =   6660
      Value           =   -1  'True
      Width           =   900
   End
   Begin VB.OptionButton ChatOpts 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Map"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   0
      TabIndex        =   217
      Top             =   6660
      Width           =   825
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   2250
      Left            =   0
      TabIndex        =   0
      Top             =   6870
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   3969
      _Version        =   393217
      BackColor       =   -2147483641
      BorderStyle     =   0
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":0A7A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picTempInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   12120
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   4
      Top             =   8940
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   13320
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   88
      Top             =   8940
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picSSMap 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   90
      Top             =   9960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picInventory 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8880
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.PictureBox picQuestLog 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   9000
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   111
      Top             =   4920
      Visible         =   0   'False
      Width           =   2910
      Begin VB.TextBox txtQuestTaskLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   270
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   113
         Top             =   2625
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ListBox lstQuestLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   2550
         Left            =   120
         TabIndex        =   112
         Top             =   360
         Width           =   2655
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   1
         Left            =   390
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   6
         Left            =   2190
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   5
         Left            =   1830
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   4
         Left            =   1470
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   3
         Left            =   1110
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   2
         Left            =   750
         Top             =   3480
         Width           =   315
      End
   End
   Begin VB.PictureBox picSpells 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   9000
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   54
      Top             =   4800
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.PictureBox picShop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   360
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   60
      Top             =   1080
      Visible         =   0   'False
      Width           =   4125
      Begin VB.PictureBox picShopItems 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   615
         ScaleHeight     =   211
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   61
         Top             =   630
         Width           =   2895
      End
      Begin VB.Image imgLeaveShop 
         Height          =   435
         Left            =   2715
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Image imgShopSell 
         Height          =   435
         Left            =   1545
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Image imgShopBuy 
         Height          =   435
         Left            =   375
         Top             =   4350
         Width           =   1035
      End
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   12120
      Top             =   9600
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.PictureBox picTempBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   12720
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   73
      Top             =   8940
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   -720
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   72
      Top             =   960
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00181C21&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9120
      Left            =   0
      ScaleHeight     =   608
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
      Begin VB.PictureBox PicBars 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Height          =   150
         Index           =   1
         Left            =   150
         ScaleHeight     =   120
         ScaleWidth      =   4140
         TabIndex        =   223
         Top             =   600
         Width           =   4170
         Begin VB.Label lblMP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100/100"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   3480
            TabIndex        =   224
            Top             =   -45
            Width           =   615
         End
         Begin VB.Image imgMPBar 
            Height          =   135
            Left            =   0
            Top             =   0
            Width           =   4200
         End
      End
      Begin VB.PictureBox PicBars 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   330
         Index           =   2
         Left            =   7440
         ScaleHeight     =   300
         ScaleWidth      =   4350
         TabIndex        =   197
         Top             =   120
         Width           =   4380
         Begin VB.Label lblEXP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100/100"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   2400
            TabIndex        =   198
            Top             =   0
            Width           =   1845
         End
         Begin VB.Image imgEXPBar 
            Height          =   390
            Left            =   0
            Top             =   0
            Width           =   4350
         End
      End
      Begin VB.PictureBox PicBars 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   450
         Index           =   0
         Left            =   75
         ScaleHeight     =   420
         ScaleWidth      =   4290
         TabIndex        =   196
         Top             =   75
         Visible         =   0   'False
         Width           =   4320
         Begin VB.Image imgHPBar 
            Height          =   450
            Left            =   0
            Top             =   0
            Width           =   4320
         End
      End
      Begin MSWinsockLib.Winsock Socket 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.PictureBox picLoad 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9120
      Left            =   0
      ScaleHeight     =   608
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   146
      Top             =   -1080
      Visible         =   0   'False
      Width           =   12000
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "...Loading Map..."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   4440
         TabIndex        =   147
         Top             =   4200
         Width           =   3255
      End
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   8
      Left            =   10755
      Top             =   9720
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   7
      Left            =   9600
      Top             =   9720
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   6
      Left            =   8460
      Top             =   9720
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   5
      Left            =   7320
      Top             =   9720
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   4
      Left            =   10755
      Top             =   9240
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   3
      Left            =   9600
      Top             =   9240
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   2
      Left            =   8460
      Top             =   9240
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   1
      Left            =   7320
      Top             =   9240
      Width           =   1035
   End
   Begin VB.Label lblGold 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 Rupees"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   900
      TabIndex        =   209
      Top             =   11250
      Width           =   1905
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ************
' ** Events **
' ************
Dim Dragging As Boolean  ' just a flag to know if we are clicking the image
Private prevX As Single, prevY As Single
Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long

Private Sub ClanChat_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Options.Chat = 3
    ' save to config.ini
    'SaveOptions
    
    'ClearChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClanChat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbChat_Click()
Select Case cmbChat.ListIndex
    Case 2
        Options.ChatToScreen = 0
        frmMain.txtChat.Visible = False
        ChatOpts(0).Top = 1595
        ChatOpts(1).Top = 1595
        ChatOpts(2).Top = 1595
        ChatOpts(3).Top = 1595
    Case 1
        Options.ChatToScreen = 1
        txtChat.Visible = True
        frmMain.txtChat.Visible = True
        ChatOpts(0).Top = 444
        ChatOpts(1).Top = 444
        ChatOpts(2).Top = 444
        ChatOpts(3).Top = 444
    Case 0
        Options.ChatToScreen = 2
        txtChat.Visible = False
        frmMain.txtChat.Visible = False
        ChatOpts(0).Top = 595
        ChatOpts(1).Top = 595
        ChatOpts(2).Top = 595
        ChatOpts(3).Top = 595
    End Select
        'play sound
        PlaySound Sound_ButtonChatBox
        
End Sub

Private Sub cmdAAction_Click()
If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
    Exit Sub
End If

SendRequestEditActions

End Sub

Private Sub cmdACustomSprite_Click()
If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
    Exit Sub
End If

SendRequestEditCustomSprites

End Sub

Private Sub cmdAAnim_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditAnimation
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAnim_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdADoor_Click()
' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
           
                Exit Sub
        End If
   
        SendRequestEditdoors
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "cmdADoor_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Private Sub cmdAKick_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    SendKick Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAKick_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdaMovement_Click()

If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
    Exit Sub
End If

SendRequestEditMovements

End Sub

Private Sub cmdAName_Click()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then

Exit Sub
End If

If Len(Trim$(txtAName.text)) < 2 Then
Exit Sub
End If

If IsNumeric(Trim$(txtAName.text)) Or IsNumeric(Trim$(txtAAccess.text)) Then
Exit Sub
End If

SendSetName Trim$(txtAName.text), (Trim$(txtAAccess.text))

' Error handler
Exit Sub
errorhandler:
HandleError "cmdAName_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub cmdAPet_Click()
If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
    Exit Sub
End If

SendRequestEditPets

End Sub

Private Sub cmdAQuest_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If
    
    SendRequestEditQuest
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdChatDisplay_Click()
    ChatOptionsInit
    frmChatDisplay.Show
End Sub

Private Sub cmdClose_Click()
picAdmin.Visible = False

frmMain.Width = 12090
End Sub

Private Sub cmdCode_Click()
    frmCode.Show
End Sub

Private Sub cmdCounters_Click(index As Integer)
'Kill Counter
Dim totalkills As Long
Dim totaldeaths As Long
Dim combatdeaths As Long
Dim alldeaths As Long

    combatdeaths = Player(MyIndex).Dead + Player(MyIndex).NpcDead
    alldeaths = combatdeaths + Player(MyIndex).EnviroDead

    Select Case index
        Case 0
            Call AddText("-Kill Counters-", DarkGrey, True)
            Call AddText("Killed by players: " + str(Player(MyIndex).Kill), White)
            Call AddText("Killed by creatures: " + str(Player(MyIndex).NpcKill), White)
            Call AddText("Total deaths: " + str(totalkills), White)
        Case 1
            Call AddText("-Death Counters", DarkGrey, True)
            Call AddText("Killed by Players: " + str(Player(MyIndex).Dead), White)
            Call AddText("Killed by Creatures: " + str(Player(MyIndex).NpcDead), White)
            Call AddText("Death Total by Combat: " + str(combatdeaths), White)
            Call AddText("Death by Accident: " + str(Player(MyIndex).EnviroDead), White)
            Call AddText("Death Total: " + str(alldeaths), White)
    End Select
    'play sound
    PlaySound Sound_ButtonClick2
End Sub

Private Sub cmdLevel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        Exit Sub
    End If

    SendRequestLevelUp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub CmdMap_Click()
        If frmMap.Visible = False Then
            frmMap.Visible = True
            ClearPics
            'play sound
            PlaySound Sound_ButtonClick
        Else
            frmMap.Visible = False
            'play sound
            PlaySound Sound_ButtonClick2
        End If
End Sub

Private Sub cmdOnline_Click()
 ClearPics
'frmMain.picPets.Visible = True
SendWhosOnline
MapEditorCancel
' play sound
PlaySound Sound_ButtonClick
picOptions.Visible = True
End Sub

Private Sub cmdSSMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    ' render the map temp
    ScreenshotMap
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdVideoOptions_Click()
    frmVideo.Show
End Sub

Private Sub Form_Resize()
 'picScreen.Width = frmMain.Width / 15.5 ' Width in tiles * 32
 '   picScreen.Height = frmMain.Height / 18 ' Height in tiles * 32
 '   ReInitSurfaces = True
End Sub

Private Sub HelpBoardButton_Click()
If frmHelpBoard.Visible = False Then
frmHelpBoard.Visible = True
Else
frmHelpBoard.Visible = False
End If
End Sub

Private Sub Form_Load()
Dim X As Long, Y As Long, xwidth As Long, yheight As Long
Dim sRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' move GUI
    'picAdmin.Left = 600
    picScreen.Width = 800 ' Width in tiles * 32
    picScreen.Height = 608 ' Height in tiles * 32
    
    Dim e As Control
    
    For Each e In Me.Controls
        If (TypeOf e Is Label) Then
            ''e.caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is CheckBox) Then
            ''e.caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is OptionButton) Then
            ''e.caption = GetTranslation(e.Caption)
        End If
    Next
        
        SetFocusOnChat
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Cancel = True
    logoutGame
    frmTutorial.Visible = False
    frmTriForce.Visible = False
    frmMap.Visible = False
    txtMyChat.Locked = False
    frmMain.txtMyChat.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblAcceptPet_Click()
    Call SpawnPet(MyIndex)
End Sub

Private Sub lblClosepicPets_Click()
frmMain.picPetStats.Visible = False
frmMain.picPets.Visible = True
End Sub

Private Sub lblGuildAcceptInvitation_Click()
Call GuildCommand(6, "")
picGuild.Visible = False
End Sub

Private Sub lblGuildDeclineInvitation_Click()
Call GuildCommand(7, "")
picGuild.Visible = False
End Sub

Private Sub cmdGuildOpen_Click()
UpdateGuildData
If Not picGuild.Visible Then
' show the window
picGuild.Visible = True
Else
picGuild.Visible = False
End If
End Sub

Private Sub lblGuildTransfer_Click()
If lstGuildMembers.ListIndex > 0 Then
Call GuildCommand(8, Trim$(GuildData.Guild_Members((lstGuildMembers.ListIndex) + 1).User_Name))
Else
If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
Call GuildCommand(8, Player(myTarget).Name)
End If
End If
End Sub
Private Sub lblInvWeight_Click()
    AddText "Weight: " & GetPlayerWeight(MyIndex) & " / " & GetPlayerMaxWeight(MyIndex), BrightGreen, True
End Sub

Private Sub lblPetAttack_Click()
Call PetAttack(MyIndex)
End Sub

Private Sub lblPetDeambulate_Click()
Call PetWander(MyIndex)
End Sub

Private Sub lblPetDisband_Click()
Call PetDisband(MyIndex)
End Sub

Private Sub lblPetFollow_Click()
Call PetFollow(MyIndex)
End Sub

Private Sub lblPetForsake_Click()
    If lblPetForsake.Caption = "Abandon Pet" Then
        lblPetForsake.Caption = "Sure?"
        lblPetForsakeYes.Visible = True
        lblPetForsakeYes.Caption = "Yes"
        lblPetForsakeNo.Visible = True
        lblPetForsakeNo.Caption = "No"
    Else
        Exit Sub
    End If
End Sub

Private Sub lblPetForsakeNo_Click()

    lblPetForsake.Caption = "Abandon Pet"
    lblPetForsakeYes.Visible = False
    lblPetForsakeNo.Visible = False

End Sub

Private Sub lblPetForsakeYes_Click()

    If Player(MyIndex).Pet(Player(MyIndex).ActualPet).NumPet > 0 Then
        Call SendPetForsake(MyIndex, Player(MyIndex).ActualPet)
    End If
    
    lblPetForsake.Caption = "Abandon Pet"
    lblPetForsakeYes.Visible = False
    lblPetForsakeNo.Visible = False
    
End Sub

Private Sub lblPetPassiveActive_Click()
With lblPetPassiveActive
'cycle through them
    Select Case Player(MyIndex).PetState
        Case 0 'Passive
            .Caption = "Assist"
            Player(MyIndex).PetState = PetState.Assist
        Case 1 'Assist
            .Caption = "Defensive"
            Player(MyIndex).PetState = PetState.Defensive
        Case 2 'Defensive
            .Caption = "Passive"
            Player(MyIndex).PetState = PetState.Passive
    End Select
End With

SendPetState MyIndex, Player(MyIndex).PetState

End Sub

Private Sub lblPetStats_Click()
    ClearPics
    picPetStats.Visible = True
    
End Sub

Private Sub lblPetTame_Click()
    If CheckFreePetSlots(MyIndex) > 0 Then
        Call SendRequestTame(MyIndex)
    End If
    'play sound
     PlaySound Sound_ButtonClick2
End Sub

Private Sub MapChat_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Options.Chat = 0
    ' save to config.ini
    'SaveOptions
    
    'ClearChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapChat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub GlobalChat_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Options.Chat = 1
    ' save to config.ini
    'SaveOptions
    
    'ClearChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GlobalChat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMiniMapOff_Click()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

Options.MiniMap = 0
'SaveOptions

'play sound
PlaySound Sound_ButtonMiniMapOff

' Error handler
Exit Sub
errorhandler:
HandleError "optMiniMapOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub optMiniMapOn_Click()
 ' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

Options.MiniMap = 1
'SaveOptions

'play sound
PlaySound Sound_ButtonMiniMapOn

' Error handler
Exit Sub
errorhandler:
HandleError "optMiniMapOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub optSafeOff_Click()
Dim i As Long
    Options.SafeMode = NO
    'SaveOptions
    SendSafeMode MyIndex, Options.SafeMode
    AddText "Be careful, Security Disabled, now you will be able to kill other civilian users.", BrightRed, True
End Sub
Private Sub optSafeOn_Click()
Dim i As Long
    Options.SafeMode = YES
    'SaveOptions
    SendSafeMode MyIndex, Options.SafeMode
    'AddText "Security Activated, now you will not attack civilians by accident.", BrightGreen, True
End Sub

Private Sub PartyChat_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Options.Chat = 2
    ' save to config.ini
    'SaveOptions
    
    'ClearChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PartyChat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub imgAcceptTrade_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    AcceptTrade
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgAcceptTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_Click(index As Integer)
Dim Buffer As clsBuffer
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    
     
    Select Case index
        Case 1
            If Not picInventory.Visible Then
                'show the window
                ClearPics
                picInventory.Visible = True
                
                BltInventory
                'play sound
                PlaySound Sound_ButtonClick
            Else
                'play sound
                PlaySound Sound_ButtonClick2
                picInventory.Visible = False
            End If
        Case 2
            If Not picSpells.Visible Then
                'send packet
                Set Buffer = New clsBuffer
                Buffer.WriteLong CSpells
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                'show the window
                ClearPics
                picSpells.Visible = True
                'play sound
                PlaySound Sound_ButtonClick
            Else
                'play sound
                PlaySound Sound_ButtonClick2
                picSpells.Visible = False
            End If
        Case 3
            If Not picCharacter.Visible Then
                ' send packet
                SendRequestPlayerData
                ' show the window
                ClearPics
                picCharacter.Visible = True
                ' play sound
                PlaySound Sound_ButtonClick
                ' Render
                BltEquipment
                BltFace
            Else
                ' play sound
                PlaySound Sound_ButtonClick2
                picCharacter.Visible = False
            End If
        Case 4
            If Not picOptions.Visible Then
                ' show the window
                ClearPics
                picOptions.Visible = True
                '/Alatar v1.2
                ' play sound
                PlaySound Sound_ButtonClick
            Else
                ' play sound
                PlaySound Sound_ButtonClick2
                picOptions.Visible = False
            End If
        Case 5
            If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                SendTradeRequest
                PlaySound Sound_ButtonClick
            Else
                AddText "Invalid target.", BrightRed, True
            End If
                'play sound
                PlaySound Sound_ButtonClick
        Case 6
            ' show the window
            If picParty.Visible = False Then
                ClearPics
                picParty.Visible = True
                'play sound
                PlaySound Sound_ButtonClick
            Else
                picParty.Visible = False
                'play sound
                PlaySound Sound_ButtonClick2
            End If
            
        Case 7 'QuestLog
            If picQuestLog.Visible = False Then
                ClearPics
                picQuestLog.Visible = True
                UpdateQuestLog
                PlaySound Sound_ButtonClick
            Else
                picQuestLog.Visible = False
                PlaySound Sound_ButtonClick2
            End If
            
        Case 8 'check pets
            
            If frmMain.picPets.Visible = False Then
                frmMain.picPets.Visible = True
                ClearPics
            'play sound
            PlaySound Sound_ButtonClick
            frmMain.picPets.Visible = True
            Else
                'play sound
                PlaySound Sound_ButtonClick2
                picPets.Visible = False
            End If

    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Main index
    
    ' change the button we're hovering on
    If Not MainButton(index).State = 2 Then ' make sure we're not clicking
        changeButtonState_Main index, 1 ' hover
    End If
    
    ' play sound
    If Not LastButtonSound_Main = index Then
        PlaySound Sound_ButtonHover
        LastButtonSound_Main = index
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
        
    ' reset all buttons
    resetButtons_Main -1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Main index
    
    ' change the button we're hovering on
    changeButtonState_Main index, 2 ' clicked
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblCurrencyCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picCurrency.Visible = False
    txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    'play sound
    PlaySound Sound_ButtonCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCurrencyCancel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgDeclineTrade_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DeclineTrade
    frmMain.picTrade.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgDeclineTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgLeaveShop_Click()
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CCloseShop
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
    
    picShop.Visible = False
    InShop = 0
    ShopAction = 0
        
    'play sound
    PlaySound Sound_ButtonClose
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgLeaveShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblCurrencyOk_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsNumeric(txtCurrency.text) Then
        If CurrencyMenu = 3 Then
            If Val(txtCurrency.text) > GetBankItemValue(MyIndex, tmpCurrencyItem) Then txtCurrency.text = GetBankItemValue(MyIndex, tmpCurrencyItem)
        ElseIf Val(txtCurrency.text) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then txtCurrency.text = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
        End If
        Select Case CurrencyMenu
            Case 1 ' drop item
                SendDropItem tmpCurrencyItem, Val(txtCurrency.text)
            Case 2 ' deposit item
                DepositItem tmpCurrencyItem, Val(txtCurrency.text)
            Case 3 ' withdraw item
                WithdrawItem tmpCurrencyItem, Val(txtCurrency.text)
            Case 4 ' offer trade item
                TradeItem tmpCurrencyItem, Val(txtCurrency.text)
        End Select
    Else
        AddText "Enter a valid amount!", BrightRed, True
        Exit Sub
    End If
    
    picCurrency.Visible = False
    tmpCurrencyItem = 0
    txtCurrency.text = vbNullString
    CurrencyMenu = 0 ' clear
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCurrencyOk_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgShopBuy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ShopAction = 1 Then
    ShopAction = 0
    'play sound
    PlaySound Sound_ButtonEnding
    AddText "Stop shopping", White, True
    Else
    ShopAction = 1 ' buying an item
    'play sound
    PlaySound Sound_ButtonInitiating
    AddText "Purchase initiated. Choose the objects to buy", White, True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgShopBuy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgShopSell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ShopAction = 2 Then
    ShopAction = 0
    'play sound
    PlaySound Sound_ButtonEnding
    AddText "Sale stop", White, True
    Else
    ShopAction = 2 ' selling an item
    'play sound
    PlaySound Sound_ButtonInitiating
    AddText "Sale initiated. Choose the objects to sell", White, True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgShopSell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblDialogue_Button_Click(index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' call the handler
    dialogueHandler index
    
    picDialogue.Visible = False
    dialogueIndex = 0
    
    'play sound
    PlaySound Sound_ButtonAccept
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblDialogue_Button_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub lblPartyInvite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
        SendPartyRequest
        AddText "A request for equipment has been sent.", BrightRed, True
    Else
        AddText "Invalid target.", BrightRed, True
    End If
    
    'play sound
    PlaySound Sound_ButtonParty
    ' Error handler
    
    Exit Sub
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblPartyLeave_Click()
        ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Party.Leader > 0 Then
        SendPartyLeave
    Else
        AddText "He's not on a team.", BrightRed, True
    End If
    
    'play sound
    PlaySound Sound_ButtonClose
    ' Error handler
    
    picParty.Visible = False
    
    Exit Sub
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblTrainStat_Click(index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case index
    Case Is <= 5
        If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
        SendTrainStat index
    Case Is <= 10
        If GetPlayerPetPOINTS(MyIndex) = 0 Then Exit Sub
        Call SendTrainPetStat(index - 5)
    End Select
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblTrainStat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub optMOff_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Music = 0
    ' stop music playing
    StopMidi
    ' save to config.ini
    'SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub optMOn_Click()
Dim MusicFile As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Music = 1
    ' start music playing
    MusicFile = Trim$(map.Music)
    If Not MusicFile = "None." Then
        PlayMidi MusicFile
    Else
        StopMidi
    End If
    ' save to config.ini
    'SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optNOff_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Names = 0
    ' save to config.ini
    'SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optNOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optNOn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Names = 1
    ' save to config.ini
    'SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optNOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSOff_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Sound = 0
    ' save to config.ini
    'SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSOn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Sound = 1
    ' save to config.ini
    'SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picGuild_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    prevX = X
    prevY = Y
End Sub
Private Sub picGuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Dragging Then
        picGuild.Move picGuild.Left - (prevX - X), picGuild.Top - (prevY - Y)
    End If
End Sub
Private Sub picGuild_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub picHotbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SlotNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SlotNum = IsHotbarSlot(X, Y)

    If Button = 1 Then
        If SlotNum <> 0 Then
            SendHotbarUse SlotNum
        End If
    ElseIf Button = 2 Then
        If SlotNum <> 0 Then
            SendHotbarChange 0, 0, SlotNum
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picHotbar_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picHotbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SlotNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SlotNum = IsHotbarSlot(X, Y)

    If SlotNum <> 0 Then
        If Hotbar(SlotNum).sType = 1 Then ' item
            X = X + picHotbar.Left + 1
            Y = Y + picHotbar.Top - picItemDesc.Height - 1
            UpdateDescWindow Hotbar(SlotNum).Slot, X, Y
            LastItemDesc = Hotbar(SlotNum).Slot ' set it so you don't re-set values
            Exit Sub
        ElseIf Hotbar(SlotNum).sType = 2 Then ' spell
            X = X + picHotbar.Left + 1
            Y = Y + picHotbar.Top - picSpellDesc.Height - 1
            UpdateSpellWindow Hotbar(SlotNum).Slot, X, Y
            LastSpellDesc = Hotbar(SlotNum).Slot  ' set it so you don't re-set values
            Exit Sub
        End If
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' no spell was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picHotbar_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    prevX = X
    prevY = Y
End Sub
Private Sub picOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging Then
    picOptions.Move picOptions.Left - (prevX - X), picOptions.Top - (prevY - Y)
    End If
End Sub
Private Sub picOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub
Private Sub picParty_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    prevX = X
    prevY = Y
End Sub
Private Sub picParty_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging Then
    picParty.Move picParty.Left - (prevX - X), picParty.Top - (prevY - Y)
    End If
End Sub
Private Sub picParty_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub picPets_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    prevX = X
    prevY = Y
End Sub


Private Sub picPets_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging Then
    picPets.Move picPets.Left - (prevX - X), picPets.Top - (prevY - Y)
    End If
End Sub

Private Sub picPets_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub
Private Sub picPetStats_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    prevX = X
    prevY = Y
End Sub
Private Sub picPetStats_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging Then
    picPetStats.Move picPetStats.Left - (prevX - X), picPetStats.Top - (prevY - Y)
    End If
End Sub
Private Sub picPetStats_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub
Private Sub picQuestDialogue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging Then
    picQuestDialogue.Move picQuestDialogue.Left - (prevX - X), picQuestDialogue.Top - (prevY - Y)
    End If
End Sub

Private Sub picQuestDialogue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    prevX = X
    prevY = Y
End Sub
Private Sub picQuestDialogue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub picQuestLog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    prevX = X
    prevY = Y
End Sub
Private Sub picQuestLog_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging Then
    picQuestLog.Move picQuestLog.Left - (prevX - X), picQuestLog.Top - (prevY - Y)
    End If
End Sub
Private Sub picQuestLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sRECT As DxVBLib.RECT
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If InMapEditor Then
        Call MapEditorMouseDown(Button, X, Y, False)
    Else
        ' left click
        If Button = vbLeftButton Then
            ' targetting
            Call PlayerSearch(CurX, CurY)
        ' right click
        ElseIf Button = vbRightButton Then
            If ShiftDown Then
                ' admin warp if we're pressing shift and right clicking
                If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
            End If
        End If
    End If

    'Call SetFocusOnChat
    
    If Options.WASD = 1 Then
        ChatFocus = False
    End If
    
    Call CheckCustomSpritePosition(X, Y)
    
    'disable main pictures
    frmMain.picOptions.Visible = False
    'frmMain.picInventory.Visible = False
    'frmMain.picSpells.Visible = False
    'frmMain.picCharacter.Visible = False
    'frmMain.picParty.Visible = False
    'frmMain.picQuestLog.Visible = False
    'frmMain.picPets.Visible = False
    'frmMain.picPetStats.Visible = False
    'frmMain.picGuild.Visible = False
    'frmMap.Visible = False
        
    picScreen.Width = 800 ' Width in tiles * 32
    picScreen.Height = 608 ' Height in tiles * 32
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picScreen_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    CurX = TileView.Left + ((X + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((Y + Camera.Top) \ PIC_Y)

    If InMapEditor Then
        frmEditor_Map.shpLoc.Visible = False

        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button, X, Y)
        End If
    End If
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picScreen_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsShopItem = 0

    For i = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(i).Item > 0 And Shop(InShop).TradeItem(i).Item <= MAX_ITEMS Then
            With tempRec
                .Top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                .Bottom = .Top + PIC_Y
                .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsShopItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsShopItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub picShop_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    If Dragging Then
    picShop.Move picShop.Left - (prevX - X), picShop.Top - (prevY - Y)
    End If
    
    ' reset all buttons
    resetButtons_Main
End Sub
Private Sub picShop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    prevX = X
    prevY = Y
End Sub
Private Sub picShop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub picShopItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim shopItem As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    shopItem = IsShopItem(X, Y)
    
    If shopItem > 0 Then
        Select Case ShopAction
            Case 0 ' no action, give cost
                With Shop(InShop).TradeItem(shopItem)
                    AddText "This can be purchased with " & .CostValue & " " & Trim$(GetShopPriceName(InShop, shopItem)) & ".", White
                End With
            Case 1 ' buy item
                ' buy item code
                BuyItem shopItem
        End Select
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShopItems_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub picShopItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim shopslot As Long
Dim x2 As Long, y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    shopslot = IsShopItem(X, Y)

    If shopslot <> 0 Then
        x2 = X + picShop.Left + picShopItems.Left + 1
        y2 = Y + picShop.Top + picShopItems.Top + 1
        UpdateDescWindow Shop(InShop).TradeItem(shopslot).Item, x2, y2
        LastItemDesc = Shop(InShop).TradeItem(shopslot).Item
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShopItems_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpeechClose_Click()
frmMain.picSpeech.Visible = False
frmMain.picSpeechClose.Visible = False
frmMain.lblSpeech.Visible = False
frmMain.picSpeechFace.Visible = False
'play sound
PlaySound Sound_ButtonClose
End Sub

Private Sub picSpellDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picSpellDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpellDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_DblClick()
Dim spellnum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = IsPlayerSpell(SpellX, SpellY)

    If spellnum <> 0 Then
        Call CastSpell(spellnum)
        Exit Sub
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub picSpells_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim spellnum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = IsPlayerSpell(SpellX, SpellY)
    If Button = 1 Then ' left click
        If spellnum <> 0 Then
            DragSpell = spellnum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' right click
        If spellnum <> 0 Then
            Dialogue "Forget Spell", "Are you sure you want to forget the skill " & Trim$(Spell(PlayerSpells(spellnum)).Name) & "ï¿½", DIALOGUE_TYPE_FORGET, True, spellnum
            Exit Sub
        End If
    End If
    
    Dragging = True
    prevX = X
    prevY = Y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim spellslot As Long
Dim x2 As Long, y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellX = X
    SpellY = Y
    
    spellslot = IsPlayerSpell(X, Y)
    
    If DragSpell > 0 Then
        Call BltDraggedSpell(X + picSpells.Left, Y + picSpells.Top)
    Else
        If spellslot <> 0 Then
            x2 = X + picSpells.Left - picSpellDesc.Width - 1
            y2 = Y + picSpells.Top - picSpellDesc.Height - 1
            UpdateSpellWindow PlayerSpells(spellslot), x2, y2
            LastSpellDesc = PlayerSpells(spellslot)
            Exit Sub
        End If
    End If
    
    picSpellDesc.Visible = False
    LastSpellDesc = 0
    
    If Dragging Then
        picSpells.Move picSpells.Left - (prevX - X), picSpells.Top - (prevY - Y)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DragSpell > 0 Then
        ' drag + drop
        For i = 1 To MAX_PLAYER_SPELLS
            With rec_pos
                .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    If DragSpell <> i Then
                        SendChangeSpellSlots DragSpell, i
                        Exit For
                    End If
                End If
            End If
        Next
        ' hotbar
        For i = 1 To MAX_HOTBAR
            With rec_pos
                .Top = picHotbar.Top - picSpells.Top
                .Left = picHotbar.Left - picSpells.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.Top - picSpells.Top + 32
            End With
            
            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    SendHotbarChange 2, DragSpell, i
                    DragSpell = 0
                    picTempSpell.Visible = False
                    Exit Sub
                End If
            End If
        Next
    End If

    DragSpell = 0
    picTempSpell.Visible = False
    
    Dragging = False
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    If Dragging Then
    picTrade.Move picTrade.Left - (prevX - X), picTrade.Top - (prevY - Y)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub picTrade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    prevX = X
    prevY = Y
End Sub
Private Sub picTrade_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub picYourTrade_DblClick()
Dim TradeNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(TradeX, TradeY, True)

    If TradeNum <> 0 Then
        UntradeItem TradeNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picYourTrade_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picYourTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeX = X
    TradeY = Y
    
    TradeNum = IsTradeItem(X, Y, True)
    
    If TradeNum <> 0 Then
        X = X + picTrade.Left + picYourTrade.Left + 4
        Y = Y + picTrade.Top + picYourTrade.Top + 4
        UpdateDescWindow GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).num), X, Y
        LastItemDesc = GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).num) ' set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picYourTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picTheirTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(X, Y, False)
    
    If TradeNum <> 0 Then
        X = X + picTrade.Left + picTheirTrade.Left + 4
        Y = Y + picTrade.Top + picTheirTrade.Top + 4
        UpdateDescWindow TradeTheirOffer(TradeNum).num, X, Y
        LastItemDesc = TradeTheirOffer(TradeNum).num ' set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picTheirTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAAmount.Caption = "Amount: " & scrlAAmount.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAAmount_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAItem.Caption = "Item: " & Trim$(Item(scrlAItem.value).Name) & ", " & scrlAItem.value
    If isItemStackable(scrlAItem.value) Then
        scrlAAmount.enabled = True
        Exit Sub
    End If
    scrlAAmount.enabled = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPet_Change()
    Dim i As Byte
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).petData.Owner = MyIndex Then
            If Player(MyIndex).ActualPet < 1 Or Player(MyIndex).ActualPet > MAX_PLAYER_PETS Then Exit Sub
            scrlPet.value = Player(MyIndex).ActualPet
            Exit Sub
        End If
    Next
    Call SendRequestChangeActualPet(MyIndex, scrlPet.value)
End Sub


Private Sub scrlPetExp_Change()
Dim Percent As Byte

Percent = 25 * (scrlPetExp.value - 1)

If Percent < 0 Or Percent > 100 Then Exit Sub

lblPetExpText.Caption = "Exp: " & Percent & "%"

Call SendPetPercent(MyIndex, Percent)

End Sub

Private Sub scrlVolume_Change()

    lblVolume.Caption = "Volume: " & scrlVolume.value
    DefaultVolume = scrlVolume.value
    Options.DefaultVolume = DefaultVolume

    ' save to config.ini
    'SaveOptions

End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsConnected Or frmMenu.Visible = True Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

If Options.WASD = 1 Then
    If ChatFocus = False Then
    'wasd
        If KeyAscii = 119 Or KeyAscii = 97 Or KeyAscii = 115 Or KeyAscii = 100 Then Exit Sub
    'wasd with shift (caps?)
        If KeyAscii = 87 Or KeyAscii = 65 Or KeyAscii = 83 Or KeyAscii = 68 Then Exit Sub
    End If
End If

    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        If Options.WASD = 1 Then
            ChatFocus = Not ChatFocus
            If ChatFocus = True Then frmMain.txtMyChat.SetFocus Else frmMain.picScreen.SetFocus
        End If
        KeyAscii = 0
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case KeyCode
        Case vbKeyInsert
            If Player(MyIndex).Access > 0 Then
                picAdmin.Visible = Not picAdmin.Visible
                If picAdmin.Visible = True Then frmMain.Width = 14835 Else frmMain.Width = 12090
            End If
        Case vbKeyEnd
            TakePicture
            PrintVideo
        Case vbKeyDelete
            If Options.MiniMap = 1 Then
                Options.MiniMap = 0
                'play sound
                PlaySound Sound_ButtonMiniMapOff
            Else
                Options.MiniMap = 1
                'play sound
                PlaySound Sound_ButtonMiniMapOn
            End If

    End Select
    
    ' hotbar
    For i = 1 To MAX_HOTBAR
        If KeyCode = 111 + i Then
            SendHotbarUse i
        End If
    Next
    
    ' Spinning
    'If KeyCode = vbKeyBack Then
    'Call Spin
    'End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

If Number = 10061 Then Exit Sub
If Number = 10053 Then Exit Sub

'MsgBox Number & ": " & Description
'DestroyGame

End Sub

Private Sub txtMyChat_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    MyText = txtMyChat
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMyChat_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub txtChat_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    'SetFocusOnChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtChat_GotFocus", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ***************
' ** Inventory **
' ***************
Private Sub picInventory_DblClick()
    Dim InvNum As Long
    Dim value As Long
    Dim multiplier As Double
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DragInvSlotNum = 0
    InvNum = IsInvItem(InvX, InvY)

    If InvNum <> 0 Then
    
        ' are we in a shop?
        If InShop > 0 Then
            Select Case ShopAction
                Case 0 ' nothing, give value
                    multiplier = Shop(InShop).BuyRate / 100
                    value = Item(GetPlayerInvItemNum(MyIndex, InvNum)).Price * multiplier
                    If value > 0 Then
                        AddText "You can sell this item for " & value & " Rupees.", White
                    Else
                        AddText "This vendor does not want this item.", BrightRed, True
                    End If
                Case 2 ' 2 = sell
                    SellItem InvNum
            End Select
            
            Exit Sub
        End If
        
        ' in bank?
        If InBank Then
            If isItemStackable(GetPlayerInvItemNum(MyIndex, InvNum)) Then
                CurrencyMenu = 2 ' deposit
                lblCurrency.Caption = "How much do you want to deposit?"
                tmpCurrencyItem = InvNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
                
            Call DepositItem(InvNum, 0)
            Exit Sub
        End If
        
        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For i = 1 To MAX_INV
                If TradeYourOffer(i).num = InvNum Then
                    ' is currency?
                    
                    If isItemStackable(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)) Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(i).value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If isItemStackable(GetPlayerInvItemNum(MyIndex, InvNum)) Then
                CurrencyMenu = 4 ' offer in trade
                lblCurrency.Caption = "How much do you want to trade?"
                tmpCurrencyItem = InvNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
            
            Call TradeItem(InvNum, 0)
            Exit Sub
        End If
        
        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(InvNum)
        Exit Sub
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Function IsEqItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsEqItem = 0

    For i = 1 To Equipment.Equipment_Count - 1

        If GetPlayerEquipment(MyIndex, i) > 0 And GetPlayerEquipment(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .Top = EqTop
                .Bottom = .Top + PIC_Y
                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsEqItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsEqItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsInvItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsInvItem = 0

    For i = 1 To MAX_INV

        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsInvItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsInvItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsPlayerSpell(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsPlayerSpell = 0

    For i = 1 To MAX_PLAYER_SPELLS

        If PlayerSpells(i) > 0 And PlayerSpells(i) <= MAX_SPELLS Then

            With tempRec
                .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsPlayerSpell = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsTradeItem(ByVal X As Single, ByVal Y As Single, ByVal Yours As Boolean) As Long
    Dim tempRec As RECT
    Dim i As Long
    Dim ItemNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsTradeItem = 0

    For i = 1 To MAX_INV
    
        If Yours Then
            ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        Else
            ItemNum = TradeTheirOffer(i).num
        End If

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then

            With tempRec
                .Top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsTradeItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTradeItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub picInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim InvNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InvNum = IsInvItem(X, Y)

    If Button = 1 Then
        If InvNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Or InShop Then Exit Sub
            DragInvSlotNum = InvNum
            Exit Sub
        End If

    ElseIf Button = 2 Then
        If Not InBank And Not InShop And Not InTrade > 0 Then
            If InvNum <> 0 Then
                If isItemStackable(GetPlayerInvItemNum(MyIndex, InvNum)) Then
                    If GetPlayerInvItemValue(MyIndex, InvNum) > 0 Then
                        CurrencyMenu = 1 ' drop
                        lblCurrency.Caption = "How much do you want to pull?"
                        tmpCurrencyItem = InvNum
                        txtCurrency.text = vbNullString
                        picCurrency.Visible = True
                        txtCurrency.SetFocus
                    End If
                Else
                    Call SendDropItem(InvNum, 0)
                End If
            End If
        End If
    End If
    
    Dragging = True
    prevX = X
    prevY = Y
    
    'SetFocusOnChat
    
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picInventory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim InvNum As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InvX = X
    InvY = Y

    If DragInvSlotNum > 0 Then
        If InTrade > 0 Then Exit Sub
        If InBank Or InShop Then Exit Sub
        Call BltInventoryItem(X + picInventory.Left, Y + picInventory.Top)
    Else
        InvNum = IsInvItem(X, Y)

        If InvNum <> 0 Then
            ' exit out if we're offering that item
            If InTrade Then
                For i = 1 To MAX_INV
                    If TradeYourOffer(i).num = InvNum Then
                        ' is currency?
                        
                        If isItemStackable(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)) Then
                            ' only exit out if we're offering all of it
                            If TradeYourOffer(i).value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then
                                Exit Sub
                            End If
                        Else
                            Exit Sub
                        End If
                    End If
                Next
            End If
            X = X + picInventory.Left - picItemDesc.Width - 1
            Y = Y + picInventory.Top - picItemDesc.Height - 1
            UpdateDescWindow GetPlayerInvItemNum(MyIndex, InvNum), X, Y
            LastItemDesc = GetPlayerInvItemNum(MyIndex, InvNum) ' set it so you don't re-set values
            Exit Sub
        End If
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    If Dragging Then
        picInventory.Move picInventory.Left - (prevX - X), picInventory.Top - (prevY - Y)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picInventory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Dragging = False
    
    If InTrade > 0 Then Exit Sub
    If InBank Or InShop Then Exit Sub

    If DragInvSlotNum > 0 Then
        ' drag + drop
        For i = 1 To MAX_INV
            With rec_pos
                .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then '
                    If DragInvSlotNum <> i Then
                        SendChangeInvSlots DragInvSlotNum, i
                        Exit For
                    End If
                End If
            End If
        Next
        ' hotbar
        For i = 1 To MAX_HOTBAR
            With rec_pos
                .Top = picHotbar.Top - picInventory.Top
                .Left = picHotbar.Left - picInventory.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.Top - picInventory.Top + 32
            End With
            
            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    SendHotbarChange 1, DragInvSlotNum, i
                    DragInvSlotNum = 0
                    picTempInv.Visible = False
                    BltHotbar
                    Exit Sub
                End If
            End If
        Next
    End If

    DragInvSlotNum = 0
    picTempInv.Visible = False
    
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picItemDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picItemDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picItemDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *****************
' ** Char window **
' *****************

Private Sub picCharacter_Click()
    Dim EqNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EqNum = IsEqItem(EqX, EqY)

    If EqNum <> 0 Then
        SendUnequip EqNum
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub picCharacter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    prevX = X
    prevY = Y
End Sub
Private Sub picCharacter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim EqNum As Long
    Dim x2 As Long, y2 As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EqX = X
    EqY = Y
    EqNum = IsEqItem(X, Y)

    If EqNum <> 0 Then
        y2 = Y + picCharacter.Top - frmMain.picItemDesc.Height - 1
        x2 = X + picCharacter.Left - frmMain.picItemDesc.Width - 1
        UpdateDescWindow GetPlayerEquipment(MyIndex, EqNum), x2, y2
        LastItemDesc = GetPlayerEquipment(MyIndex, EqNum) ' set it so you don't re-set values
        Exit Sub
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    If Dragging Then
        picCharacter.Move picCharacter.Left - (prevX - X), picCharacter.Top - (prevY - Y)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ****************
' ** Admin Menu **
' ****************

Private Sub cmdALoc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    BLoc = Not BLoc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdALoc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If
    
    SendRequestEditMap
    
    Options.MappingMode = 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp2Me_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp2Me_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarpMe2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarpMe2_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp_Click()
Dim N As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAMap.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.text)) Then
        Exit Sub
    End If

    N = CLng(Trim$(txtAMap.text))

    ' Check to make sure its a valid map #
    If N > 0 And N <= MAX_MAPS Then
        Call WarpTo(N)
    Else
        Call AddText("Map number not visible.", Red, True)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASprite_Click()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then

Exit Sub
End If

If Len(Trim$(txtASprite.text)) < 1 Then
Exit Sub
End If

If Not IsNumeric(Trim$(txtASprite.text)) Then
Exit Sub
End If

If Len(Trim$(txtAName.text)) > 1 Then
SendSetSprite CLng(Trim$(txtASprite.text)), txtAName.text
Else
SendSetSprite CLng(Trim$(txtASprite.text)), GetPlayerName(MyIndex)
End If

' Error handler
Exit Sub
errorhandler:
HandleError "cmdASprite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub cmdAMapReport_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    SendMapReport
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMapReport_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdARespawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    SendMapRespawn
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdARespawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdABan_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdABan_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditItem
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdANpc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditNpc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdANpc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAResource_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditResource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAResource_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditShop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditSpell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAAccess_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 2 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Or Not IsNumeric(Trim$(txtAAccess.text)) Then
        Exit Sub
    End If

    SendSetAccess Trim$(txtAName.text), CLng(Trim$(txtAAccess.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAccess_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdADestroy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendBanDestroy
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdADestroy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If
    
    SendSpawnItem scrlAItem.value, scrlAAmount.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' bank
Private Sub picBank_DblClick()
Dim bankNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DragBankSlotNum = 0

    bankNum = IsBankItem(BankX, BankY)
    If bankNum <> 0 Then
         If GetBankItemNum(bankNum) = ITEM_TYPE_NONE Then Exit Sub


             If isItemStackable(GetBankItemNum(bankNum)) Then
                CurrencyMenu = 3 ' withdraw
                lblCurrency.Caption = "How much do you want to withdraw?"
                tmpCurrencyItem = bankNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
            
         WithdrawItem bankNum, 0
         Exit Sub
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub picBank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bankNum As Long, ItemNum As Long, ItemType As Long
Dim x2 As Long, y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    BankX = X
    BankY = Y
    
    If DragBankSlotNum > 0 Then
        Call BltBankItem(X + picBank.Left, Y + picBank.Top)
    Else
        bankNum = IsBankItem(X, Y)
        
        If bankNum <> 0 Then
            
            x2 = X + picBank.Left + 1
            y2 = Y + picBank.Top + 1
            UpdateDescWindow Bank.Item(bankNum).num, x2, y2
            Exit Sub
        End If
    End If
    
    frmMain.picItemDesc.Visible = False
    LastBankDesc = 0
    
    If Dragging Then
    picBank.Move picBank.Left - (prevX - X), picBank.Top - (prevY - Y)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub picBank_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bankNum As Long
                        
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    bankNum = IsBankItem(X, Y)
    
    If bankNum <> 0 Then
        If Button = 1 Then
            DragBankSlotNum = bankNum
        End If
    Else
        Dragging = True
        prevX = X
        prevY = Y
    End If
    

    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

' TODO : Add sub to change bankslots client side first so there's no delay in switching
    If DragBankSlotNum > 0 Then
        For i = 1 To MAX_BANK
            With rec_pos
                .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    If DragBankSlotNum <> i Then
                        ChangeBankSlots DragBankSlotNum, i
                        Exit For
                    End If
                End If
            End If
        Next
    End If

    DragBankSlotNum = 0
    picTempBank.Visible = False
    Dragging = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsBankItem(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsBankItem = 0
    
    For i = 1 To MAX_BANK
        If GetBankItemNum(i) > 0 And GetBankItemNum(i) <= MAX_ITEMS Then
        
            With tempRec
                .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With
            
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    
                    IsBankItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsBankItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

'ALATAR

'QuestDialogue:

Private Sub lblQuestAccept_Click()
    PlayerHandleQuest CLng(lblQuestAccept.Tag), 1
    picQuestDialogue.Visible = False
    lblQuestAccept.Visible = False
    lblQuestAccept.Tag = vbNullString
    lblQuestSay = "-"
    RefreshQuestLog
    
    'play sound
    PlaySound Sound_ButtonAccept
    
End Sub

Private Sub lblQuestExtra_Click()
    RunQuestDialogueExtraLabel
    
    'play sound
    PlaySound Sound_ButtonCancel
    
End Sub

Private Sub lblQuestClose_Click()
    picQuestDialogue.Visible = False
    lblQuestExtra.Visible = False
    lblQuestAccept.Visible = False
    lblQuestAccept.Tag = vbNullString
    lblQuestSay = "-"
    
    'play sound
    PlaySound Sound_ButtonClose
    
End Sub

'QuestLog:
'Private Sub picQuestButton_Click()
'    'Need to be replaced with imgButton(X) and a proper image
'    UpdateQuestLog
'    picQuestLog.Visible = Not picQuestLog.Visible
'    PlaySound Sound_ButtonClick
'End Sub

Private Sub imgQuestButton_Click(index As Integer)
    If Trim$(lstQuestLog.text) = vbNullString Then Exit Sub
    LoadQuestlogBox index
    
    'play sound
    PlaySound Sound_ButtonQuest
    
End Sub
Private Sub optLvlOn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Level = 1
    ' save to config.ini
    'SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optLvlOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optLvlOff_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Level = 0
    ' save to config.ini
    'SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optLvlOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblGuildInv_Click()
If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
Call GuildCommand(2, Player(myTarget).Name)
Else
AddText "The user cannot be invited to the clan or has not been selected.", BrightRed, True
End If
End Sub

Private Sub lblGuildLeave_Click()
Call GuildCommand(3, "")
picGuild.Visible = False
End Sub

Private Sub lblGuildKick_Click()
If lstGuildMembers.ListIndex > 0 Then
Call GuildCommand(9, Trim$(GuildData.Guild_Members((lstGuildMembers.ListIndex) + 1).User_Name))
Else
If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
Call GuildCommand(9, Player(myTarget).Name)
End If
End If
End Sub

Private Sub lblGuildDisband_Click()
lblGuildDisband.Caption = "Are you sure?"
lblGuildYes.Visible = True
lblGuildNo.Visible = True
lblGuildYes.Caption = "Yes"
lblGuildNo.Caption = "No"
End Sub

Private Sub lblGuildYes_Click()
lblGuildYes.Visible = False
lblGuildNo.Visible = False
picGuild.Visible = False
lblGuildDisband.Caption = "Disband Guild"
Call GuildCommand(10, "")
End Sub

Private Sub lblGuildNo_Click()
lblGuildYes.Visible = False
lblGuildNo.Visible = False
lblGuildDisband.Caption = "Disband Guild"
End Sub

Private Sub lblGuildCAccept_Click()
Call GuildCommand(1, txtGuildC.text)
frmGuildC.Visible = False
picGuild.Visible = False
picGuildInvitation.Visible = False
End Sub

Private Sub lblGuildCCancel_Click()
lblGuildC.Visible = True
frmGuildC.Visible = False
End Sub

Private Sub lblGuildC_Click()
lblGuildC.Visible = False
frmGuildC.Visible = True
End Sub

Private Sub lblGuildFounder_Click()
If lstGuildMembers.ListIndex > 0 Then
Call GuildCommand(8, Trim$(GuildData.Guild_Members((lstGuildMembers.ListIndex) + 1).User_Name))
Else
If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
Call GuildCommand(8, Player(myTarget).Name)
End If
End If
End Sub

Private Sub lblGuildAdminPanel_Click()
Call GuildCommand(4, "")
End Sub

Private Sub ClearPics()
            picCharacter.Visible = False
            picInventory.Visible = False
            picSpells.Visible = False
            picOptions.Visible = False
            picParty.Visible = False
            picQuestLog.Visible = False
            picPets.Visible = False
            picPetStats.Visible = False
            picGuild.Visible = False
End Sub

Private Sub txtMyChat_Click()
ChatFocus = True
End Sub
