VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMap 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "World Map"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab WorldMap 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Hyrule"
      TabPicture(0)   =   "frmMap.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MiniMapHyrule(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Termina"
      TabPicture(1)   =   "frmMap.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MiniMapTermina(1)"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox MiniMapHyrule 
         BackColor       =   &H00004000&
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
         Height          =   5520
         Index           =   0
         Left            =   240
         Picture         =   "frmMap.frx":0038
         ScaleHeight     =   5520
         ScaleWidth      =   8580
         TabIndex        =   18
         Top             =   480
         Width           =   8580
         Begin VB.OptionButton NoneHyrule 
            BackColor       =   &H00004000&
            Caption         =   "Nothing"
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
            Height          =   210
            Index           =   0
            Left            =   4320
            TabIndex        =   22
            Top             =   5160
            Width           =   1215
         End
         Begin VB.OptionButton DungeonsHyrule 
            BackColor       =   &H00004000&
            Caption         =   "Temple"
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
            Height          =   210
            Index           =   0
            Left            =   3000
            TabIndex        =   21
            Top             =   5160
            Width           =   1215
         End
         Begin VB.OptionButton WorldsHyrule 
            BackColor       =   &H00004000&
            Caption         =   "Hyrule"
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
            Height          =   210
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   5160
            Width           =   1215
         End
         Begin VB.OptionButton CitiesHyrule 
            BackColor       =   &H00004000&
            Caption         =   "City"
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
            Height          =   210
            Index           =   0
            Left            =   1560
            TabIndex        =   19
            Top             =   5160
            Width           =   1215
         End
         Begin VB.Label lblCities 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gerudo"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   6
            Left            =   720
            TabIndex        =   40
            Top             =   1200
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label lblWorlds 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gerudo Canyon"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   6
            Left            =   1680
            TabIndex        =   39
            Top             =   480
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.Label lblDungeons 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Temple of the Spirit"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   4
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label CloseWorldMap 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   7680
            TabIndex        =   37
            Top             =   5160
            Width           =   735
         End
         Begin VB.Label lblDungeons 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Water Temple"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   36
            Top             =   2160
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label lblDungeons 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Temple of Fire"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   540
            Index           =   2
            Left            =   5400
            TabIndex        =   35
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblDungeons 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Temple of the Forest"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   1800
            TabIndex        =   34
            Top             =   4560
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.Label lblWorlds 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Lon Lon Ranch"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   4080
            TabIndex        =   33
            Top             =   3000
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label lblWorlds 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rio Zora"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   4
            Left            =   5880
            TabIndex        =   32
            Top             =   2400
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblWorlds 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hylia Lake"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   360
            TabIndex        =   31
            Top             =   3480
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lblWorlds 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mountain of Death"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   585
            Index           =   2
            Left            =   5520
            TabIndex        =   30
            Top             =   1320
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label lblWorlds 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Lost Forests"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   2160
            TabIndex        =   29
            Top             =   3720
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label lblCities 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gorons"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   4
            Left            =   5640
            TabIndex        =   28
            Top             =   960
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label lblCities 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Zoras"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   7200
            TabIndex        =   27
            Top             =   2160
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label lblCities 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hyrule"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   4080
            TabIndex        =   26
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblCities 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Kokiri"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   3120
            TabIndex        =   25
            Top             =   3000
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblCities 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Kakariko"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   6720
            TabIndex        =   24
            Top             =   840
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lblDungeons 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Temple of Shadows"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   540
            Index           =   5
            Left            =   7080
            TabIndex        =   23
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin VB.PictureBox MiniMapTermina 
         BackColor       =   &H00404000&
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
         Height          =   5520
         Index           =   1
         Left            =   -74760
         Picture         =   "frmMap.frx":8DC7C
         ScaleHeight     =   5520
         ScaleWidth      =   8580
         TabIndex        =   1
         Top             =   480
         Width           =   8580
         Begin VB.OptionButton CitiesTermina 
            BackColor       =   &H00404000&
            Caption         =   "City"
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
            Height          =   210
            Index           =   1
            Left            =   1560
            TabIndex        =   5
            Top             =   5160
            Width           =   1215
         End
         Begin VB.OptionButton WorldsTermina 
            BackColor       =   &H00404000&
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
            Height          =   210
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   5160
            Width           =   1215
         End
         Begin VB.OptionButton DungeonsTermina 
            BackColor       =   &H00404000&
            Caption         =   "Temple"
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
            Height          =   210
            Index           =   1
            Left            =   3000
            TabIndex        =   3
            Top             =   5160
            Width           =   1215
         End
         Begin VB.OptionButton NoneTermina 
            BackColor       =   &H00404000&
            Caption         =   "Nothing"
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
            Height          =   210
            Index           =   1
            Left            =   4320
            TabIndex        =   2
            Top             =   5160
            Width           =   1215
         End
         Begin VB.Label CloseWorldMap 
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
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   1
            Left            =   7680
            TabIndex        =   17
            Top             =   5160
            Width           =   735
         End
         Begin VB.Label lblWorlds 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Swamp"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   7
            Left            =   4080
            TabIndex        =   16
            Top             =   3600
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblWorlds 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Snowy Summit"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   8
            Left            =   3840
            TabIndex        =   15
            Top             =   120
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label lblWorlds 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Great Bay"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   9
            Left            =   720
            TabIndex        =   14
            Top             =   2760
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lblWorlds 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ikana Valley"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   10
            Left            =   6600
            TabIndex        =   13
            Top             =   3000
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblCities 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Clock City"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   7
            Left            =   3600
            TabIndex        =   12
            Top             =   3360
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label lblCities 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Goron Refuge"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   8
            Left            =   4440
            TabIndex        =   11
            Top             =   600
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label lblCities 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Zora City"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   9
            Left            =   960
            TabIndex        =   10
            Top             =   3360
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblDungeons 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Temple of the Swamp"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   6
            Left            =   3600
            TabIndex        =   9
            Top             =   4440
            Visible         =   0   'False
            Width           =   2595
         End
         Begin VB.Label lblDungeons 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Snowy Temple"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   7
            Left            =   5280
            TabIndex        =   8
            Top             =   960
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label lblDungeons 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Temple of the Great Bay"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   540
            Index           =   8
            Left            =   240
            TabIndex        =   7
            Top             =   1320
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label lblDungeons 
            Alignment       =   2  'Center
            BackColor       =   &H00A4D7DB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Temple of the Stone Tower"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   540
            Index           =   9
            Left            =   6240
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   2355
         End
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CloseWorldMap_Click(index As Integer)
    Me.Visible = False
    'play sound
    PlaySound Sound_ButtonClose
End Sub

Private Sub Form_Load()
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
End Sub

Private Sub WorldsHyrule_Click(index As Integer)
        lblWorlds(1).Visible = True
        lblWorlds(2).Visible = True
        lblWorlds(3).Visible = True
        lblWorlds(4).Visible = True
        lblWorlds(5).Visible = True
        lblWorlds(6).Visible = True
        lblCities(1).Visible = False
        lblCities(2).Visible = False
        lblCities(3).Visible = False
        lblCities(4).Visible = False
        lblCities(5).Visible = False
        lblCities(6).Visible = False
        lblDungeons(1).Visible = False
        lblDungeons(2).Visible = False
        lblDungeons(3).Visible = False
        lblDungeons(4).Visible = False
        lblDungeons(5).Visible = False
        ' play sound
        PlaySound Sound_ButtonChange
End Sub

Private Sub CitiesHyrule_Click(index As Integer)
        lblWorlds(1).Visible = False
        lblWorlds(2).Visible = False
        lblWorlds(3).Visible = False
        lblWorlds(4).Visible = False
        lblWorlds(5).Visible = False
        lblWorlds(6).Visible = False
        lblCities(1).Visible = True
        lblCities(2).Visible = True
        lblCities(3).Visible = True
        lblCities(4).Visible = True
        lblCities(5).Visible = True
        lblCities(6).Visible = True
        lblDungeons(1).Visible = False
        lblDungeons(2).Visible = False
        lblDungeons(3).Visible = False
        lblDungeons(4).Visible = False
        lblDungeons(5).Visible = False
        ' play sound
        PlaySound Sound_ButtonChange
End Sub
Private Sub DungeonsHyrule_Click(index As Integer)
        lblWorlds(1).Visible = False
        lblWorlds(2).Visible = False
        lblWorlds(3).Visible = False
        lblWorlds(4).Visible = False
        lblWorlds(5).Visible = False
        lblWorlds(6).Visible = False
        lblCities(1).Visible = False
        lblCities(2).Visible = False
        lblCities(3).Visible = False
        lblCities(4).Visible = False
        lblCities(5).Visible = False
        lblCities(6).Visible = False
        lblDungeons(1).Visible = True
        lblDungeons(2).Visible = True
        lblDungeons(3).Visible = True
        lblDungeons(4).Visible = True
        lblDungeons(5).Visible = True
        ' play sound
        PlaySound Sound_ButtonChange
End Sub
Private Sub NoneHyrule_Click(index As Integer)
        lblWorlds(1).Visible = False
        lblWorlds(2).Visible = False
        lblWorlds(3).Visible = False
        lblWorlds(4).Visible = False
        lblWorlds(5).Visible = False
        lblWorlds(6).Visible = False
        lblCities(1).Visible = False
        lblCities(2).Visible = False
        lblCities(3).Visible = False
        lblCities(4).Visible = False
        lblCities(5).Visible = False
        lblCities(6).Visible = False
        lblDungeons(1).Visible = False
        lblDungeons(2).Visible = False
        lblDungeons(3).Visible = False
        lblDungeons(4).Visible = False
        lblDungeons(5).Visible = False
        ' play sound
        PlaySound Sound_ButtonChange
End Sub

Private Sub WorldsTermina_Click(index As Integer)
        lblWorlds(7).Visible = True
        lblWorlds(8).Visible = True
        lblWorlds(9).Visible = True
        lblWorlds(10).Visible = True
        lblCities(7).Visible = False
        lblCities(8).Visible = False
        lblCities(9).Visible = False
        lblDungeons(6).Visible = False
        lblDungeons(7).Visible = False
        lblDungeons(8).Visible = False
        lblDungeons(9).Visible = False
        ' play sound
        PlaySound Sound_ButtonChange
End Sub
Private Sub CitiesTermina_Click(index As Integer)
        lblWorlds(7).Visible = False
        lblWorlds(8).Visible = False
        lblWorlds(9).Visible = False
        lblWorlds(10).Visible = False
        lblCities(7).Visible = True
        lblCities(8).Visible = True
        lblCities(9).Visible = True
        lblDungeons(6).Visible = False
        lblDungeons(7).Visible = False
        lblDungeons(8).Visible = False
        lblDungeons(9).Visible = False
        ' play sound
        PlaySound Sound_ButtonChange
End Sub
Private Sub DungeonsTermina_Click(index As Integer)
        lblWorlds(7).Visible = False
        lblWorlds(8).Visible = False
        lblWorlds(9).Visible = False
        lblWorlds(10).Visible = False
        lblCities(7).Visible = False
        lblCities(8).Visible = False
        lblCities(9).Visible = False
        lblDungeons(6).Visible = True
        lblDungeons(7).Visible = True
        lblDungeons(8).Visible = True
        lblDungeons(9).Visible = True
        ' play sound
        PlaySound Sound_ButtonChange
End Sub
Private Sub NoneTermina_Click(index As Integer)
        lblWorlds(7).Visible = False
        lblWorlds(8).Visible = False
        lblWorlds(9).Visible = False
        lblWorlds(10).Visible = False
        lblCities(7).Visible = False
        lblCities(8).Visible = False
        lblCities(9).Visible = False
        lblDungeons(6).Visible = False
        lblDungeons(7).Visible = False
        lblDungeons(8).Visible = False
        lblDungeons(9).Visible = False
        ' play sound
        PlaySound Sound_ButtonChange
End Sub

