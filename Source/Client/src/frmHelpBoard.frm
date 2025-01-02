VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHelpBoard 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5805
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab HelpBoard 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabHeight       =   529
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Controls"
      TabPicture(0)   =   "frmHelpBoard.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CloseHelpBoard(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Training"
      TabPicture(1)   =   "frmHelpBoard.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CloseHelpBoard(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Picture7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Stats"
      TabPicture(2)   =   "frmHelpBoard.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "CloseHelpBoard(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Picture8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Abilities"
      TabPicture(3)   =   "frmHelpBoard.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CloseHelpBoard(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Picture9(0)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Pets"
      TabPicture(4)   =   "frmHelpBoard.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "CloseHelpBoard(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Picture9(1)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   120
         ScaleHeight     =   4635
         ScaleWidth      =   5355
         TabIndex        =   39
         Top             =   720
         Width           =   5415
         Begin VB.Label Stats 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Force"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   49
            Top             =   240
            Width           =   3885
         End
         Begin VB.Label Stats 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Agility"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   48
            Top             =   960
            Width           =   3885
         End
         Begin VB.Label Stats 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Defense"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   47
            Top             =   2040
            Width           =   3885
         End
         Begin VB.Label Stats 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Spirit"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   46
            Top             =   2880
            Width           =   3885
         End
         Begin VB.Label Stats 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Intelligence"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   45
            Top             =   3720
            Width           =   3885
         End
         Begin VB.Label Stats 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Increases the melee attack with weapons"
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
            Height          =   495
            Index           =   5
            Left            =   1080
            TabIndex        =   44
            Top             =   480
            Width           =   3195
         End
         Begin VB.Label Stats 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Increases the chance of dodging attack and increases the attack with projectiles"
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
            Height          =   615
            Index           =   6
            Left            =   1080
            TabIndex        =   43
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label Stats 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Increases the user's melee defense"
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
            Height          =   495
            Index           =   7
            Left            =   960
            TabIndex        =   42
            Top             =   2280
            Width           =   3375
         End
         Begin VB.Label Stats 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "It raises the regeneration of life and energy and raises the magical defense"
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
            Height          =   495
            Index           =   8
            Left            =   960
            TabIndex        =   41
            Top             =   3120
            Width           =   3375
         End
         Begin VB.Label Stats 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Expand the energy and power up the magic attack"
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
            Height          =   495
            Index           =   9
            Left            =   1080
            TabIndex        =   40
            Top             =   3960
            Width           =   3135
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   -74880
         ScaleHeight     =   4275
         ScaleWidth      =   5355
         TabIndex        =   33
         Top             =   720
         Width           =   5415
         Begin VB.Label lblKeys 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Grabbing things off the floor: ENTER"
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
            Index           =   2
            Left            =   840
            TabIndex        =   38
            Top             =   1920
            Width           =   3855
         End
         Begin VB.Label lblKeys 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Attack/Talk to NPC: Cntrl"
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
            Index           =   0
            Left            =   840
            TabIndex        =   37
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label lblKeys 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Moving around the world: Arrows"
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
            Left            =   840
            TabIndex        =   36
            Top             =   1200
            Width           =   3855
         End
         Begin VB.Label lblKeys 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Running while moving: Shift"
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
            Index           =   3
            Left            =   840
            TabIndex        =   35
            Top             =   2640
            Width           =   3855
         End
         Begin VB.Label lblKeys 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Use items or abilities using F1, F2... F12: Drag item or Skill to the desired skill bar box"
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
            Height          =   735
            Index           =   4
            Left            =   840
            TabIndex        =   34
            Top             =   3240
            Width           =   3855
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   -74880
         ScaleHeight     =   4635
         ScaleWidth      =   5355
         TabIndex        =   18
         Top             =   720
         Width           =   5415
         Begin VB.Label lblPlaces 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Gerudo Canyon, Gerudo Fortress, Desert, Temple of the Spirit"
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
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   32
            Top             =   3000
            Width           =   5055
         End
         Begin VB.Label lblLevels 
            BackStyle       =   0  'Transparent
            Caption         =   "Level 40-50"
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
            Index           =   4
            Left            =   120
            TabIndex        =   31
            Top             =   2760
            Width           =   3255
         End
         Begin VB.Label lblLevels 
            BackStyle       =   0  'Transparent
            Caption         =   "Level 5-10"
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
            Left            =   120
            TabIndex        =   30
            Top             =   0
            Width           =   3255
         End
         Begin VB.Label lblLevels 
            BackStyle       =   0  'Transparent
            Caption         =   "Level 10-20"
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
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   3255
         End
         Begin VB.Label lblPlaces 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Deku Tree, Dodongo Cavern, Great Jabu Jabu, Hyrule Fields, Zora River"
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
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   5085
         End
         Begin VB.Label lblPlaces 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Lost Forests, Forest Temple, Swamp, Swamp Temple, Ruins"
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
            Height          =   615
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   5055
         End
         Begin VB.Label lblLevels 
            BackStyle       =   0  'Transparent
            Caption         =   "Level 20-30"
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
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Label lblPlaces 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Death Mountain, Fire Temple, Snowy Pass, Ice Temple"
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
            Height          =   615
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   1560
            Width           =   5055
         End
         Begin VB.Label lblLevels 
            BackStyle       =   0  'Transparent
            Caption         =   "Level 30-40"
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
            Index           =   3
            Left            =   120
            TabIndex        =   24
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label lblPlaces 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hylia Lake, Water Temple, Cemetery, Tombs and Royal Tomb"
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
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   23
            Top             =   2280
            Width           =   5055
         End
         Begin VB.Label lblLevels 
            BackStyle       =   0  'Transparent
            Caption         =   "Level 50-60"
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
            Index           =   5
            Left            =   120
            TabIndex        =   22
            Top             =   3480
            Width           =   3255
         End
         Begin VB.Label lblPlaces 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cave of Aquamentus, Secret Sanctuary, Castle of Ganon"
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
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   21
            Top             =   3720
            Width           =   5055
         End
         Begin VB.Label lblLevels 
            BackStyle       =   0  'Transparent
            Caption         =   "Level 60-70"
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
            Index           =   6
            Left            =   120
            TabIndex        =   20
            Top             =   4080
            Width           =   3255
         End
         Begin VB.Label lblPlaces 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cyclopean Clouds, Pyramid"
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
            Index           =   6
            Left            =   120
            TabIndex        =   19
            Top             =   4320
            Width           =   5055
         End
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Index           =   0
         Left            =   -74880
         ScaleHeight     =   4275
         ScaleWidth      =   5355
         TabIndex        =   8
         Top             =   720
         Width           =   5415
         Begin VB.Label lblSpells 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "All classes can use skills, and everyone has different from the others"
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
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   4875
         End
         Begin VB.Label lblSpells 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hylian"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   840
            Width           =   4875
         End
         Begin VB.Label lblSpells 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Goron"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   15
            Top             =   1680
            Width           =   4875
         End
         Begin VB.Label lblSpells 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Zora"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   14
            Top             =   2520
            Width           =   4875
         End
         Begin VB.Label lblSpells 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "In the Temple of Time, in the city of Hyrule, to the northeast"
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
            Height          =   495
            Index           =   4
            Left            =   720
            TabIndex        =   13
            Top             =   1080
            Width           =   3915
         End
         Begin VB.Label lblSpells 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "In the hall of Darunia, in Goron City, north of the central hall"
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
            Height          =   495
            Index           =   5
            Left            =   720
            TabIndex        =   12
            Top             =   1920
            Width           =   3915
         End
         Begin VB.Label lblSpells 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "In the hall of King Zora, in the City Zora, north of the central hall"
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
            Height          =   495
            Index           =   6
            Left            =   720
            TabIndex        =   11
            Top             =   2760
            Width           =   3915
         End
         Begin VB.Label lblSpells 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "In the square of the Gerudo Fortress, next to the mountain on the left"
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
            Height          =   495
            Index           =   10
            Left            =   720
            TabIndex        =   10
            Top             =   3600
            Width           =   3915
         End
         Begin VB.Label lblSpells 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Gerudo"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   9
            Top             =   3360
            Width           =   4875
         End
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   4275
         ScaleWidth      =   5355
         TabIndex        =   1
         Top             =   720
         Width           =   5415
         Begin VB.Label lblPets 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Taming Creature"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Width           =   4875
         End
         Begin VB.Label lblSpells 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"frmHelpBoard.frx":008C
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
            Height          =   1095
            Index           =   7
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   5115
         End
         Begin VB.Label lblPets 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Benefit"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   1440
            Width           =   4875
         End
         Begin VB.Label lblSpells 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"frmHelpBoard.frx":0159
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
            Height          =   1095
            Index           =   8
            Left            =   120
            TabIndex        =   4
            Top             =   1680
            Width           =   5115
         End
         Begin VB.Label lblPets 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Use Pet"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   3
            Top             =   2760
            Width           =   4875
         End
         Begin VB.Label lblSpells 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "To be able to use a pet, choose it, having clicked on the ""Pet"" button, from the pet selection bar and then click on ""Accept"""
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
            Height          =   855
            Index           =   9
            Left            =   120
            TabIndex        =   2
            Top             =   3000
            Width           =   5115
         End
      End
      Begin VB.Label CloseHelpBoard 
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
         Height          =   255
         Index           =   0
         Left            =   -70320
         TabIndex        =   54
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label CloseHelpBoard 
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
         Height          =   255
         Index           =   1
         Left            =   -70320
         TabIndex        =   53
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label CloseHelpBoard 
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
         Height          =   255
         Index           =   2
         Left            =   4680
         TabIndex        =   52
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label CloseHelpBoard 
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
         Height          =   255
         Index           =   3
         Left            =   -70320
         TabIndex        =   51
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label CloseHelpBoard 
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
         Height          =   255
         Index           =   4
         Left            =   -70320
         TabIndex        =   50
         Top             =   5040
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmHelpBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseHelpBoard_Click(index As Integer)
    Me.Visible = False
End Sub

Private Sub Close1HelpBoard_Click()
    Me.Visible = False
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

Private Sub HelpBoard_DblClick()

End Sub
