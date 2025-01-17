VERSION 5.00
Begin VB.Form frmEditor_Item 
   Caption         =   "Item Editor"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   681
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   818
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraVitals 
      Caption         =   "Consume Data"
      Height          =   3135
      Left            =   3360
      TabIndex        =   48
      Top             =   4680
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CheckBox chkMPPercent 
         Caption         =   "Percent"
         Height          =   255
         Left            =   3720
         TabIndex        =   109
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chkHPPercent 
         Caption         =   "Percent"
         Height          =   255
         Left            =   3720
         TabIndex        =   108
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkInstant 
         Caption         =   "Instant Cast?"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCastSpell 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   64
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   120
         Max             =   30000
         TabIndex        =   62
         Top             =   1800
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         Max             =   30000
         TabIndex        =   60
         Top             =   1200
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   120
         Max             =   30000
         TabIndex        =   49
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblCastSpell 
         AutoSize        =   -1  'True
         Caption         =   "Cast Spell: None"
         Height          =   180
         Left            =   120
         TabIndex        =   65
         Top             =   2160
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         Caption         =   "Add Exp: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   63
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         Caption         =   "Add MP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   61
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         Caption         =   "Add HP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   50
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame fraWeight 
      Caption         =   "Weight"
      Height          =   1215
      Left            =   9840
      TabIndex        =   100
      Top             =   6600
      Width           =   2295
      Begin VB.TextBox txtWeight 
         Height          =   270
         Left            =   240
         TabIndex        =   101
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblWeight 
         Caption         =   "Weight: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   102
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CheckBox ChkTwoh 
      Caption         =   "ChkTwoh"
      Height          =   255
      Left            =   9840
      TabIndex        =   89
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Projectiles"
      Height          =   2055
      Left            =   3360
      TabIndex        =   75
      Top             =   7920
      Width           =   6255
      Begin VB.HScrollBar scrlDepth 
         Height          =   255
         Left            =   1560
         Max             =   255
         TabIndex        =   107
         Top             =   1680
         Width           =   1455
      End
      Begin VB.HScrollBar scrlAmmo 
         Height          =   255
         Left            =   1560
         TabIndex        =   85
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox ChkAmmo 
         Caption         =   "Ammunition"
         Height          =   375
         Left            =   120
         TabIndex        =   84
         Top             =   1080
         Width           =   1215
      End
      Begin VB.HScrollBar scrlProjectileDamage 
         Height          =   255
         Left            =   4560
         Max             =   5000
         TabIndex        =   83
         Top             =   720
         Width           =   1455
      End
      Begin VB.HScrollBar scrlProjectileRange 
         Height          =   255
         Left            =   1560
         Max             =   100
         TabIndex        =   82
         Top             =   720
         Width           =   1455
      End
      Begin VB.HScrollBar scrlProjectilePic 
         Height          =   255
         Left            =   1560
         Max             =   500
         TabIndex        =   81
         Top             =   240
         Width           =   1455
      End
      Begin VB.HScrollBar scrlProjectileSpeed 
         Height          =   255
         Left            =   4560
         Max             =   200
         TabIndex        =   80
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblDepth 
         Caption         =   "Depth: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblammo 
         Height          =   255
         Left            =   3120
         TabIndex        =   86
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label lblProjectileDamage 
         Caption         =   "Damage: 0"
         Height          =   255
         Left            =   3240
         TabIndex        =   79
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblProjectilesSpeed 
         Caption         =   "Speed: 0"
         Height          =   255
         Left            =   3240
         TabIndex        =   78
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblProjectileRange 
         Caption         =   "Range: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblProjectilePic 
         Caption         =   "Pic: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   3375
      Left            =   3360
      TabIndex        =   16
      Top             =   120
      Width           =   8655
      Begin VB.Frame Frame6 
         Caption         =   "Extra HP"
         Height          =   975
         Left            =   5760
         TabIndex        =   117
         Top             =   2280
         Width           =   2655
         Begin VB.HScrollBar spinExtHp 
            Height          =   255
            Left            =   240
            Max             =   30000
            TabIndex        =   119
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label lblExtraHP 
            Caption         =   "Extra HP: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   118
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Impact"
         Height          =   1215
         Left            =   5760
         TabIndex        =   114
         Top             =   1080
         Visible         =   0   'False
         Width           =   2775
         Begin VB.CheckBox chkImpactarF 
            Caption         =   "Activate Without Impact"
            Height          =   255
            Left            =   600
            TabIndex        =   120
            Top             =   840
            Width           =   1935
         End
         Begin VB.HScrollBar spinImpactar 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   115
            Top             =   480
            Value           =   1
            Width           =   2535
         End
         Begin VB.Label lblImpactar 
            Caption         =   "Impactar: None"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.HScrollBar scrlArmyRangeReq 
         Height          =   255
         Left            =   7200
         TabIndex        =   113
         Top             =   600
         Width           =   1335
      End
      Begin VB.HScrollBar scrlArmyTypeReq 
         Height          =   255
         Left            =   7200
         TabIndex        =   112
         Top             =   240
         Width           =   1335
      End
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         Left            =   4200
         TabIndex        =   87
         Top             =   3000
         Width           =   1455
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4200
         Max             =   99
         TabIndex        =   73
         Top             =   2700
         Width           =   1455
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   71
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0A4E
         Left            =   3840
         List            =   "frmEditor_Item.frx":0A50
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2040
         Width           =   1815
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtDesc 
         Height          =   1455
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   59
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   24
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0A52
         Left            =   4200
         List            =   "frmEditor_Item.frx":0A68
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPrice 
         Height          =   255
         LargeChange     =   10
         Left            =   4200
         Max             =   30000
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0AA2
         Left            =   120
         List            =   "frmEditor_Item.frx":0AD6
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
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
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   17
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblArmyRange 
         Caption         =   "Range Req: None"
         Height          =   495
         Left            =   5760
         TabIndex        =   111
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblArmyType 
         Caption         =   "Army Req: None"
         Height          =   255
         Left            =   5760
         TabIndex        =   110
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblItem 
         Caption         =   "Item: None"
         Height          =   330
         Left            =   2880
         TabIndex        =   88
         Top             =   2955
         Width           =   1095
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   74
         Top             =   2700
         Width           =   900
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   72
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Class Req:"
         Height          =   180
         Left            =   2880
         TabIndex        =   70
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2880
         TabIndex        =   67
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   30
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Special Type:"
         Height          =   180
         Left            =   2880
         TabIndex        =   29
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   28
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   27
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   9840
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   9975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   9600
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirements"
      Height          =   975
      Left            =   3360
      TabIndex        =   5
      Top             =   3600
      Width           =   6255
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   14
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   13
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   11
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   480
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   3135
      Left            =   3360
      TabIndex        =   31
      Top             =   4680
      Visible         =   0   'False
      Width           =   6255
      Begin VB.PictureBox picPaperdoll 
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
         Height          =   1080
         Left            =   120
         ScaleHeight     =   72
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   400
         TabIndex        =   57
         Top             =   1920
         Width           =   6000
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   5040
         TabIndex        =   56
         Top             =   1560
         Width           =   1095
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   4560
         Max             =   3000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   39
         Top             =   840
         Value           =   100
         Width           =   1575
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   38
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   37
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5280
         Max             =   255
         TabIndex        =   36
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   35
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   1320
         Max             =   255
         TabIndex        =   34
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0B64
         Left            =   1320
         List            =   "frmEditor_Item.frx":0B7A
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   360
         Width           =   4815
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   32
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         Caption         =   "Paperdoll: 0"
         Height          =   180
         Left            =   3960
         TabIndex        =   55
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 0.1 sec"
         Height          =   180
         Left            =   3240
         TabIndex        =   47
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2160
         TabIndex        =   46
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   45
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4440
         TabIndex        =   44
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End: 0"
         Height          =   180
         Index           =   2
         Left            =   2160
         TabIndex        =   43
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         Caption         =   "Damage: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Object Tool:"
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1215
      Left            =   3360
      TabIndex        =   51
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1080
         Max             =   255
         Min             =   1
         TabIndex        =   52
         Top             =   720
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   53
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.Frame fraBag 
      Caption         =   "Bags"
      Height          =   4215
      Left            =   3360
      TabIndex        =   90
      Top             =   3600
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlBag 
         Height          =   255
         Left            =   240
         Max             =   10
         TabIndex        =   91
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblBag 
         Caption         =   "Add Bags: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   92
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fraAddWeight 
      Caption         =   "AddWeight"
      Height          =   4215
      Left            =   3360
      TabIndex        =   103
      Top             =   3600
      Width           =   6255
      Begin VB.HScrollBar scrlAddWeight 
         Height          =   255
         Left            =   360
         Max             =   10000
         TabIndex        =   104
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblAddWeight 
         Caption         =   "Add Weight: 0"
         Height          =   255
         Left            =   360
         TabIndex        =   105
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame frameContainer 
      Caption         =   "Container"
      Height          =   4215
      Left            =   3360
      TabIndex        =   93
      Top             =   3600
      Width           =   6255
      Begin VB.HScrollBar scrlContainer 
         Height          =   255
         Left            =   240
         Max             =   800
         TabIndex        =   96
         Top             =   1560
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAmount 
         Height          =   255
         Left            =   2640
         Max             =   100
         TabIndex        =   95
         Top             =   1560
         Width           =   1095
      End
      Begin VB.HScrollBar scrlContainerIndex 
         Height          =   255
         Left            =   240
         Max             =   10
         TabIndex        =   94
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblContainer 
         Caption         =   "Item: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   99
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount: 0"
         Height          =   255
         Left            =   2640
         TabIndex        =   98
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblContainerIndex 
         Caption         =   "Container Index: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   97
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastIndex As Long

Private Sub ChkAmmo_Click()
 If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
Item(EditorIndex).ammoreq = ChkAmmo.value
End Sub

Private Sub chkExpPercent_Click()

End Sub

Private Sub chkHPPercent_Click()
    Item(EditorIndex).AddHPPercent = chkHPPercent.value
    'Call CheckItemPercentChange(EditorIndex, CBool(chkHPPercent.Value), 1)
End Sub



Private Sub chkMPPercent_Click()
    Item(EditorIndex).AddMPPercent = chkMPPercent.value
    'Call CheckItemPercentChange(EditorIndex, CBool(chkMPPercent.Value), 2)
End Sub

Private Sub ChkTwoh_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ChkTwoh.value = 0 Then
        Item(EditorIndex).istwohander = False
    Else
        Item(EditorIndex).istwohander = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkTwoh", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbBind_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBind_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).Sound = cmbSound.list(cmbSound.ListIndex)
    Else
        Item(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlPic.Max = NumItems
    scrlAnim.Max = MAX_ANIMATIONS
    scrlPaperdoll.Max = NumPaperdolls
    scrlAmmo.Max = MAX_ITEMS
    scrlItem.Max = MAX_ITEMS
    scrlContainer.Max = MAX_ITEMS
    scrlContainerIndex.Max = MAX_ITEM_CONTAINERS
    scrlBag.Max = MAX_RUPEE_BAGS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.Visible = True
        Frame4.Visible = True
        Frame6.Visible = True
        'scrlDamage_Change
    Else
        Frame6.Visible = False
        fraEquipment.Visible = False
    End If
    
     If (cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
        Frame5.Visible = True
    Else
        Frame5.Visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.Visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_CONTAINER) Then
        frameContainer.Visible = True
        Frame1.Visible = False
    Else
        frameContainer.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_BAG) Then
        fraBag.Visible = True
    Else
        fraBag.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_ADDWEIGHT) Then
        fraAddWeight.Visible = True
    Else
        fraAddWeight.Visible = False
    End If
    
    Item(EditorIndex).Type = cmbType.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ItemEditorInit
    DoEvents
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccessReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.Caption = "Access Req: " & scrlAccessReq.value
    Item(EditorIndex).AccessReq = scrlAccessReq.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddHp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddHP.Caption = "Add HP: " & scrlAddHp.value
    Item(EditorIndex).AddHP = scrlAddHp.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddMp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddMP.Caption = "Add MP: " & scrlAddMP.value
    Item(EditorIndex).AddMP = scrlAddMP.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddExp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddExp.Caption = "Add Exp: " & scrlAddExp.value
    Item(EditorIndex).AddEXP = scrlAddExp.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddWeight_Change()
    Item(EditorIndex).Data1 = scrlAddWeight.value
    lblAddWeight.Caption = "Add Weight: " & scrlAddWeight.value
End Sub

Private Sub scrlAmmo_Change()
If scrlAmmo.value > 0 Then
lblammo.Caption = "Weapon: " + Item(scrlAmmo.value).Name
Else
lblammo.Caption = "Weapon: None"
End If

Item(EditorIndex).ammo = scrlAmmo.value
End Sub

Private Sub scrlAnim_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.value).Name)
    End If
    lblAnim.Caption = "Anim: " & sString
    Item(EditorIndex).Animation = scrlAnim.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlArmyRangeReq_Change()
    lblArmyRange.Caption = "Range Req: " & RangeToStr(scrlArmyRangeReq.value, scrlArmyTypeReq.value)
    Item(EditorIndex).ArmyRange_Req = scrlArmyRangeReq.value
    
End Sub

Private Sub scrlArmyTypeReq_Change()
    lblArmyType.Caption = "Army Req: " & JusticeToStr(scrlArmyTypeReq.value)
    Item(EditorIndex).ArmyType_Req = scrlArmyTypeReq.value
    CheckItemEditorRangeScrolls
End Sub

Private Sub scrlBag_Change()

Item(EditorIndex).AddBags = scrlBag.value
lblBag.Caption = "Add Bags: " & scrlBag.value

End Sub




Private Sub scrlDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblDamage.Caption = "Damage: " & scrlDamage.value
    Item(EditorIndex).Data2 = scrlDamage.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDepth_Change()
    Item(EditorIndex).ProjecTile.Depth = scrlDepth.value
    lblDepth.Caption = "Depth: " & scrlDepth.value
End Sub

Private Sub scrlItem_Change()
If scrlItem.value = 0 Then
         lblItem.Caption = "Item: None"
         Exit Sub
End If

lblItem.Caption = "Item: " & Item(scrlItem.value).Name
Item(EditorIndex).ConsumeItem = scrlItem.value
End Sub

Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = "Level req: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.value
    Call EditorItem_BltPaperdoll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Pic: " & scrlPic.value
    Item(EditorIndex).Pic = scrlPic.value
    Call EditorItem_BltItem
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPrice_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPrice.Caption = "Price: " & scrlPrice.value
    Item(EditorIndex).Price = scrlPrice.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRarity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.Caption = "Rarity: " & scrlRarity.value
    Item(EditorIndex).Rarity = scrlRarity.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "Speed: " & scrlSpeed.value / 1000 & " sec"
    Item(EditorIndex).Speed = scrlSpeed.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatBonus_Change(index As Integer)
Dim text As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case index
        Case 1
            text = "+ Str: "
        Case 2
            text = "+ End: "
        Case 3
            text = "+ Int: "
        Case 4
            text = "+ Agi: "
        Case 5
            text = "+ Will: "
    End Select
            
    lblStatBonus(index).Caption = text & scrlStatBonus(index).value
    Item(EditorIndex).Add_Stat(index) = scrlStatBonus(index).value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatBonus_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatReq_Change(index As Integer)
    Dim text As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case index
        Case 1
            text = "Str: "
        Case 2
            text = "End: "
        Case 3
            text = "Int: "
        Case 4
            text = "Agi: "
        Case 5
            text = "Will: "
    End Select
    
    lblStatReq(index).Caption = text & scrlStatReq(index).value
    Item(EditorIndex).Stat_Req(index) = scrlStatReq(index).value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If Len(Trim$(Spell(scrlSpell.value).Name)) > 0 Then
        lblSpellName.Caption = "Name: " & Trim$(Spell(scrlSpell.value).Name)
    Else
        lblSpellName.Caption = "Name: None"
    End If
    
    lblSpell.Caption = "Spell: " & scrlSpell.value
    
    Item(EditorIndex).Data1 = scrlSpell.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub spinExtHp_Change()
  ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

  
    lblExtraHP.Caption = "Extra HP: " & spinExtHp.value
    Item(EditorIndex).ExtraHP = spinExtHp.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub spinImpactar_Change()
  
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If spinImpactar.value > 0 Then

        If Len(Trim$(Spell(spinImpactar.value).Name)) > 0 Then
            lblImpactar.Caption = "Impactar: " & Trim$(Spell(spinImpactar.value).Name)
        End If

    Else
        lblImpactar.Caption = "Impactar: None"
    End If
    
    
    Item(EditorIndex).Impactar.Spell = spinImpactar.value
    Exit Sub
End Sub

Private Sub chkImpactarF_Click()
 If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Impactar.Auto = chkImpactarF.value
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectileDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileDamage.Caption = "Damage: " & scrlProjectileDamage.value
    Item(EditorIndex).ProjecTile.Damage = scrlProjectileDamage.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectilePic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectilePic.Caption = "Pic: " & scrlProjectilePic.value
    Item(EditorIndex).ProjecTile.Pic = scrlProjectilePic.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ProjecTile
Private Sub scrlProjectileRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRange.Caption = "Range: " & scrlProjectileRange.value
    Item(EditorIndex).ProjecTile.range = scrlProjectileRange.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRange_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectileSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectilesSpeed.Caption = "Speed: " & scrlProjectileSpeed.value
    Item(EditorIndex).ProjecTile.Speed = scrlProjectileSpeed.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    lblAmount.Caption = "Amount: " & scrlAmount.value
    Item(EditorIndex).Container(scrlContainerIndex.value).value = scrlAmount.value
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAmount_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub scrlContainer_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If scrlContainer.value = 0 Then
        lblContainer.Caption = "Item: 0"
    Else
        lblContainer.Caption = "Item: " & Item(scrlContainer.value).Name
    End If
    Item(EditorIndex).Container(scrlContainerIndex.value).ItemNum = scrlContainer.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlContainer_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlContainerIndex_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If scrlContainerIndex.value > MAX_ITEM_CONTAINERS Then Exit Sub
    
    lblContainerIndex.Caption = "Container Index: " & scrlContainerIndex.value
    
    If Item(EditorIndex).Container(scrlContainerIndex.value).ItemNum <= MAX_ITEMS And Item(EditorIndex).Container(scrlContainerIndex.value).ItemNum > 0 Then
        scrlContainer.value = Item(EditorIndex).Container(scrlContainerIndex.value).ItemNum
    End If
    
    If Item(EditorIndex).Container(scrlContainerIndex.value).value <= scrlAmount.Max Then
        scrlAmount.value = Item(EditorIndex).Container(scrlContainerIndex.value).value
    End If
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlContainerIndex_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

