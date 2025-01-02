VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Custom Server"
      Height          =   255
      Left            =   5880
      TabIndex        =   38
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Troll Server"
      Height          =   255
      Left            =   5880
      TabIndex        =   37
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Normal Server"
      Height          =   255
      Left            =   5880
      TabIndex        =   36
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox picCharacter 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   525
      ScaleHeight     =   243
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   16
      Top             =   450
      Visible         =   0   'False
      Width           =   6630
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   4800
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   26
         Top             =   1320
         Width           =   480
      End
      Begin VB.ComboBox cmbClass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmMenu.frx":0A4E
         Left            =   2280
         List            =   "frmMenu.frx":0A50
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton optMale 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Male"
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
         Left            =   2280
         TabIndex        =   19
         Top             =   1935
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optFemale 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Female"
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
         Left            =   3360
         TabIndex        =   18
         Top             =   1935
         Width           =   1095
      End
      Begin VB.TextBox txtCName 
         Appearance      =   0  'Flat
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
         Height          =   225
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   21
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblClassInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   855
         Left            =   1200
         TabIndex        =   29
         Top             =   2550
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label lblSprite 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[Change Appearance]"
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
         Left            =   2700
         TabIndex        =   25
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
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
         Left            =   1320
         TabIndex        =   24
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Clase:"
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
         Left            =   1560
         TabIndex        =   23
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblCAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
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
         Left            =   2760
         TabIndex        =   17
         Top             =   2280
         Width           =   1215
      End
   End
   Begin VB.PictureBox picRegister 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   525
      ScaleHeight     =   243
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   7
      Top             =   450
      Visible         =   0   'False
      Width           =   6630
      Begin VB.TextBox txtRPass2 
         Appearance      =   0  'Flat
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
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtRPass 
         Appearance      =   0  'Flat
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
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtRUser 
         Appearance      =   0  'Flat
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
         Height          =   225
         Left            =   2520
         MaxLength       =   12
         TabIndex        =   8
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Rewriting:"
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
         Left            =   1200
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label txtRAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
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
         Left            =   2760
         TabIndex        =   12
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Index           =   7
         Left            =   1200
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblBlank 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.PictureBox picCredits 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   525
      ScaleHeight     =   3645
      ScaleWidth      =   6630
      TabIndex        =   15
      Top             =   450
      Visible         =   0   'False
      Width           =   6630
      Begin VB.Label Label1 
         BackColor       =   &H80000008&
         Caption         =   "Managers: Farid, Nicoxlitox"
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
         Index           =   4
         Left            =   480
         TabIndex        =   35
         Top             =   2400
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000008&
         Caption         =   "Graphics: Dace, Sebas"
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
         Left            =   480
         TabIndex        =   34
         Top             =   2040
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000008&
         Caption         =   "Mappers: Dace, Luis Lara, Rolexgamer, Kevin"
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
         Left            =   480
         TabIndex        =   33
         Top             =   1680
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000008&
         Caption         =   "Programmers: Joan, Dace"
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
         Left            =   480
         TabIndex        =   32
         Top             =   1320
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000008&
         Caption         =   "Project Manager: Dace"
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
         Left            =   480
         TabIndex        =   31
         Top             =   960
         Width           =   5535
      End
   End
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1350
      Left            =   525
      ScaleHeight     =   1350
      ScaleWidth      =   3000
      TabIndex        =   27
      Top             =   2760
      Width           =   3000
      Begin VB.Label lblNews 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Zelda Online"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   120
         TabIndex        =   28
         Top             =   75
         Width           =   2775
      End
   End
   Begin VB.PictureBox picLogin 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   525
      ScaleHeight     =   242
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   6630
      Begin VB.CheckBox chkPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Save Password?"
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
         Left            =   2520
         TabIndex        =   5
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtLPass 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtLUser 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2520
         MaxLength       =   12
         TabIndex        =   1
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label ServerStatus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   4080
         TabIndex        =   30
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblLAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
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
         Left            =   2760
         TabIndex        =   6
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Left            =   1200
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Label lblVer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   39
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Image imgButton 
      Height          =   465
      Index           =   4
      Left            =   5520
      Top             =   4440
      Width           =   1665
   End
   Begin VB.Image imgButton 
      Height          =   465
      Index           =   3
      Left            =   3840
      Top             =   4440
      Width           =   1665
   End
   Begin VB.Image imgButton 
      Height          =   465
      Index           =   2
      Left            =   2160
      Top             =   4440
      Width           =   1665
   End
   Begin VB.Image imgButton 
      Height          =   465
      Index           =   1
      Left            =   480
      Top             =   4440
      Width           =   1665
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClass_Click()
Dim lblClassInfo As String

    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    NewCharacterBltSprite
    
    frmMenu.lblClassInfo.Visible = True
    
    If newCharClass = 0 Then
    frmMenu.lblClassInfo.Caption = "The Guardian of Light is adept at handling weapons and is very well-balanced in hand-to-hand combat"
    ElseIf newCharClass = 1 Then
        frmMenu.optMale.Visible = True
        frmMenu.optFemale.Visible = True
        frmMenu.lblClassInfo.Caption = "The Dark Guardian wields certain weapons and possesses powerful arcane spells with which to conjure"
    ElseIf newCharClass = 2 Then
        frmMenu.optMale.Visible = True
        frmMenu.optFemale.Visible = True
        frmMenu.lblClassInfo.Caption = "The Light Magician uses electric power to paralyze and attack, as well as to heal wounds"
    ElseIf newCharClass = 3 Then
        frmMenu.optMale.Visible = True
        frmMenu.optFemale.Visible = True
        frmMenu.lblClassInfo.Caption = "The Elemental Mage controls fire and ice as a means of burning or freezing his opponents"
    ElseIf newCharClass = 4 Then
        frmMenu.optMale.Visible = True
        frmMenu.optFemale.Visible = True
        frmMenu.lblClassInfo.Caption = "The Fighting Goron is a brave and strong fighter who fights without weapons and hand-to-hand"
    ElseIf newCharClass = 5 Then
        frmMenu.optMale.Visible = True
        frmMenu.optFemale.Visible = True
        frmMenu.lblClassInfo.Caption = "The Explosive Goron dominates fire and explosions and is very resistant, withstanding many blows"
    ElseIf newCharClass = 6 Then
        frmMenu.optMale.Visible = True
        frmMenu.optFemale.Visible = True
        frmMenu.lblClassInfo.Caption = "The Mystic Zora uses the nature of water to defend herself from her enemies or to heal herself"
    ElseIf newCharClass = 7 Then
        frmMenu.optMale.Visible = True
        frmMenu.optFemale.Visible = True
        frmMenu.lblClassInfo.Caption = "The Stormy Zora uses his knowledge about nature to attack his opponents from afar"
    ElseIf newCharClass = 8 Then
        frmMenu.optMale.Visible = False
        frmMenu.optFemale.value = True
        frmMenu.lblClassInfo.Caption = "The Gerudo Warrior is an excellent hand-to-hand combatant, who is also lethal in throwing weapons"
    ElseIf newCharClass = 9 Then
        frmMenu.optMale.Visible = False
        frmMenu.optFemale.value = True
        frmMenu.lblClassInfo.Caption = "The Gerudo Sorceress uses devastating magical powers to defeat her opponents, her specialty is magic"
    Else
    frmMenu.lblClassInfo.Caption = "Class without information"
    End If
    
    frmMenu.lblClassInfo.Caption = frmMenu.lblClassInfo.Caption
    
End Sub

Private Sub Command1_Click()
DestroyTCP
Options.ip = "127.0.0.1"
Options.port = "4000"
frmMain.Caption = Options.Game_Name & " - Official Server"
TcpInit
End Sub

Private Sub Command2_Click()

If MsgBox("The 'Troll Server' is an evil place where everyone is an admin" & vbNewLine & "However, kicking/banning is disabled." & vbNewLine & _
    "You may edit maps, items, spells, quests, whatever." & vbNewLine & "You can even kill other players." & vbNewLine & _
    "type /admin in chat, or press the Insert key to open the admin menu." & vbNewLine & vbNewLine & _
    "Would you like to play on this server?", vbYesNo, "Play on the Troll Server?") = vbNo Then Exit Sub
    

DestroyTCP
Options.ip = "127.0.0.1"
Options.port = "4001"
frmMain.Caption = Options.Game_Name & " - Troll Server"
TcpInit
End Sub

Private Sub Command3_Click()
DestroyTCP

Dim ip As String, port As String
ip = InputBox("Please enter a custom server's Address..", "Server Address", "127.0.0.1")
If LenB(ip) <= 0 Then Exit Sub
port = InputBox("Please enter a server port..", "Server Port", "4000")
If LenB(port) <= 0 Then Exit Sub

Options.ip = ip
Options.port = port

frmMain.Caption = Options.Game_Name & " - " & ip

TcpInit
End Sub

Private Sub Form_Load()
    Dim tmpTxt As String, tmpArray() As String, i As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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
    
    lblVer.Caption = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision
    
    'For i = 1 To lblBlank.UBound
    '    With lblBlank(i)
    '        .Caption = GetTranslation(.Caption)
    '    End With
    'Next i
    
    'With lblCAccept
    '.Caption = GetTranslation(.Caption)
    'End With
    
    
    
    ' general menu stuff
    Me.Caption = Options.Game_Name
    lblNews.Caption = vbNullString
    ' load news
    Open App.Path & "\data\news.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, tmpTxt
        lblNews.Caption = lblNews.Caption & tmpTxt & vbNewLine
    Loop
    Close #1

    'For i = 0 To UBound(tmpArray)
    '    lblNews.Caption = lblNews.Caption & tmpArray(i) & vbNewLine
    'Next

    ' Load the username + pass
    txtLUser.text = Trim$(Options.Username)
    If Options.SavePass = 1 Then
        txtLPass.text = Trim$(Options.Password)
        chkPass.value = Options.SavePass
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not EnteringGame Then DestroyGame
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_Click(index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case index
        Case 1
            If Not picLogin.Visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.Visible = False
                picLogin.Visible = True
                picRegister.Visible = False
                picCharacter.Visible = False
                picMain.Visible = False
                If Len(txtLUser.text) = 0 Then
                    txtLUser.SetFocus
                Else
                    txtLPass.SetFocus
                    txtLPass.SelLength = Len(txtLPass.text)
                End If

                ' play sound
                PlaySound Sound_ButtonClick
            End If
        Case 2
            If Not picRegister.Visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.Visible = False
                picLogin.Visible = False
                picRegister.Visible = True
                picCharacter.Visible = False
                picMain.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
            End If
        Case 3
            If Not picCredits.Visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.Visible = True
                picLogin.Visible = False
                picRegister.Visible = False
                picCharacter.Visible = False
                picMain.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
            End If
        Case 4
            Call DestroyGame
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Menu index
    
    ' change the button we're hovering on
    changeButtonState_Menu index, 2 ' clicked
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseDown", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Menu index
    
    ' change the button we're hovering on
    If Not MenuButton(index).State = 2 Then ' make sure we're not clicking
        changeButtonState_Menu index, 1 ' hover
    End If
    
    ' play sound
    If Not LastButtonSound_Menu = index Then
        PlaySound Sound_ButtonHover
        LastButtonSound_Menu = index
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
        
    ' reset all buttons
    resetButtons_Menu -1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseUp", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblLAccept_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If isLoginLegal(txtLUser.text, txtLPass.text) Then
        Call MenuState(MENU_STATE_LOGIN)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblLAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblSprite_Click()
Dim spritecount As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If optMale.value Then
        spritecount = UBound(Class(cmbClass.ListIndex + 1).MaleSprite)
    Else
        spritecount = UBound(Class(cmbClass.ListIndex + 1).FemaleSprite)
    End If

    If newCharSprite >= spritecount Then
        newCharSprite = 0
    Else
        newCharSprite = newCharSprite + 1
    End If
    
    NewCharacterBltSprite
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblSprite_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optFemale_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    NewCharacterBltSprite
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optFemale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMale_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    NewCharacterBltSprite
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picCredits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCredits_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If ServerStatus.Caption = "Connecting.." Or ServerStatus.Caption = "Online!" Then
        Exit Sub
    End If
    ServerStatus.Caption = "Connecting.."
    ServerStatus.ForeColor = RGB(250, 200, 100)
        
    resetButtons_Menu
    
    If ConnectToServer(1) Then
        ServerStatus.Caption = "Online!"
        ServerStatus.ForeColor = RGB(0, 250, 0)
    Else
        ServerStatus.Caption = "...Offline"
        ServerStatus.ForeColor = RGB(250, 0, 0)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picLogin_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picMain_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picRegister_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picRegister_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ServerStatus_Change()
'ServerStatus.Caption = GetTranslation(ServerStatus.Caption)
End Sub

Private Sub txtLPass_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    lblLAccept_Click
    KeyAscii = 0
    Exit Sub
End If

End Sub

' Register
Private Sub txtRAccept_Click()
    Dim Name As String
    Dim Password As String
    Dim PasswordAgain As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Name = Trim$(txtRUser.text)
    Password = Trim$(txtRPass.text)
    PasswordAgain = Trim$(txtRPass2.text)

    If isLoginLegal(Name, Password) Then
        If Password <> PasswordAgain Then
            Call MsgBox("Passwords don't match.")
            Exit Sub
        End If

        If Not isStringLegal(Name) Then
            Exit Sub
        End If

        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtRAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' New Char
Private Sub lblCAccept_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MenuState(MENU_STATE_ADDCHAR)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
