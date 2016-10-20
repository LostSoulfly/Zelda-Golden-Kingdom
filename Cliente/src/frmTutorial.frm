VERSION 5.00
Begin VB.Form frmTutorial 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tutorial"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTutorial 
      BackColor       =   &H000C0E0F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   6960
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7020
      Begin VB.Label lblTutorialText 
         BackStyle       =   0  'Transparent
         Caption         =   "¡Bienvenido a ""The Legend of Zelda: El Reino Dorado""!"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   5895
      End
      Begin VB.Label lblTutorialText 
         BackStyle       =   0  'Transparent
         Caption         =   "Pulsando en el botón ""Opc"" (Opciones) tienes a tu disposición un botón para abrir el Panel de Mini-Manual."
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
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   6735
      End
      Begin VB.Label lblTutorialText 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTutorial.frx":0000
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
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   6735
      End
      Begin VB.Label lblTutorialText 
         BackStyle       =   0  'Transparent
         Caption         =   "Recomendación de pimeros pasos para:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   6615
      End
      Begin VB.Label lblTutorialText 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrenar en el Gran Árbol Deku, al este de Ciudad Kokiri, por el camino que indica Mido."
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
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   5655
      End
      Begin VB.Label lblTutorialText 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrenar en la caverna Dodongo, dentro de Ciudad Goron, hacia el norte tomando la cueva de la derecha arriba."
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
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   5655
      End
      Begin VB.Label lblTutorialText 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrenar en el interior del Gran Jabu Jabu, hacia el norte dentro de la Ciudad Zora, a la izquierda del Rey Zora."
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
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   3720
         Width           =   5655
      End
      Begin VB.Label lblTutorialText 
         BackStyle       =   0  'Transparent
         Caption         =   "Hylians"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblTutorialText 
         BackStyle       =   0  'Transparent
         Caption         =   "Gorons"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lblTutorialText 
         BackStyle       =   0  'Transparent
         Caption         =   "Zoras"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   4
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lblTutorialText 
         BackStyle       =   0  'Transparent
         Caption         =   "Gerudos"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   3
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lblTutorialText 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrenar en la Zona de Entrenamiento Gerudo, al este de la Fortaleza Gerudo, hay una cueva."
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
         Height          =   495
         Index           =   11
         Left            =   120
         TabIndex        =   2
         Top             =   4440
         Width           =   5655
      End
      Begin VB.Label lblTutorialExit 
         BackStyle       =   0  'Transparent
         Caption         =   "Cerrar Ventana"
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
         Left            =   5280
         TabIndex        =   1
         Top             =   4800
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim e As Control
    
    For Each e In Me.Controls
        If (TypeOf e Is Label) Then
            e.Caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is CheckBox) Then
            e.Caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is OptionButton) Then
            e.Caption = GetTranslation(e.Caption)
        End If
    Next
End Sub

Private Sub lblTutorialExit_Click()
Me.Visible = False
frmMain.txtMyChat.Locked = False
End Sub
