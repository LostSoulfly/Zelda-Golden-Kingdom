VERSION 5.00
Begin VB.Form frmVideo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Video"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Aceptar Ajustes"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CheckBox chkActivateVideo 
      Caption         =   "Activar Grabación de video"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
   End
   Begin VB.HScrollBar scrlVideoSeconds 
      Height          =   255
      Left            =   240
      Max             =   10
      Min             =   1
      TabIndex        =   0
      Top             =   1320
      Value           =   1
      Width           =   3495
   End
   Begin VB.Label lblInfo3 
      Caption         =   "Si Guardas los ajustes la grabación actual se eliminará. Puedes Guardarla antes"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label lblInfo2 
      Caption         =   "Para Guardar el video presiona la tecla ""Fin"" o ""End"""
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label lblInfo 
      Caption         =   "Atención: La grabación de video utiliza mucha memoria RAM: Uso Medio: 28 MB"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblVideoSeconds 
      Caption         =   "Segundos Memorizados:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewFrames As Long
Dim Activate As Boolean

Private Sub cmdAccept_Click()
    ClearVideo
    Maxframes = NewFrames
    RecordingActived = Activate
    Unload Me
End Sub

Private Sub form_load()
    chkActivateVideo.Value = BTI(RecordingActived)
    scrlVideoSeconds.Value = Maxframes \ 4
    
End Sub

Private Sub chkActivateVideo_Click()
    If ITB(chkActivateVideo.Value) Then
        Activate = True
    Else
        Activate = False
    End If
End Sub



Private Sub scrlVideoSeconds_Change()
    Dim seconds As Long
    seconds = scrlVideoSeconds.Value
    If seconds > 0 And seconds <= 10 Then
        lblVideoSeconds.Caption = "Segundos Memorizados: " & seconds
        NewFrames = seconds * 4
    End If
End Sub
