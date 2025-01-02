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
      Caption         =   "Accept Settings"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CheckBox chkActivateVideo 
      Caption         =   "Activate Video Recording"
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
      Caption         =   "If you save the settings the current recording will be deleted. You can save it before."
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label lblInfo2 
      Caption         =   "To save the video press the key ""Order"" o ""End"""
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label lblInfo 
      Caption         =   "Attention: Video recording uses a lot of RAM memory: Average Usage: 28 MB"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblVideoSeconds 
      Caption         =   "Memorized Seconds:"
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

Private Sub Form_Load()
    chkActivateVideo.value = BTI(RecordingActived)
    scrlVideoSeconds.value = Maxframes \ 4
    
        Dim e As Control
    
    For Each e In Me.Controls
        If (TypeOf e Is Label) Then
            'e.caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is CheckBox) Then
            'e.caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is OptionButton) Then
            'e.caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is Frame) Then
            'e.caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is CommandButton) Then
            'e.caption = GetTranslation(e.Caption)
        End If
        If (TypeOf e Is TextBox) Then
            'e.text = GetTranslation(e.text)
        End If
    Next
    
    'me.caption = GetTranslation(Me.Caption)
    
    
End Sub

Private Sub chkActivateVideo_Click()
    If ITB(chkActivateVideo.value) Then
        Activate = True
    Else
        Activate = False
    End If
End Sub



Private Sub scrlVideoSeconds_Change()
    Dim seconds As Long
    seconds = scrlVideoSeconds.value
    If seconds > 0 And seconds <= 10 Then
        lblVideoSeconds.Caption = "Memorized Seconds:" & seconds
        NewFrames = seconds * 4
    End If
End Sub
