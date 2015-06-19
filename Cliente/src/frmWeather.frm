VERSION 5.00
Begin VB.Form frmWeather 
   Caption         =   "Weather"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAccept 
      Caption         =   "Accept"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Frame FrmSettings 
      Caption         =   "Loop Settings"
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.Frame FrmCustom 
         Caption         =   "Custom"
         Height          =   1095
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   4575
         Begin VB.HScrollBar ScrlProb 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   10
            Top             =   240
            Width           =   1695
         End
         Begin VB.HScrollBar ScrlTime 
            Height          =   255
            Left            =   1080
            Max             =   1440
            TabIndex        =   9
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton OptOther 
            Caption         =   "Other"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton OptRandom 
            Caption         =   "Random"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.Label LblProb 
            Caption         =   "Prob(%): 0"
            Height          =   255
            Left            =   3120
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label LblTime 
            Caption         =   "Time(minutes): 0"
            Height          =   375
            Left            =   3120
            TabIndex        =   6
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.OptionButton OptCustom 
         Caption         =   "Custom Lapse of time"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2775
      End
      Begin VB.OptionButton OptDisable 
         Caption         =   "Disable Weather"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton OptActivate 
         Caption         =   "Permanently Activate"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmWeather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub OptCustom_Click()
    FrmCustom.Visible = True
End Sub


Private Sub ScrlTime_Change()
    LblTime.Caption = "Time(minutes): " & ScrlTime.Value
End Sub

Private Sub scrlTime_Scroll()
    ScrlTime_Change
End Sub

Private Sub ScrlProb_Change()
    LblProb.Caption = "Prob(%): " & ScrlProb.Value
End Sub

Private Sub scrlProb_Scroll()
    ScrlProb_Change
End Sub


Private Sub CmdAccept_Click()

    ' First look for what option was choosen
    If OptActivate.Value = True Then
        WeatherTime = 0
        Call ActivateWeather
    ElseIf OptDisable = True Then
        WeatherTime = 0
        Call DisableWeather
    ElseIf OptCustom = True Then
        ' Custom, now guess if it was random or time choose
        If OptRandom.Value = True Then
            Call CalculateWeatherUpdateTime(-1, CByte(ScrlProb.Value))
        ElseIf OptOther.Value = True Then
            Call CalculateWeatherUpdateTime(Int(ScrlTime.Value), 0)
        End If
    End If
    
    frmWeather.Visible = False
    

End Sub
