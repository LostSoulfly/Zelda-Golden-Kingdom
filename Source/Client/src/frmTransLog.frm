VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTransLog 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Translation"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckTranslate 
      Left            =   840
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   6969
   End
   Begin VB.Timer tmrLog 
      Interval        =   3000
      Left            =   240
      Top             =   480
   End
   Begin RichTextLib.RichTextBox txtLog 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmTransLog.frx":0000
   End
End
Attribute VB_Name = "frmTransLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private onlyOnce As Boolean

Private Sub Form_Load()
sckTranslate.Connect
Sleep 100
DoEvents
End Sub

Private Sub Form_Resize()
txtLog.Width = Me.Width - 140
txtLog.Height = Me.Height - 485
End Sub

Private Sub sckTranslate_Connect()
    If onlyOnce = False Then
        frmTransLog.Visible = True
        frmTransLog.Show
        DoEvents
    End If
    sckTranslate.SendData "15" + vbCrLf
End Sub

Private Sub sckTranslate_DataArrival(ByVal bytesTotal As Long)
Static Buffer As String
Dim NewData As String
Dim Msgs() As String
Dim MD5 As String
Dim i As Integer

sckTranslate.GetData NewData
Msgs = Split(Buffer & NewData, vbNewLine)
Buffer = Msgs(UBound(Msgs))
For i = 0 To UBound(Msgs) - 1
If Mid(Msgs(i), 1, 2) = 99 Then
    modTranslate.AddTransLog Mid(Msgs(i), 3)
Else
    MD5 = Mid(Msgs(i), 1, 32)
    Msgs(i) = Mid(Msgs(i), 33)
    Msgs(i) = Replace(Msgs(i), "\r", vbCr)
    Msgs(i) = Replace(Msgs(i), "\n", vbLf)
    Msgs(i) = Replace(Msgs(i), "\r\n", vbNewLine)
    modTranslate.AddToCache MD5, Msgs(i), modTranslate.transCol
End If
    'MsgBox Msgs(I)
Next
End Sub

Private Sub tmrLog_Timer()
If Len(txtLog.text) >= 100000 Then txtLog.text = vbNullString
If sckTranslate.State <> sckConnected Then
    sckTranslate.Close
    sckTranslate.Connect
End If

End Sub

Private Sub txtLog_Change()
With txtLog
    .SelStart = 2000000000
End With
End Sub
