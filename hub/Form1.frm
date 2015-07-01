VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hub Server"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrStats 
      Interval        =   1000
      Left            =   5520
      Top             =   4800
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   6000
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Log"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtLog"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tmrLog"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Config"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkChat"
      Tab(1).Control(1)=   "cmdShutdown"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Stats"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblBytesReceived"
      Tab(2).Control(1)=   "lblBytesSent"
      Tab(2).Control(2)=   "lblPacketsSent"
      Tab(2).Control(3)=   "lblPacketsReceived"
      Tab(2).ControlCount=   4
      Begin VB.CheckBox chkChat 
         Caption         =   "Global Hub Chat"
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         Top             =   840
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton cmdShutdown 
         Caption         =   "Shutdown Servers"
         Height          =   255
         Left            =   -74760
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.Timer tmrLog 
         Interval        =   60000
         Left            =   5040
         Top             =   4680
      End
      Begin RichTextLib.RichTextBox txtLog 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   7858
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":0054
      End
      Begin VB.Label lblBytesReceived 
         Caption         =   "Packets Received / Second: 1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   5
         Top             =   2160
         Width           =   4095
      End
      Begin VB.Label lblBytesSent 
         Caption         =   "Packets Received / Second: 1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   4
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label lblPacketsSent 
         Caption         =   "Packets Received / Second: 1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   3
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label lblPacketsReceived 
         Caption         =   "Packets Received / Second: 1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdShutdown_Click()
    SendDataToAllHub BuildGeneric(HShutdown, "")
    AddLog "Sending shutdown command to all servers."
End Sub

Private Sub Form_Load()
test = GetRealTickCount
End Sub

Private Sub Socket_Close(Index As Integer)
    Call CloseSocket(Index)
End Sub

Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Call AcceptConnection(Index, requestID)
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Call IncomingData(Index, bytesTotal)
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)
    Call AcceptConnection(Index, SocketId)
End Sub

Private Sub Socket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call CloseSocket(Index)
End Sub

Private Sub tmrLog_Timer()
    If Len(txtLog.text) >= 600000 Then txtLog.text = vbNullString
End Sub

Private Sub tmrStats_Timer()
    UpdateTrafficStadistics
End Sub

Private Sub txtLog_Change()
With txtLog
    .SelStart = 2000000000
End With
End Sub
