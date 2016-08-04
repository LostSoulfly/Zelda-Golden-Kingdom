VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Server Select"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Play on this server!"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   3015
   End
   Begin VB.ListBox lstServers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSelect_Click()
Dim index As Integer
    
    index = lstServers.ListIndex
    
    'With Server(index)
    '    MsgBox .Name & " " & .CurrentPlayers, vbOKOnly, "test"
    'End With
    
    SelectedServer = index + 1
    If CheckServerFull(SelectedServer) = True Then
        If MsgBox("This server seems to be full! Are you sure you want to try to play on it?", vbYesNo, "Server Full!") = vbYes Then
            frmSelect.Visible = False
            Exit Sub
        End If
    Else
        frmSelect.Visible = False
    End If
    
    frmMain.tmrServerStatus.Enabled = False
    frmMain.tmrServerStatus.Enabled = True
    
End Sub

Private Sub Form_Load()
On Error Resume Next
If SelectedServer = 0 Then SelectedServer = 1

    lstServers.ListIndex = SelectedServer + 1
    
End Sub

Private Sub lstServers_Click()
'MsgBox Server(lstServers.ListIndex + 1).Name
End Sub
