VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Loading..."
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSWinsockLib.Winsock hubSocket 
      Left            =   0
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   5000
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Console"
      TabPicture(0)   =   "frmServer.frx":1708A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblCPS"
      Tab(0).Control(1)=   "lblCpsLock"
      Tab(0).Control(2)=   "txtText"
      Tab(0).Control(3)=   "txtChat"
      Tab(0).Control(4)=   "tmrIsServerBug"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":170A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwInfo"
      Tab(1).Control(1)=   "cmdCopy"
      Tab(1).Control(2)=   "texto"
      Tab(1).Control(3)=   "cmdDisableAdmins"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Control "
      TabPicture(2)   =   "frmServer.frx":170C2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraDatabase"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "FraExp"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdTransLog"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdSave"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "fraServer"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "FraExtras"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Info"
      TabPicture(3)   =   "frmServer.frx":170DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblBytesReceived"
      Tab(3).Control(1)=   "lblBytesSent"
      Tab(3).Control(2)=   "lblPacketsReceived"
      Tab(3).Control(3)=   "lblPacketsSent"
      Tab(3).Control(4)=   "lblLoopTime"
      Tab(3).Control(5)=   "lblMapTime"
      Tab(3).ControlCount=   6
      Begin VB.Frame FraExtras 
         Caption         =   "Extras"
         Height          =   855
         Left            =   3000
         TabIndex        =   19
         Top             =   2400
         Width           =   3375
         Begin VB.CommandButton cmdMapLinkReport 
            Caption         =   "Map link report"
            Height          =   495
            Left            =   2280
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdWeather 
            Caption         =   "Weather"
            Height          =   495
            Left            =   1200
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton CharacterEditor 
            Caption         =   "Character Editor"
            Height          =   495
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraServer 
         Caption         =   "Server"
         Height          =   2055
         Left            =   3000
         TabIndex        =   1
         Top             =   480
         Width           =   1815
         Begin VB.CheckBox chkHub 
            Caption         =   "Use Hub"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CheckBox chkTroll 
            Caption         =   "TrollMode!"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CheckBox chkPass 
            Caption         =   "Verify Passw."
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chkServerLog 
            Caption         =   "Server Log"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton cmdShutDown 
            Caption         =   "Shut Down"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save."
         Height          =   375
         Left            =   4920
         TabIndex        =   37
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdTransLog 
         Caption         =   "Show TransLog"
         Height          =   375
         Left            =   4920
         TabIndex        =   35
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Frame FraExp 
         Caption         =   "Exp: 1"
         Height          =   855
         Left            =   4920
         TabIndex        =   31
         Top             =   480
         Width           =   1455
         Begin VB.HScrollBar spExp 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   50
            Min             =   1
            TabIndex        =   32
            Top             =   360
            Value           =   1
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdDisableAdmins 
         Caption         =   "On/Off Gm"
         Height          =   250
         Left            =   -74880
         TabIndex        =   30
         Top             =   3150
         Width           =   1080
      End
      Begin VB.TextBox texto 
         Height          =   285
         Left            =   -73320
         TabIndex        =   29
         Top             =   3120
         Width           =   4455
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "loc"
         Height          =   250
         Left            =   -73750
         TabIndex        =   28
         Top             =   3150
         Width           =   360
      End
      Begin VB.Timer tmrIsServerBug 
         Interval        =   60000
         Left            =   -69000
         Top             =   0
      End
      Begin VB.Frame fraDatabase 
         Caption         =   "Reload"
         Height          =   2775
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2775
         Begin VB.CommandButton cmdLoadOptions 
            Caption         =   "Options"
            Height          =   375
            Left            =   1440
            TabIndex        =   36
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadLang 
            Caption         =   "Language"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1440
            TabIndex        =   34
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "Animations"
            Height          =   375
            Left            =   1440
            TabIndex        =   16
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "Resources"
            Height          =   375
            Left            =   1440
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Items"
            Height          =   375
            Left            =   1440
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "Npcs"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Shops"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Spells"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Maps"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "Classes"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtChat 
         Height          =   285
         Left            =   -74880
         TabIndex        =   3
         Top             =   3090
         Width           =   6255
      End
      Begin VB.TextBox txtText 
         Height          =   2175
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   840
         Width           =   6255
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4683
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label lblBytesReceived 
         Caption         =   "Bytes Received / Second: 0"
         Height          =   375
         Left            =   -74880
         TabIndex        =   27
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label lblBytesSent 
         Caption         =   "Bytes Sent / Second: 0"
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   1920
         Width           =   3375
      End
      Begin VB.Label lblPacketsReceived 
         Caption         =   "Packets Received / Second: 0"
         Height          =   255
         Left            =   -74880
         TabIndex        =   25
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label lblPacketsSent 
         Caption         =   "Packets Sent / Second: 0"
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop(ms): 0"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblMapTime 
         Caption         =   "MapUpdate(ms): 0"
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblCpsLock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[Unlock]"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74880
         TabIndex        =   18
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblCPS 
         Caption         =   "CPS: 0"
         Height          =   255
         Left            =   -74040
         TabIndex        =   17
         Top             =   645
         Width           =   1815
      End
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
      Begin VB.Menu mnuGodPlayer 
         Caption         =   "Turn GodPlayer?"
      End
      Begin VB.Menu mnuSpellPlayer 
         Caption         =   "Turn SpellPlayer?"
      End
      Begin VB.Menu mnuSail 
         Caption         =   "Set Sail!"
      End
      Begin VB.Menu mnuRide 
         Caption         =   "Start Riding."
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkHub_Click()
useHubServer = chkHub.Value
End Sub

Private Sub cmdCopy_Click()
    Dim a As String
    Dim i As Byte
    For i = 1 To 3
        a = a + lvwInfo.SelectedItem.SubItems(i) + ","
    Next
    
    i = FindPlayer(lvwInfo.SelectedItem.SubItems(3))

    a = a + "GM: " & GetPlayerAccess_Mode(i) & ", LOCALIP: " & GetPlayerHost(i)

    
    texto.Text = a
End Sub

Private Sub cmdDisableAdmins_Click()
If Options.DisableAdmins = 0 Then
    Options.DisableAdmins = 1
    cmdDisableAdmins.Caption = "On GM"
    Else
    Options.DisableAdmins = 0
    cmdDisableAdmins.Caption = "Off GM"
End If
End Sub


Private Sub cmdLoadOptions_Click()
Call TextAdd("Options reloaded.")
LoadOptions
SendHubCommand CommandsType.SOptions, ""
End Sub

Private Sub cmdMapLinkReport_Click()
Dim i As Long
On Error Resume Next
cmdMapLinkReport.Enabled = False
Kill App.Path & "\data\logs\MAP_CHECK.log"
    TextAdd "Checking maps for unlinked/unwarped maps.."
    For i = 1 To MAX_MAPS
    DoEvents
        If CheckMapUnlinked(i) = True Then
        'dead-end maps should be here. Next to check if any map warps here or links here..
            If CheckMapsConnectedTo(i) = False Then
                AddLog2 "Map " & Trim$(map(i).Name) & " (" & i & ") is not connected anywhere! (spell warps unchecked)", "MAP_CHECK.log"
            End If
        End If
    Next
    Shell "notepad.exe " & App.Path & "\data\logs\MAP_CHECK.log", vbNormalFocus
    cmdMapLinkReport.Enabled = True
End Sub

Private Sub cmdTransLog_Click()
frmTransLog.Visible = True
End Sub

Private Sub cmdWeather_Click()
frmWeather.Visible = True
End Sub

Private Sub cmdSave_Click()
Dim i As Long
'get
SaveClasses
SaveItems
SaveResources
SaveNpcs
SaveShops
SaveSpells
SaveAnimations
SaveQuests
SaveDoors
SaveActions
Savemovements
SavePets
SaveCustomSprites

MsgBox "Done."

End Sub

Private Sub Command1_Click()
SendHubCommand CommandsType.SOptions, ""
End Sub

Private Sub hubSocket_DataArrival(ByVal bytesTotal As Long)
    Call HubIncomingData(bytesTotal)
End Sub

Private Sub hubSocket_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    If Number <> 0 And Number <> 10061 Then
        TextAdd "Hub Server disconnected: " & Number & ": " & Description
    Else
        If isHubConnected = True Then
            TextAdd "Hub Server disconnected normally."
            'todo global message about disconnect.
            hubSocket.Close
            isHubConnected = False
        End If
    End If

End Sub

Private Sub lblCPSLock_Click()
    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Unlock]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Lock]"
    End If
End Sub

Private Sub CharacterEditor_Click()
'If GetAD Then Exit Sub
frmAccountEditor.Visible = True
End Sub

Private Sub mnuGodPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        TurnGodPlayer FindPlayer(Name)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You are now a GodPlayer. Wtf does that mean?", BrightCyan, True)
    End If

End Sub

Private Sub mnuSail_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        If GetPlayerState(FindPlayer(Name)) = StateSailing Then
            ClearSailing (FindPlayer(Name))
        Else
            StartSailing (FindPlayer(Name))
        End If
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "Allll abbbooooooooarddddd!", BrightCyan, True)
    End If

End Sub

Private Sub mnuRide_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        If GetPlayerState(FindPlayer(Name)) = StateRiding Then
            ClearRiding (FindPlayer(Name))
        Else
            StartRiding (FindPlayer(Name))
        End If
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "Allll abbbooooooooarddddd!", BrightCyan, True)
    End If

End Sub

Private Sub mnuSpellPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        TurnSpellPlayer FindPlayer(Name)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You are now a SpellPlayer. Wtf does that mean?", BrightCyan, True)
    End If

End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)

    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If

End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub

' ********************
Private Sub chkServerLog_Click()

    ' if its not 0, then its true
    If Not chkServerLog.Value Then
        ServerLog = True
    End If

End Sub

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmdReloadClasses_Click()
Dim i As Long
    Call LoadClasses
    Call TextAdd("All classes reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendClasses i
        End If
    Next
    
End Sub

Private Sub cmdReloadItems_Click()
Dim i As Long
    Call LoadItems
    Call TextAdd("All items reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendItems i
        End If
    Next

End Sub

Private Sub cmdReloadMaps_Click()
Dim i As Long

    Call LoadMaps
    Call TextAdd("All maps reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            PlayerSpawn i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i)
        End If
    Next

End Sub

Private Sub cmdReloadNPCs_Click()
Dim i As Long
    Call LoadNpcs
    Call TextAdd("All npcs reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendNpcs i
        End If
    Next

End Sub

Private Sub cmdReloadShops_Click()
Dim i As Long
    Call LoadShops
    Call TextAdd("All shops reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendShops i
        End If
    Next

End Sub

Private Sub cmdReloadSpells_Click()
Dim i As Long
    Call LoadSpells
    Call TextAdd("All spells reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendSpells i
        End If
    Next

End Sub

Private Sub cmdReloadResources_Click()
Dim i As Long
    Call LoadResources
    Call TextAdd("All Resources reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendResources i
        End If
    Next

End Sub

Private Sub cmdReloadAnimations_Click()
Dim i As Long
    Call LoadAnimations
    Call TextAdd("All Animations reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendAnimations i
        End If
    Next

End Sub

Private Sub cmdShutDown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown cancelled.", Cyan
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancel"
    End If
End Sub

Private Sub Form_Load()
    Call UsersOnline_Start
End Sub

Private Sub Form_Resize()

    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader - 1
    'Set Sorted to True to sort the list.
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.index - 1
    lvwInfo.Sorted = True
End Sub

Private Sub spExp_Change()

    Options.ExpMultiplier = spExp.Value
    FraExp.Caption = "Exp: " & Options.ExpMultiplier
    
    Call SaveOptions
    SendHubCommand CommandsType.SOptions, ""
End Sub

Private Sub tmrIsServerBug_Timer()
    IsServerBug = True
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg("Server: " & txtChat.Text, White, False)
            Call TextAdd("Server: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If

End Sub

Sub UsersOnline_Start()
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (i)

        If i < 10 Then
            frmServer.lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmServer.lvwInfo.ListItems(i).Text = "0" & i
        Else
            frmServer.lvwInfo.ListItems(i).Text = i
        End If

        frmServer.lvwInfo.ListItems(i).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(3) = vbNullString
    Next

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If

End Sub

Private Sub mnuKickPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" And Not LPE(FindPlayer(Name)) Then
        Call AlertMsg(FindPlayer(Name), "You have been kicked by the server.")
    End If

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" And Not LPE(FindPlayer(Name)) Then
        CloseSocket (FindPlayer(Name))
    End If

End Sub

Sub mnuBanPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" And Not LPE(FindPlayer(Name)) Then
        Call ServerBanIndex(FindPlayer(Name))
    End If

End Sub

Sub mnuAdminPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    
    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 4)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been promoted to administrator.", BrightCyan)
    End If

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" And Not LPE(FindPlayer(Name)) Then
        Call SetPlayerAccess(FindPlayer(Name), 0)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been revoked from your Administrator position", BrightRed)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lmsg As Long
    lmsg = X / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.Text)
    End Select

End Sub

