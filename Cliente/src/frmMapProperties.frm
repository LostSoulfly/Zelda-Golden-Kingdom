VERSION 5.00
Begin VB.Form frmEditor_MapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMapProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraAllowedStates 
      Caption         =   "Allowed States"
      Height          =   1335
      Left            =   6720
      TabIndex        =   40
      Top             =   2280
      Width           =   2055
      Begin VB.CheckBox chkAllow 
         Caption         =   "Allow"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   1575
      End
      Begin VB.HScrollBar scrlState 
         Height          =   255
         Left            =   120
         Max             =   2
         Min             =   1
         TabIndex        =   41
         Top             =   600
         Value           =   1
         Width           =   1815
      End
      Begin VB.Label lblState 
         Caption         =   "State: ""None"""
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame frmMapNpcProperties 
      Caption         =   "MapNPCProperties"
      Height          =   2055
      Left            =   6720
      TabIndex        =   33
      Top             =   120
      Width           =   2055
      Begin VB.ComboBox cmbAction 
         Height          =   315
         Left            =   240
         TabIndex        =   36
         Text            =   "Action"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox cmbMovement 
         Height          =   315
         Left            =   240
         TabIndex        =   35
         Text            =   "Movemnt."
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblnpcnum 
         Caption         =   "NPC: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.ComboBox CmbWeather 
      Height          =   315
      ItemData        =   "frmMapProperties.frx":0A4E
      Left            =   120
      List            =   "frmMapProperties.frx":0A5E
      TabIndex        =   32
      Text            =   "Weather"
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Music"
      Height          =   2775
      Left            =   4440
      TabIndex        =   27
      Top             =   1440
      Width           =   2055
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   1815
      End
      Begin VB.ListBox lstMusic 
         Height          =   1815
         ItemData        =   "frmMapProperties.frx":0A7F
         Left            =   120
         List            =   "frmMapProperties.frx":0A81
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frmMaxSizes 
      Caption         =   "Max Sizes"
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   2055
      Begin VB.TextBox txtMaxX 
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtMaxY 
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Max X:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Max Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   630
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Links"
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   2055
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   20
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   19
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   18
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblMap 
         BackStyle       =   0  'Transparent
         Caption         =   "Current map: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Map Settings"
      Height          =   855
      Left            =   2280
      TabIndex        =   13
      Top             =   480
      Width           =   4215
      Begin VB.ComboBox cmbMoral 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0A83
         Left            =   960
         List            =   "frmMapProperties.frx":0A96
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moral:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boot Settings"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
      Begin VB.TextBox txtBootMap 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Boot Map:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Boot X:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Boot Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   630
      End
   End
   Begin VB.Frame fraNPCs 
      Caption         =   "NPCs"
      Height          =   3615
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
      Begin VB.CommandButton cmdPasteNPC 
         Caption         =   "P"
         Height          =   255
         Left            =   1560
         TabIndex        =   39
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton cmdCopyNPC 
         Caption         =   "C"
         Height          =   255
         Left            =   1080
         TabIndex        =   38
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton cmdClearNPC 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   3240
         Width           =   735
      End
      Begin VB.ListBox lstNpcs 
         Height          =   2400
         ItemData        =   "frmMapProperties.frx":0AC4
         Left            =   120
         List            =   "frmMapProperties.frx":0AC6
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2760
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEditor_MapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private NPCCopied As Long





Private Sub cmbAction_Click()

    If Not lstNpcs.ListCount > 0 Then Exit Sub
    
    If Actions(cmbAction.ListIndex + 1).Name <> vbNullString Then
        
    With map
    .NPCSProperties(lstNpcs.ListIndex + 1).Action = cmbAction.ListIndex
    End With
    
    End If

End Sub
Private Sub cmbMovement_Click()

    If Not lstNpcs.ListCount > 0 Then Exit Sub
    
    If Movements(cmbMovement.ListIndex + 1).Name <> vbNullString Then
        
    With map
    .NPCSProperties(lstNpcs.ListIndex + 1).movement = cmbMovement.ListIndex
    End With
    
    End If
End Sub



Private Sub cmdClearNPC_Click()
If lstNpcs.ListIndex >= 0 Then

    

    map.NPC(lstNpcs.ListIndex + 1) = 0
    Dim X As Long
    Dim tmpIndex As Long
    ' re-load the list
    tmpIndex = lstNpcs.ListIndex
    lstNpcs.Clear
    For X = 1 To MAX_MAP_NPCS
        If map.NPC(X) > 0 Then
        lstNpcs.AddItem X & ": " & Trim$(NPC(map.NPC(X)).Name)
        Else
            lstNpcs.AddItem X & ": No NPC"
        End If
    Next
    lstNpcs.ListIndex = tmpIndex
    
End If
End Sub

Private Sub cmdCopyNPC_Click()
If lstNpcs.ListIndex >= 0 Then
    NPCCopied = map.NPC(lstNpcs.ListIndex + 1)
End If

End Sub

Private Sub cmdPasteNPC_Click()
If lstNpcs.ListIndex >= 0 Then
    If NPCCopied > 0 And NPCCopied < MAX_NPCS Then
        map.NPC(lstNpcs.ListIndex + 1) = NPCCopied
        Dim X As Long
        Dim tmpIndex As Long
        ' re-load the list
        tmpIndex = lstNpcs.ListIndex
        lstNpcs.Clear
        For X = 1 To MAX_MAP_NPCS
            If map.NPC(X) > 0 Then
            lstNpcs.AddItem X & ": " & Trim$(NPC(map.NPC(X)).Name)
            Else
                lstNpcs.AddItem X & ": No NPC"
            End If
        Next
        lstNpcs.ListIndex = tmpIndex
    End If
End If
End Sub

Private Sub cmdPlay_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    StopMidi
    PlayMidi lstMusic.list(lstMusic.ListIndex)
    PlayMp3 lstMusic.list(lstMusic.ListIndex)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdPlay_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdStop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    StopMidi
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdStop_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOk_Click()
    Dim i As Long
    Dim sTemp As Long
    Dim X As Long, x2 As Long
    Dim Y As Long, y2 As Long
    Dim tempArr() As TileRec
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not IsNumeric(txtMaxX.text) Then txtMaxX.text = map.MaxX
    If Val(txtMaxX.text) < MAX_MAPX Then txtMaxX.text = MAX_MAPX
    If Val(txtMaxX.text) > MAX_BYTE Then txtMaxX.text = MAX_BYTE
    If Not IsNumeric(txtMaxY.text) Then txtMaxY.text = map.MaxY
    If Val(txtMaxY.text) < MAX_MAPY Then txtMaxY.text = MAX_MAPY
    If Val(txtMaxY.text) > MAX_BYTE Then txtMaxY.text = MAX_BYTE

    With map
        .Name = Trim$(txtName.text)
        If lstMusic.ListIndex >= 0 Then
            .Music = lstMusic.list(lstMusic.ListIndex)
        Else
            .Music = vbNullString
        End If
        .Up = Val(txtUp.text)
        .Down = Val(txtDown.text)
        .Left = Val(txtLeft.text)
        .Right = Val(txtRight.text)
        .moral = cmbMoral.ListIndex
        .BootMap = Val(txtBootMap.text)
        .BootX = Val(txtBootX.text)
        .BootY = Val(txtBootY.text)
        .Weather = CmbWeather.ListIndex

        ' set the data before changing it
        tempArr = map.Tile
        x2 = map.MaxX
        y2 = map.MaxY
        ' change the data
        .MaxX = Val(txtMaxX.text)
        .MaxY = Val(txtMaxY.text)
        ReDim map.Tile(0 To .MaxX, 0 To .MaxY)

        If x2 > .MaxX Then x2 = .MaxX
        If y2 > .MaxY Then y2 = .MaxY

        For X = 0 To x2
            For Y = 0 To y2
                .Tile(X, Y) = tempArr(X, Y)
            Next
        Next

        ClearTempTile
    End With

    Call UpdateDrawMapName
    Unload frmEditor_MapProperties
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdOk_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Unload frmEditor_MapProperties
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub lstNpcs_Click()
Dim tmpString() As String
Dim NPCNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' exit out if needed
    If Not cmbNpc.ListCount > 0 Then Exit Sub
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    
    ' set the combo box properly
    tmpString = Split(lstNpcs.list(lstNpcs.ListIndex))
    NPCNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
    cmbNpc.ListIndex = map.NPC(NPCNum)
    
    'Actions and movements
    lblNpcNum.Caption = "NPC: " & NPCNum
    
    With map
    
    If .NPCSProperties(NPCNum).movement > 0 Then
        'cmbMovement.ListIndex = 0
        cmbMovement.ListIndex = .NPCSProperties(NPCNum).movement
    Else
        cmbMovement.text = "No movement"
    End If
    
    If .NPCSProperties(NPCNum).Action > 0 Then
        'cmbAction.ListIndex = 0
        cmbAction.ListIndex = .NPCSProperties(NPCNum).Action
    Else
        cmbAction.text = "No Action"
    End If
    
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstNpcs_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbNpc_Click()
Dim tmpString() As String
Dim NPCNum As Long
Dim X As Long, tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' exit out if needed
    If Not cmbNpc.ListCount > 0 Then Exit Sub
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    
    ' set the combo box properly
    tmpString = Split(cmbNpc.list(cmbNpc.ListIndex))
    ' make sure it's not a clear
    If Not cmbNpc.list(cmbNpc.ListIndex) = "No NPC" Then
        NPCNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        map.NPC(lstNpcs.ListIndex + 1) = NPCNum
    Else
        map.NPC(lstNpcs.ListIndex + 1) = 0
    End If
    
    ' re-load the list
    tmpIndex = lstNpcs.ListIndex
    lstNpcs.Clear
    For X = 1 To MAX_MAP_NPCS
        If map.NPC(X) > 0 Then
        lstNpcs.AddItem X & ": " & Trim$(NPC(map.NPC(X)).Name)
        Else
            lstNpcs.AddItem X & ": No NPC"
        End If
    Next
    lstNpcs.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbNpc_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub UpdateAllowedStateInfo(ByVal value As Byte)
    lblState.Caption = "State: " & StateToStr(value)
    chkAllow.value = BTI(map.AllowedStates(value))
End Sub

Private Sub scrlState_Change()
    Call UpdateAllowedStateInfo(scrlState.value)
End Sub

Public Sub InitAllowedStates()
    scrlState.Min = 1
    scrlState.Max = Max_States - 1
    scrlState.value = 1
End Sub

Private Sub chkAllow_Click()
    map.AllowedStates(scrlState.value) = ITB(chkAllow.value)
End Sub
