VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12915
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Spell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   861
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spell Properties"
      Height          =   7335
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   9495
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   6840
         Width           =   1215
      End
      Begin VB.TextBox txtDesc 
         Height          =   975
         Left            =   1440
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Top             =   6240
         Width           =   5295
      End
      Begin VB.Frame Frame6 
         Caption         =   "Data"
         Height          =   5895
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Width           =   5895
         Begin VB.Frame fraStates 
            Caption         =   "Change State"
            Height          =   855
            Left            =   3360
            TabIndex        =   66
            Top             =   2160
            Width           =   2055
            Begin VB.HScrollBar scrlState 
               Height          =   255
               Left            =   120
               Max             =   2
               Min             =   1
               TabIndex        =   67
               Top             =   480
               Value           =   1
               Width           =   1815
            End
            Begin VB.Label lblState 
               Caption         =   "State: ""None"""
               Height          =   255
               Left            =   120
               TabIndex        =   68
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.ComboBox cmbStatDefense 
            Height          =   300
            Left            =   3360
            TabIndex        =   65
            Text            =   "Defense Stat"
            Top             =   1680
            Width           =   1815
         End
         Begin VB.ComboBox cmbStatDamage 
            Height          =   300
            Left            =   3360
            TabIndex        =   64
            Text            =   "Damage Stat"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Frame fraStun 
            Caption         =   "StunOptions"
            Height          =   1335
            Left            =   3360
            TabIndex        =   59
            Top             =   4440
            Width           =   2175
            Begin VB.CheckBox chkPlayerAction 
               Caption         =   "Enabled"
               Height          =   180
               Left            =   240
               TabIndex        =   61
               Top             =   960
               Width           =   1575
            End
            Begin VB.HScrollBar scrlPlayerActions 
               Height          =   255
               Left            =   240
               Max             =   5
               Min             =   1
               TabIndex        =   60
               Top             =   600
               Value           =   1
               Width           =   1695
            End
            Begin VB.Label lblPlayerAction 
               Caption         =   "Action: None"
               Height          =   255
               Left            =   240
               TabIndex        =   62
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Left            =   3240
            TabIndex        =   57
            Top             =   480
            Width           =   1935
         End
         Begin VB.HScrollBar scrlStun 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   51
            Top             =   5520
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnim 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   49
            Top             =   4920
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnimCast 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   47
            Top             =   4320
            Width           =   2895
         End
         Begin VB.CheckBox chkAOE 
            Caption         =   "Area of Effect spell?"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   3240
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAOE 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   42
            Top             =   3720
            Width           =   3015
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   40
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlInterval 
            Height          =   255
            Left            =   1680
            Max             =   60
            TabIndex        =   38
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlDuration 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   36
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlVital 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   30000
            TabIndex        =   34
            Top             =   1680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlDir 
            Height          =   255
            Left            =   1680
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   1680
            Max             =   100
            TabIndex        =   20
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   18
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   200
            TabIndex        =   16
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblStat 
            Caption         =   "Buffer Stat: None"
            Height          =   255
            Left            =   3240
            TabIndex        =   58
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblStun 
            Caption         =   "Stun Duration: None"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   5280
            Width           =   2895
         End
         Begin VB.Label lblAnim 
            Caption         =   "Animation: None"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   4680
            Width           =   2895
         End
         Begin VB.Label lblAnimCast 
            Caption         =   "Cast Anim: None"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   4080
            Width           =   2895
         End
         Begin VB.Label lblAOE 
            Caption         =   "AoE: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   3480
            Width           =   3015
         End
         Begin VB.Label lblRange 
            Caption         =   "Range: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   2640
            Width           =   3015
         End
         Begin VB.Label lblInterval 
            Caption         =   "Interval: 0s"
            Height          =   255
            Left            =   1680
            TabIndex        =   37
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblDuration 
            Caption         =   "Duration: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblVital 
            Caption         =   "Vital: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label lblDir 
            Caption         =   "Dir: Down"
            Height          =   255
            Left            =   1680
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Basic Information"
         Height          =   5895
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3255
         Begin VB.CheckBox cmdPercent 
            Caption         =   "%?"
            Height          =   255
            Left            =   2640
            TabIndex        =   63
            Top             =   1680
            Width           =   615
         End
         Begin VB.PictureBox picSprite 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2640
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   50
            Top             =   5160
            Width           =   480
         End
         Begin VB.HScrollBar scrlIcon 
            Height          =   255
            Left            =   120
            Max             =   200
            TabIndex        =   45
            Top             =   5400
            Width           =   2415
         End
         Begin VB.HScrollBar scrlCool 
            Height          =   255
            Left            =   120
            Max             =   180
            TabIndex        =   32
            Top             =   4680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlCast 
            Height          =   255
            Left            =   120
            Max             =   200
            TabIndex        =   30
            Top             =   4080
            Width           =   3015
         End
         Begin VB.ComboBox cmbClass 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   3480
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAccess 
            Height          =   255
            Left            =   120
            Max             =   5
            TabIndex        =   26
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlLevel 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   100
            TabIndex        =   24
            Top             =   2280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMP 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   2000
            TabIndex        =   13
            Top             =   1680
            Width           =   2415
         End
         Begin VB.ComboBox cmbType 
            Height          =   300
            ItemData        =   "frmEditor_Spell.frx":0A4E
            Left            =   120
            List            =   "frmEditor_Spell.frx":0A6A
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblIcon 
            Caption         =   "Icon: None"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   5160
            Width           =   3015
         End
         Begin VB.Label lblCool 
            Caption         =   "Cooldown Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   4440
            Width           =   2535
         End
         Begin VB.Label lblCast 
            Caption         =   "Casting Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   3840
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Class Required:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label lblAccess 
            Caption         =   "Access Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label lblLevel 
            Caption         =   "Level Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblMP 
            Caption         =   "MP Cost: None"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   6240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Spell List"
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6900
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   7560
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAOE_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkAOE.value = 0 Then
        Spell(EditorIndex).IsAoE = False
    Else
        Spell(EditorIndex).IsAoE = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkAOE_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkPlayerAction_Click()
    Spell(EditorIndex).BlockActions(scrlPlayerActions.value) = chkPlayerAction.value
End Sub

Private Sub cmbClass_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Spell(EditorIndex).ClassReq = cmbClass.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClass_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Spell(EditorIndex).Type = cmbType.ListIndex
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearSpell EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    SpellEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdPercent_Click()
    Spell(EditorIndex).UsePercent = cmdPercent.value
    Call CheckMPScroll(Spell(EditorIndex).UsePercent)
End Sub



Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccess_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAccess.value > 0 Then
        lblAccess.Caption = "Access Required: " & scrlAccess.value
    Else
        lblAccess.Caption = "Access Required: None"
    End If
    Spell(EditorIndex).AccessReq = scrlAccess.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccess_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnim.value > 0 Then
        lblAnim.Caption = "Animation: " & Trim$(Animation(scrlAnim.value).Name)
    Else
        lblAnim.Caption = "Animation: None"
    End If
    Spell(EditorIndex).SpellAnim = scrlAnim.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnimCast.value > 0 Then
        lblAnimCast.Caption = "Cast Anim: " & Trim$(Animation(scrlAnimCast.value).Name)
    Else
        lblAnimCast.Caption = "Cast Anim: None"
    End If
    Spell(EditorIndex).CastAnim = scrlAnimCast.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAOE_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAOE.value > 0 Then
        lblAOE.Caption = "AoE: " & scrlAOE.value & " tiles."
    Else
        lblAOE.Caption = "AoE: Self-cast"
    End If
    Spell(EditorIndex).AoE = scrlAOE.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAOE_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblCast.Caption = "Casting Time: " & scrlCast.value & "s"
    Spell(EditorIndex).CastTime = scrlCast.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCool_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblCool.Caption = "Cooldown Time: " & scrlCool.value & "s"
    Spell(EditorIndex).CDTime = scrlCool.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCool_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDir_Change()
Dim sDir As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlDir.value
        Case DIR_UP
            sDir = "Up"
        Case DIR_DOWN
            sDir = "Down"
        Case DIR_RIGHT
            sDir = "Right"
        Case DIR_LEFT
            sDir = "Left"
    End Select
    lblDir.Caption = "Dir: " & sDir
    Spell(EditorIndex).dir = scrlDir.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDir_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDuration_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblDuration.Caption = "Duration: " & scrlDuration.value & "s"
    Spell(EditorIndex).Duration = scrlDuration.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDuration_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlIcon_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlIcon.value > 0 Then
        lblIcon.Caption = "Icon: " & scrlIcon.value
    Else
        lblIcon.Caption = "Icon: None"
    End If
    Spell(EditorIndex).Icon = scrlIcon.value
    EditorSpell_BltIcon
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlIcon_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlInterval_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblInterval.Caption = "Interval: " & scrlInterval.value & "s"
    Spell(EditorIndex).Interval = scrlInterval.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlInterval_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlLevel.value > 0 Then
        lblLevel.Caption = "Level Required: " & scrlLevel.value
    Else
        lblLevel.Caption = "Level Required: None"
    End If
    Spell(EditorIndex).LevelReq = scrlLevel.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevel_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMap_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblMap.Caption = "Map: " & scrlMap.value
    Spell(EditorIndex).map = scrlMap.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMap_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlMP.value > 0 Then
        lblMP.Caption = "MP Cost: " & scrlMP.value
    Else
        lblMP.Caption = "MP Cost: None"
    End If
    Spell(EditorIndex).MPCost = scrlMP.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMP_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPlayerActions_Change()
    If 0 <= Spell(EditorIndex).BlockActions(scrlPlayerActions.value) <= 1 Then
        chkPlayerAction.value = BTI(Spell(EditorIndex).BlockActions(scrlPlayerActions.value))
        lblPlayerAction.Caption = "Action: " & ActionToStr(scrlPlayerActions.value)
    Else
        chkPlayerAction.value = False
        lblPlayerAction.Caption = "Action: " & ActionToStr(scrlPlayerActions.value)
    End If
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlRange.value > 0 Then
        lblRange.Caption = "Range: " & scrlRange.value & " tiles."
    Else
        lblRange.Caption = "Range: Self-cast"
    End If
    Spell(EditorIndex).range = scrlRange.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStat_Change()
    Spell(EditorIndex).stat = scrlStat.value
    lblStat.Caption = "Buffer Stat: " & StatToStr(scrlStat.value)
End Sub

Public Sub UpdateChangeStateInfo(ByVal value As Byte)
    lblState.Caption = "State: " & StateToStr(value)
    Spell(EditorIndex).ChangeState = value
End Sub

Private Sub scrlState_Change()
    Call UpdateChangeStateInfo(scrlState.value)
End Sub

Public Sub InitChangeStates()
    scrlState.Min = 0
    scrlState.Max = Max_States - 1
    scrlState.value = Spell(EditorIndex).ChangeState
End Sub


Private Sub scrlStun_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlStun.value > 0 Then
        lblStun.Caption = "Stun Duration: " & scrlStun.value & "s"
    Else
        lblStun.Caption = "Stun Duration: None"
    End If
    Spell(EditorIndex).StunDuration = scrlStun.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStun_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlVital_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblVital.Caption = "Vital: " & scrlVital.value
    Spell(EditorIndex).vital = scrlVital.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlVital_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblX.Caption = "X: " & scrlX.value
    Spell(EditorIndex).X = scrlX.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlX_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblY.Caption = "Y: " & scrlY.value
    Spell(EditorIndex).y = scrlY.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlY_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Spell(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Spell(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Spell(EditorIndex).Sound = cmbSound.list(cmbSound.ListIndex)
    Else
        Spell(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbStatDamage_Click()
Spell(EditorIndex).StatDamage = cmbStatDamage.ListIndex
End Sub

Private Sub cmbStatDefense_Click()
Spell(EditorIndex).StatDefense = cmbStatDefense.ListIndex
End Sub

