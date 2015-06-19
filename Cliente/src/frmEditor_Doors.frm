VERSION 5.00
Begin VB.Form frmEditor_Doors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Door Editor"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11085
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5760
      TabIndex        =   26
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   25
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4200
      TabIndex        =   24
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Door/Switch Properties"
      Height          =   5775
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Frame fraDoor 
         Caption         =   "Door/Switch Properties"
         Height          =   4455
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   7335
         Begin VB.Frame fraInitialState 
            Caption         =   "InitialState"
            Height          =   735
            Left            =   4560
            TabIndex        =   34
            Top             =   1680
            Width           =   2175
            Begin VB.CheckBox chkInitialState 
               Caption         =   "Firstly Opened"
               Height          =   255
               Left            =   240
               TabIndex        =   35
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.Frame fraTime 
            Caption         =   "Lock Time"
            Height          =   1215
            Left            =   4800
            TabIndex        =   30
            Top             =   2880
            Width           =   2295
            Begin VB.HScrollBar scrlTime 
               Height          =   255
               Left            =   240
               Max             =   60
               TabIndex        =   32
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label lblTime 
               Caption         =   "time: 0"
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Unlock What door?"
            Height          =   855
            Left            =   3600
            TabIndex        =   27
            Top             =   480
            Width           =   3135
            Begin VB.HScrollBar scrlSwitch 
               Height          =   255
               Left            =   240
               TabIndex        =   28
               Top             =   480
               Width           =   2415
            End
            Begin VB.Label lblSwitch 
               Caption         =   "Door: None"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   2895
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Warp"
            Height          =   1935
            Left            =   240
            TabIndex        =   16
            Top             =   480
            Width           =   2175
            Begin VB.HScrollBar scrlY 
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   1560
               Width           =   1935
            End
            Begin VB.HScrollBar scrlX 
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   1080
               Width           =   1935
            End
            Begin VB.HScrollBar scrlMap 
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label lblY 
               Caption         =   "Map y: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lblX 
               Caption         =   "Map x: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   840
               Width           =   855
            End
            Begin VB.Label lblMap 
               Caption         =   "Map: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.Frame fraToUnlock 
            Caption         =   "How To Unlock"
            Height          =   1455
            Left            =   240
            TabIndex        =   9
            Top             =   2760
            Width           =   4215
            Begin VB.OptionButton OptUnlock 
               Caption         =   "None"
               Height          =   255
               Index           =   2
               Left            =   3120
               TabIndex        =   15
               Top             =   240
               Width           =   975
            End
            Begin VB.Frame Frame4 
               Caption         =   "Key"
               Height          =   855
               Left            =   120
               TabIndex        =   12
               Top             =   480
               Width           =   3375
               Begin VB.HScrollBar scrlKey 
                  Height          =   255
                  Left            =   240
                  TabIndex        =   13
                  Top             =   480
                  Width           =   2415
               End
               Begin VB.Label lblKey 
                  Caption         =   "Key: None"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   14
                  Top             =   240
                  Width           =   3015
               End
            End
            Begin VB.OptionButton OptUnlock 
               Caption         =   "Switch"
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   11
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton OptUnlock 
               Caption         =   "Key"
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   10
               Top             =   240
               Width           =   975
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Door or switch?"
         Height          =   735
         Left            =   3600
         TabIndex        =   3
         Top             =   240
         Width           =   3615
         Begin VB.OptionButton optDoor 
            Caption         =   "WSwitch"
            Height          =   195
            Index           =   2
            Left            =   2280
            TabIndex        =   33
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optDoor 
            Caption         =   "Switch"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optDoor 
            Caption         =   "Door"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Door/Switch List"
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.ListBox lstIndex 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5325
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Doors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkInitialState_Click()
    Doors(EditorIndex).InitialState = chkInitialState.Value
End Sub

Private Sub Form_Load()

scrlMap.Max = MAX_MAPS
scrlX.Max = MAX_BYTE
scrlY.Max = MAX_BYTE
scrlSwitch.Max = MAX_DOORS
scrlKey.Max = MAX_ITEMS
End Sub

Private Sub cmdCancel_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call DoorEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_DOORS Then Exit Sub
    
    ClearDoor EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Doors(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    DoorEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call DoorEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    DoorEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optDoor_Click(Index As Integer)
Doors(EditorIndex).DoorType = Index
If Index = 0 Then
    Frame6.Visible = True
    fraToUnlock.Visible = True
    Frame5.Visible = False
Else
    Frame6.Visible = False
    fraToUnlock.Visible = False
    Frame5.Visible = True
End If
End Sub

Private Sub OptUnlock_Click(Index As Integer)
Doors(EditorIndex).UnlockType = Index
If Index = 0 Then
    Frame4.Visible = True
Else
    Frame4.Visible = False
End If
End Sub

Private Sub scrlKey_Change()
If scrlKey.Value > 0 Then
lblKey.Caption = "Key: " & Trim$(Item(scrlKey.Value).Name)
Else
lblKey.Caption = "Key: None"
End If
Doors(EditorIndex).key = scrlKey.Value
End Sub



Private Sub scrlMap_Change()

lblMap.Caption = "Map: " & scrlMap.Value
Doors(EditorIndex).WarpMap = scrlMap.Value

End Sub

Private Sub scrlSwitch_Change()
If (scrlSwitch.Value > 0) Then
lblSwitch.Caption = "Door: " & Trim$(Doors(scrlSwitch.Value).Name)
Else
lblSwitch.Caption = "Door: None"
End If
Doors(EditorIndex).Switch = scrlSwitch.Value
End Sub



Private Sub scrlTime_Change()
lblTime.Caption = "Time: " & scrlTime.Value
Doors(EditorIndex).Time = scrlTime.Value
End Sub

Private Sub scrlX_Change()
lblX.Caption = "Map x: " & scrlX.Value
Doors(EditorIndex).WarpX = scrlX.Value
End Sub

Private Sub scrlY_Change()
lblY.Caption = "Map y: " & scrlY.Value
Doors(EditorIndex).WarpY = scrlY.Value
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Doors(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Doors(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
