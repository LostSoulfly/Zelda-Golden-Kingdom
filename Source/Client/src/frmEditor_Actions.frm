VERSION 5.00
Begin VB.Form frmEditor_Actions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actions"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmAction 
      Caption         =   "Action"
      Height          =   5535
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      Begin VB.Frame frmSubVital 
         Caption         =   "Sub-Vital"
         Height          =   2175
         Left            =   360
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   4815
         Begin VB.Frame frmByDiv 
            Caption         =   "By Div"
            Height          =   735
            Left            =   2280
            TabIndex        =   27
            Top             =   1200
            Visible         =   0   'False
            Width           =   2055
            Begin VB.TextBox txtVitalAbstract 
               Height          =   285
               Left            =   1320
               TabIndex        =   29
               Text            =   "1"
               Top             =   240
               Width           =   375
            End
            Begin VB.Label lblVitalAbstract 
               Caption         =   "GetPlayerVital /"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame frmByNum 
            Caption         =   "By Num"
            Height          =   855
            Left            =   2280
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   2415
            Begin VB.HScrollBar scrlVitalNum 
               Height          =   255
               Left            =   240
               TabIndex        =   25
               Top             =   480
               Width           =   1935
            End
            Begin VB.Label lblVitalNum 
               Caption         =   "Vital Num: 0"
               Height          =   255
               Left            =   240
               TabIndex        =   26
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.CheckBox chkLoseExp 
            Caption         =   "Don't lose exp"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton optVitalAbstract 
            Caption         =   "By Division"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1680
            Width           =   1335
         End
         Begin VB.OptionButton optVitalNum 
            Caption         =   "By Num"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1320
            Width           =   1695
         End
         Begin VB.HScrollBar scrlVital 
            Height          =   255
            Left            =   240
            Max             =   2
            Min             =   1
            TabIndex        =   13
            Top             =   600
            Value           =   1
            Width           =   1935
         End
         Begin VB.Label lblVitalName 
            Caption         =   "Vital: None"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame frmWarp 
         Caption         =   "Warp"
         Height          =   2055
         Left            =   1680
         TabIndex        =   16
         Top             =   2280
         Visible         =   0   'False
         Width           =   2175
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   120
            Max             =   30
            TabIndex        =   22
            Top             =   1680
            Width           =   1935
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            Max             =   30
            TabIndex        =   21
            Top             =   1080
            Width           =   1935
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   400
            Min             =   1
            TabIndex        =   20
            Top             =   480
            Value           =   1
            Width           =   1935
         End
         Begin VB.Label lblY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblMap 
            Caption         =   "Map: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame frmMoment 
         Caption         =   "Moment"
         Height          =   975
         Left            =   2880
         TabIndex        =   9
         Top             =   960
         Width           =   2415
         Begin VB.ComboBox cmbMoment 
            Height          =   315
            ItemData        =   "frmEditor_Actions.frx":0000
            Left            =   240
            List            =   "frmEditor_Actions.frx":0010
            TabIndex        =   10
            Text            =   "Moment"
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame frmActionType 
         Caption         =   "Action Type"
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2415
         Begin VB.ComboBox cmbActionType 
            Height          =   315
            ItemData        =   "frmEditor_Actions.frx":004C
            Left            =   240
            List            =   "frmEditor_Actions.frx":0056
            TabIndex        =   8
            Text            =   "Action Type"
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   3720
         TabIndex        =   6
         Top             =   4800
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   2040
         TabIndex        =   5
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.ListBox lstIndex 
      Height          =   5520
      ItemData        =   "frmEditor_Actions.frx":006B
      Left            =   120
      List            =   "frmEditor_Actions.frx":006D
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmEditor_Actions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkLoseExp_Click()
    Actions(EditorIndex).Data4 = chkLoseExp.value '1 = don't lose exp, 0 = lose exp
End Sub

Private Sub cmbActionType_Click()

    Actions(EditorIndex).Type = cmbActionType.ListIndex
    Call ActionsShowWindow(cmbActionType.ListIndex)
End Sub

Private Sub cmbMoment_Click()

    Actions(EditorIndex).Moment = cmbMoment.ListIndex
    
End Sub


Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ActionsEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAction EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Actions(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ActionsEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
       
    Call ActionsEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionsEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optVitalAbstract_Click()
        Call ResetActionForms
        frmByDiv.Visible = True
        Call txtVitalAbstract_Change 'Reset Info
        Actions(EditorIndex).Data2 = 1
        'scrlVitalNum.Value = 0
End Sub

Private Sub optVitalNum_Click()
        Call ResetActionForms
        frmByNum.Visible = True
        Call scrlVitalNum_Change 'Reset Info
        Actions(EditorIndex).Data2 = 0
        'txtVitalAbstract.text = "0"
End Sub

Private Sub scrlMap_Change()
    Actions(EditorIndex).Data1 = scrlMap.value
    lblMap.Caption = "Map: " & scrlMap.value
    scrlX.Max = MAX_BYTE
    scrlY.Max = MAX_BYTE
End Sub

Private Sub scrlVital_Change()
'Find vital
'Change text
    Actions(EditorIndex).Data1 = scrlVital.value
    lblVitalName.Caption = "Vital " & VitalsEnumToString(scrlVital.value)


End Sub

Private Sub scrlVitalNum_Change()
    
    
    Actions(EditorIndex).Data3 = scrlVitalNum.value
    lblVitalNum.Caption = "Vital Num: " & scrlVitalNum.value

    
        
End Sub

Private Sub scrlX_Change()
    
    Actions(EditorIndex).Data2 = scrlX.value
    lblX.Caption = "X: " & scrlX.value
    
End Sub

Private Sub scrlY_Change()

    Actions(EditorIndex).Data3 = scrlY.value
    lblY.Caption = "Y: " & scrlY.value
    
End Sub

Private Sub txtVitalAbstract_Change()

    
    If IsNumeric(txtVitalAbstract.text) Then
        If txtVitalAbstract.text > 0 Then
            Actions(EditorIndex).Data3 = Val(txtVitalAbstract.text)
        End If
    End If
    
End Sub

Public Function VitalsEnumToString(ByVal EnumIndex As Long) As String

If EnumIndex > 0 And EnumIndex < Vitals.Vital_Count Then
    Select Case EnumIndex
    Case 1
        VitalsEnumToString = "HP"
    Case 2
        VitalsEnumToString = "MP"
    End Select
Else
    VitalsEnumToString = ""
End If



End Function

Private Sub ActionsShowWindow(ByVal index As Long)

    Call ClearActionTypeFrames
    Select Case index
    Case 0
        frmSubVital.Visible = True
    Case 1
        frmWarp.Visible = True
    End Select
              
End Sub

Public Sub ClearActionTypeFrames()
    frmSubVital.Visible = False
    frmWarp.Visible = False
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Actions(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Actions(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub ResetActionForms()
frmByNum.Visible = False
frmByDiv.Visible = False
End Sub



