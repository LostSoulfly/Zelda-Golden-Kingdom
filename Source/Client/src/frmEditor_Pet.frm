VERSION 5.00
Begin VB.Form frmEditor_Pets 
   Caption         =   "Pets Editor"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar scrlTame 
      Height          =   255
      LargeChange     =   10
      Left            =   2760
      TabIndex        =   15
      Top             =   4200
      Width           =   4815
   End
   Begin VB.HScrollBar scrlExp 
      Height          =   255
      LargeChange     =   10
      Left            =   2760
      Max             =   15
      TabIndex        =   13
      Top             =   3480
      Width           =   4815
   End
   Begin VB.HScrollBar scrlPoints 
      Height          =   255
      LargeChange     =   10
      Left            =   2760
      Max             =   15
      TabIndex        =   11
      Top             =   2760
      Width           =   4815
   End
   Begin VB.HScrollBar scrlMaxLevel 
      Height          =   255
      LargeChange     =   10
      Left            =   2760
      Max             =   100
      TabIndex        =   8
      Top             =   2040
      Width           =   4815
   End
   Begin VB.HScrollBar scrlNpcNum 
      Height          =   255
      LargeChange     =   10
      Left            =   2760
      Max             =   400
      TabIndex        =   7
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   5520
      Width           =   1455
   End
   Begin VB.ListBox lstIndex 
      Height          =   5910
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblTame 
      Caption         =   "TamePoints: 0"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lblExp 
      Caption         =   "Exp Progression: 0"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblPoints 
      Caption         =   "Points Progression: Normal"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label lblMaxLvl 
      Caption         =   "Max Lvl: 0"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblNpcNum 
      Caption         =   "NPC: 0"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmEditor_Pets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub scrlExp_Change()
    Pet(EditorIndex).ExpProgression = scrlExp.value
    lblEXP.Caption = "Exp Progression: " & scrlExp.value
End Sub

Private Sub scrlMaxLevel_Change()
    Pet(EditorIndex).MaxLevel = scrlMaxLevel.value
    lblMaxLvl.Caption = "Max Lvl: " & scrlMaxLevel.value
End Sub

Private Sub scrlNpcNum_Change()
    Pet(EditorIndex).NPCNum = scrlNpcNum.value
    If scrlNpcNum.value <> 0 Then
        lblNpcNum.Caption = "NPC: " & scrlNpcNum.value & ", " & NPC(scrlNpcNum.value).Name
    Else
        lblNpcNum.Caption = "NPC: 0"
    End If
End Sub

Private Sub scrlPoints_Change()
    Pet(EditorIndex).PointsProgression = scrlPoints.value
    lblPoints.Caption = "Points Progression: " & GetPointsProgressionString(scrlPoints.value)
End Sub

Public Function GetPointsProgressionString(ByVal points As Byte) As String
Dim sum As Double
Dim N As Byte

If points < 0 Or points > MAX_PET_POINTS_PERLVL Then Exit Function

sum = 0
N = 0

Do While sum < points
    sum = sum + CDbl(MAX_PET_POINTS_PERLVL / 5)
    N = N + 1
Loop

Select Case N
Case 1
    GetPointsProgressionString = "Very Low"
Case 2
    GetPointsProgressionString = "Low"
Case 3
    GetPointsProgressionString = "Normal"
Case 4
    GetPointsProgressionString = "High"
Case 5
    GetPointsProgressionString = "Very High"
Case Else
    GetPointsProgressionString = "Constant"
End Select


End Function

Private Sub scrlTame_Change()
    Pet(EditorIndex).TamePoints = scrlTame.value
    lblTame.Caption = "TamePoints: " & scrlTame.value
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Pet(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Pet(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call PetEditorCancel
    
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
    
    ClearPet EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Pet(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    PetsEditorInit
    
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
       
    Call PetsEditorOk
    
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
    
    PetsEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
