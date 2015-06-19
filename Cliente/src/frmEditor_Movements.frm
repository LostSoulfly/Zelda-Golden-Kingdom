VERSION 5.00
Begin VB.Form frmEditor_Movements 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movements"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmMovement 
      Caption         =   "Movement Editor"
      Height          =   6375
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3240
         TabIndex        =   25
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CommandButton cmdErase 
         Caption         =   "Delete"
         Height          =   375
         Left            =   1800
         TabIndex        =   24
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   5760
         Width           =   1215
      End
      Begin VB.ComboBox cmbMovementType 
         Height          =   315
         ItemData        =   "frmEditor_Movements.frx":0000
         Left            =   840
         List            =   "frmEditor_Movements.frx":000A
         TabIndex        =   15
         Text            =   "Type"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtMovementName 
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   3615
      End
      Begin VB.Frame frmCustom 
         Caption         =   "Custom"
         Height          =   2895
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   4335
         Begin VB.Frame fraRepeat 
            Caption         =   "Repeat"
            Height          =   615
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   1335
            Begin VB.CheckBox chkRepeat 
               Caption         =   "Repeat"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdAddtoList 
            Caption         =   "Add to list"
            Height          =   615
            Left            =   1800
            TabIndex        =   18
            Top             =   2040
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cmbCustomType 
            Height          =   315
            ItemData        =   "frmEditor_Movements.frx":0027
            Left            =   240
            List            =   "frmEditor_Movements.frx":0034
            TabIndex        =   17
            Text            =   "Custom Type"
            Top             =   480
            Width           =   1455
         End
         Begin VB.Frame frmDirection 
            Caption         =   "Direction"
            Height          =   975
            Left            =   120
            TabIndex        =   11
            Top             =   1800
            Visible         =   0   'False
            Width           =   1575
            Begin VB.ComboBox cmbDirection 
               Height          =   315
               ItemData        =   "frmEditor_Movements.frx":0056
               Left            =   120
               List            =   "frmEditor_Movements.frx":0069
               TabIndex        =   12
               Text            =   "Direction"
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame frmNumTiles 
            Caption         =   "Num Tiles"
            Height          =   975
            Left            =   2520
            TabIndex        =   8
            Top             =   1800
            Visible         =   0   'False
            Width           =   1695
            Begin VB.HScrollBar scrlNumTiles 
               Height          =   255
               Left            =   120
               Max             =   255
               TabIndex        =   21
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label lblNumTilesInfo 
               Caption         =   "Val: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame frmNumMovements 
            Caption         =   "Num Movements"
            Height          =   1695
            Left            =   1920
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   2295
            Begin VB.CommandButton cmdDelete 
               Caption         =   "Delete"
               Height          =   255
               Left            =   1440
               TabIndex        =   14
               Top             =   1320
               Width           =   735
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "Add"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   1320
               Width           =   735
            End
            Begin VB.HScrollBar scrlMovements 
               Height          =   270
               Left            =   120
               Max             =   0
               TabIndex        =   7
               Top             =   960
               Width           =   2055
            End
            Begin VB.Label lblDirection 
               Caption         =   "Dir: None"
               Height          =   255
               Left            =   1320
               TabIndex        =   20
               Top             =   600
               Width           =   855
            End
            Begin VB.Label lblNumTiles 
               Caption         =   "Tiles: 0"
               Height          =   255
               Left            =   1320
               TabIndex        =   19
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblNumMovement 
               Caption         =   "Movement: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   360
               Width           =   1095
            End
         End
      End
      Begin VB.Frame frmOnlyDirectional 
         Caption         =   "OnlyDirectional"
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
         Begin VB.ComboBox cmbOnlyDirectionDir 
            Height          =   315
            ItemData        =   "frmEditor_Movements.frx":008A
            Left            =   240
            List            =   "frmEditor_Movements.frx":009D
            TabIndex        =   16
            Text            =   "Starting Dir"
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame frmMovements 
      Caption         =   "Movements"
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.ListBox lstIndex 
         Height          =   5715
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmEditor_Movements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkRepeat_Click()
    Movements(EditorIndex).Repeat = chkRepeat.Value
End Sub

Private Sub cmbCustomType_Click()

Select Case cmbCustomType.ListIndex
    
    Case 0
        Movements(EditorIndex).Type = ByDirection
        frmEditor_Movements.frmNumTiles.Visible = False
        frmEditor_Movements.frmNumMovements.Visible = True
        frmEditor_Movements.frmDirection.Visible = True
        frmEditor_Movements.cmdAddtoList.Visible = True
        frmEditor_Movements.lblNumTiles.Visible = False
    Case 1
        Movements(EditorIndex).Type = Bytile
        frmEditor_Movements.frmNumTiles.Visible = True
        frmEditor_Movements.frmNumMovements.Visible = True
        frmEditor_Movements.frmDirection.Visible = True
        frmEditor_Movements.cmdAddtoList.Visible = True
        frmEditor_Movements.lblNumTiles.Visible = True
    Case 2
        Movements(EditorIndex).Type = Random
        frmEditor_Movements.frmNumTiles.Visible = False
        frmEditor_Movements.frmNumMovements.Visible = False
        frmEditor_Movements.frmDirection.Visible = False
        frmEditor_Movements.cmdAddtoList.Visible = False
        
End Select

End Sub

Private Sub cmbMovementType_Click()

    Select Case cmbMovementType.ListIndex
    
    Case 0
        Call LMCreate(Movements(EditorIndex).MovementsTable)
        scrlMovements.Min = 0
        scrlMovements.Max = 0
        Movements(EditorIndex).Type = Onlydirectional
        frmEditor_Movements.frmOnlyDirectional.Visible = True
        frmEditor_Movements.frmCustom.Visible = False
    Case 1
        Call LMCreate(Movements(EditorIndex).MovementsTable)
        scrlMovements.Min = 0
        scrlMovements.Max = 0
        frmEditor_Movements.frmOnlyDirectional.Visible = False
        frmEditor_Movements.frmCustom.Visible = True
    End Select
    
End Sub

Private Sub cmbOnlyDirectionDir_Click()
    Call LMCreate(Movements(EditorIndex).MovementsTable)
    Call InsertListElement(frmEditor_Movements.cmbOnlyDirectionDir.ListIndex)
End Sub

Private Sub cmdAdd_Click()

Call InsertListElement(4) 'Insert null element

scrlMovements.Max = scrlMovements.Max + 1
scrlMovements.Min = 1

End Sub

Private Sub cmdAddtoList_Click()
    If Movements(EditorIndex).MovementsTable.nelem > 0 Then
    Call ParseElement(cmbDirection.ListIndex, scrlNumTiles.Value)
    Call ActualizeInfoByActual
    End If
End Sub

Private Sub cmdDelete_Click()
's'elimina l'element actual, i es desplaça el punt d'interés al següent element, si existeix
'actualitza totes les dades del següent element
If scrlMovements.Max > 1 Then
    Call LMDelete(Movements(EditorIndex).MovementsTable)
    scrlMovements.Max = scrlMovements.Max - 1
    Call ActualizeInfoByActual
End If

End Sub


Private Sub lstIndex_Click()

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MovementsEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub scrlMovements_Change()

lblNumMovement.Caption = "Movement: " & scrlMovements.Value

Dim found As Boolean
found = False
Call LMFirst(Movements(EditorIndex).MovementsTable)
Do While Not (LMEnd(Movements(EditorIndex).MovementsTable)) And Not (found)
    If Movements(EditorIndex).MovementsTable.Actual = scrlMovements.Value Then
        found = True
    Else
        Call LMNext(Movements(EditorIndex).MovementsTable)
    End If
Loop

If found Then
    Call ActualizeInfoByActual
Else
    Call resetinfo
End If
's'avança i es retrocedeix en la llista mitjançant la crida ja definida
'actualizta totes les dades del següent element
End Sub


Private Sub scrlNumTiles_Change()
        lblNumTilesInfo.Caption = "Val: " & scrlNumTiles.Value
End Sub

Private Sub txtMovementName_Change()
'es guarda en el moviment actual (editorindex) el nom desitjat
's'actualitza la barra d'indexs
End Sub

Private Sub ParseElement(Optional ByVal direction As Integer = 4, Optional ByVal numtiles As Byte = 0)

If cmbDirection.ListIndex = -1 Then
    direction = 4
End If

If LMempty(Movements(EditorIndex).MovementsTable) = True Then
    Call InsertListElement(CByte(direction), numtiles)
Else
    Call ModifyListElement(CByte(direction), numtiles)
End If
        
End Sub

Private Sub ActualizeInfoByActual()
Dim auxiliar As SingularMovementRec
auxiliar = LMGet(Movements(EditorIndex).MovementsTable)

frmEditor_Movements.lblDirection.Caption = "Dir: " & DirtoStr(auxiliar.direction)
frmEditor_Movements.lblNumTiles.Caption = "Tiles: " & auxiliar.NumberOfTiles

End Sub

Private Sub InsertListElement(Optional ByVal direction As Byte = 4, Optional ByVal tiles As Byte = 0)
Dim auxiliar As SingularMovementRec

auxiliar.direction = direction
auxiliar.NumberOfTiles = tiles

Call LMAdd(Movements(EditorIndex).MovementsTable, auxiliar)

End Sub

Private Sub ModifyListElement(Optional ByVal direction As Byte = 4, Optional ByVal tiles As Byte = 0)
Dim auxiliar As SingularMovementRec

auxiliar.direction = direction
auxiliar.NumberOfTiles = tiles

Call LMModify(Movements(EditorIndex).MovementsTable, auxiliar)

End Sub

Private Sub resetinfo()

frmEditor_Movements.lblDirection.Caption = "Dir: None"
frmEditor_Movements.lblNumTiles.Caption = "Tiles: 0"

End Sub

Private Sub txtNumTiles_Change()
If IsNumeric(txtNumTiles.text) Then
    If txtNumTiles.text > 0 And txtNumTiles <= 255 Then
        'Call ParseElement(cmbDirection.ListIndex, txtNumTiles.text)
        lblNumTiles.Caption = "Tiles: " & txtNumTiles.text
    Else
        txtNumTiles.text = 0
        lblNumTiles.Caption = "Tiles: 0"
    End If
End If

End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    
    
    Call MovementsEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MovementsEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdErase_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearMovement EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Movements(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    MovementsEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMovementName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Movements(EditorIndex).Name = Trim$(txtMovementName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Movements(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


