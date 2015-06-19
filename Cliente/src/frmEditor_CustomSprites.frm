VERSION 5.00
Begin VB.Form frmEditor_CustomSprites 
   Caption         =   "Custom Sprites"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstIndex 
      Height          =   6690
      ItemData        =   "frmEditor_CustomSprites.frx":0000
      Left            =   120
      List            =   "frmEditor_CustomSprites.frx":0002
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Frame fraCustomSprite 
      Caption         =   "Custom Sprite"
      Height          =   6855
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.Frame fraLayers 
         Caption         =   "Layers"
         Height          =   4935
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   8175
         Begin VB.CommandButton cmdSetMySprite 
            Caption         =   "Set My Sprite"
            Height          =   615
            Left            =   6840
            TabIndex        =   27
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton cmdPasteLayer 
            Caption         =   "Paste"
            Height          =   255
            Left            =   3240
            TabIndex        =   26
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdCopyLayer 
            Caption         =   "Copy"
            Height          =   255
            Left            =   3240
            TabIndex        =   25
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Frame fraLayerSettings 
            Caption         =   "Layer Settings"
            Height          =   1095
            Left            =   4800
            TabIndex        =   22
            Top             =   1200
            Width           =   1935
            Begin VB.CheckBox chkCenter 
               Caption         =   "Fix Center Position"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   720
               Width           =   1695
            End
            Begin VB.CheckBox chkPlayerSprite 
               Caption         =   "Use Player Sprite"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.CommandButton cmdDeleteLayer 
            Caption         =   "Delete Layer"
            Height          =   615
            Left            =   1800
            TabIndex        =   21
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddLayer 
            Caption         =   "Add Layer"
            Height          =   615
            Left            =   480
            TabIndex        =   20
            Top             =   1560
            Width           =   1095
         End
         Begin VB.HScrollBar scrlLayers 
            Height          =   255
            Left            =   480
            Max             =   1
            Min             =   1
            TabIndex        =   18
            Top             =   1080
            Value           =   1
            Width           =   3975
         End
         Begin VB.Frame fraPos 
            Caption         =   "Fix Position"
            Height          =   2055
            Left            =   2880
            TabIndex        =   11
            Top             =   2760
            Visible         =   0   'False
            Width           =   5175
            Begin VB.Frame fraAnims 
               Caption         =   "Enabled Anims"
               Height          =   1815
               Left            =   2880
               TabIndex        =   29
               Top             =   120
               Width           =   2175
               Begin VB.HScrollBar scrlAnimto 
                  Height          =   255
                  Left            =   240
                  Max             =   3
                  TabIndex        =   33
                  Top             =   1080
                  Width           =   1695
               End
               Begin VB.HScrollBar scrlAnims 
                  Height          =   255
                  Left            =   240
                  Max             =   3
                  TabIndex        =   31
                  Top             =   480
                  Width           =   1695
               End
               Begin VB.CommandButton cmdEnableAll 
                  Caption         =   "Default"
                  Height          =   255
                  Left            =   720
                  TabIndex        =   30
                  Top             =   1440
                  Width           =   735
               End
               Begin VB.Label lblSetDir 
                  Caption         =   "Set: 0"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   34
                  Top             =   840
                  Width           =   975
               End
               Begin VB.Label lblAnims 
                  Caption         =   "Anim: 0"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   32
                  Top             =   240
                  Width           =   615
               End
            End
            Begin VB.CommandButton cmdCenterLayer 
               Caption         =   "Center"
               Height          =   375
               Left            =   1800
               TabIndex        =   28
               Top             =   1200
               Width           =   735
            End
            Begin VB.TextBox txtY 
               Height          =   285
               Left            =   360
               TabIndex        =   14
               Top             =   1560
               Width           =   975
            End
            Begin VB.TextBox txtX 
               Height          =   285
               Left            =   360
               TabIndex        =   13
               Top             =   1200
               Width           =   975
            End
            Begin VB.HScrollBar scrlDir 
               Height          =   255
               Left            =   360
               Max             =   3
               TabIndex        =   12
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label lblY 
               Caption         =   "Y:"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   1560
               Width           =   255
            End
            Begin VB.Label lblX 
               Caption         =   "X:"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   1200
               Width           =   255
            End
            Begin VB.Label lblDir 
               Caption         =   "Dir: 0"
               Height          =   255
               Left            =   360
               TabIndex        =   15
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.Frame fraSprite 
            Caption         =   "Sprite"
            Height          =   1815
            Left            =   120
            TabIndex        =   8
            Top             =   3000
            Width           =   2655
            Begin VB.HScrollBar scrlSprite 
               Height          =   255
               Left            =   240
               Max             =   500
               TabIndex        =   9
               Top             =   960
               Value           =   1
               Width           =   2175
            End
            Begin VB.Label lblSprite 
               Caption         =   "Sprite: 0"
               Height          =   255
               Left            =   240
               TabIndex        =   10
               Top             =   600
               Width           =   1455
            End
         End
         Begin VB.Label lblLayer 
            Caption         =   "Layer: 1"
            Height          =   255
            Left            =   480
            TabIndex        =   19
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   6120
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   2160
         TabIndex        =   2
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   3720
         TabIndex        =   1
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmEditor_CustomSprites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CopiedLayerNum As Byte

Private Sub chkCenter_Click()
    If Not IsLayerLegal Then Exit Sub
    Dim value As Boolean
    value = chkCenter.value
    CustomSprites(EditorIndex).Layers(GetLayerNum).UseCenterPosition = value
    If value Then
        fraPos.Visible = True
    Else
        fraPos.Visible = False
    End If
End Sub

Private Sub cmdCenterLayer_Click()
    If Not IsLayerLegal Then Exit Sub
    
    Dim sprite As Long
    If IsLayerUsingPlayerSprite(CustomSprites(EditorIndex).Layers(GetLayerNum)) Then
        sprite = GetPlayerSprite(MyIndex)
    Else
        sprite = GetLayerSprite(CustomSprites(EditorIndex).Layers(GetLayerNum))
    End If
    
    With DDSD_Character(sprite)
        Dim Dir As Byte
        Dim spritetop As Byte
        For Dir = 0 To MAX_DIRECTIONS - 1
            Select Case Dir
                Case DIR_UP
                    spritetop = 3
                Case DIR_DOWN
                    spritetop = 0
                Case DIR_LEFT
                    spritetop = 1
                Case DIR_RIGHT
                    spritetop = 2
            End Select
            
            scrlDir.value = Dir
            
            txtX.text = CStr((.lWidth / 4) / 2)
            Call txtX_Validate(False)
            txtY.text = CStr((.lHeight / 4) / 2 + spritetop * (.lHeight / 4))
            Call txtY_Validate(False)
        Next
    
    End With
End Sub

Private Sub cmdEnableAll_Click()
    Dim i As Byte, j As Byte
    If IsLayerLegal Then
        For j = 0 To MAX_DIRECTIONS - 1
            For i = 0 To MAX_SPRITE_ANIMS - 1
                CustomSprites(EditorIndex).Layers(GetLayerNum).fixed.EnabledAnims(j, i) = i
            Next
        Next
    End If
    
    scrlAnims.value = 0
    
End Sub

Private Sub chkEnableAnim_Click()
    
End Sub


Private Sub chkPlayerSprite_Click()
    If Not IsLayerLegal Then Exit Sub
    Dim value As Boolean
    value = chkPlayerSprite.value
    CustomSprites(EditorIndex).Layers(GetLayerNum).UsePlayerSprite = value
    If value Then
        fraSprite.Visible = False
    Else
        fraSprite.Visible = True
    End If
End Sub

Private Sub cmdAddLayer_Click()
    Call AddLayer(CustomSprites(EditorIndex), GetLayerNum)
    scrlLayers.Max = scrlLayers.Max + 1
End Sub

Private Sub cmdCopyLayer_Click()
    If IsLayerLegal Then
        CopiedLayerNum = scrlLayers.value
    End If
End Sub

Private Sub cmdDeleteLayer_Click()
    If IsLayerLegal Then
    If GetCustomSpriteNLayers(CustomSprites(EditorIndex)) > 1 Then
        Call DeleteLayer(CustomSprites(EditorIndex), GetLayerNum)
        scrlLayers.Max = scrlLayers.Max - 1
    End If
    Call UpdateInfo
    End If

End Sub



Private Sub cmdPasteLayer_Click()
    If Not IsLayerLegal Then Exit Sub
    Call CopyLayer(scrlLayers.value, CopiedLayerNum)
End Sub

Private Sub cmdSetMySprite_Click()
    'Player(MyIndex).CustomSprite = EditorIndex
    Call SetPlayerCustomSprite(MyIndex, EditorIndex)
End Sub





Private Sub scrlAnims_Change()
    lblAnims.Caption = "Anim: " & scrlAnims.value
    
    If Not IsLayerLegal Then Exit Sub
    
    Call LoadAnim(scrlDir.value, scrlAnims.value)
    'chkEnableAnim.Value = BTI(CustomSprites(EditorIndex).Layers(GetLayerNum).fixed.EnabledAnims(scrlAnims.Value))
End Sub

Private Sub scrlAnimto_Change()
    If Not IsLayerLegal Then Exit Sub
    Dim value As Byte
    value = scrlAnimto.value
    SetAnim value
End Sub

Public Sub LoadAnim(ByVal Dir As Byte, ByVal Anim As Byte)
    scrlAnimto.value = CustomSprites(EditorIndex).Layers(GetLayerNum).fixed.EnabledAnims(Dir, Anim)
End Sub

Public Sub SetAnim(ByVal value As Byte)
    If value < 0 Or value >= MAX_SPRITE_ANIMS Then Exit Sub
    
    lblSetDir.Caption = "Set: " & value
    CustomSprites(EditorIndex).Layers(GetLayerNum).fixed.EnabledAnims(scrlDir.value, scrlAnims.value) = value
End Sub

Private Sub scrlDir_Change()
    lblDir.Caption = "Dir: " & DirtoStr(scrlDir.value)
    If Not IsLayerLegal Then Exit Sub
    
    With CustomSprites(EditorIndex).Layers(GetLayerNum)
        txtX.text = CStr(.CentersPositions(scrlDir.value).X)
        txtY.text = CStr(.CentersPositions(scrlDir.value).Y)
    End With
    
    scrlAnims.value = 0
End Sub

Private Sub scrlLayers_Change()
    Call UpdateInfo
    lblLayer.Caption = "Layer: " & scrlLayers.value
    scrlDir.value = 0
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = "Sprite: " & scrlSprite.value
    If IsLayerLegal Then
        CustomSprites(EditorIndex).Layers(GetLayerNum).sprite = scrlSprite.value
    End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    CustomSprites(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & CustomSprites(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    CustomSpritesEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
       
    Call CustomSpritesEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearCustomSprite EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & CustomSprites(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    CustomSpritesEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call CustomSpritesEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtX_Validate(Cancel As Boolean)
If Not IsLayerLegal Then Exit Sub

If IsNumeric(txtX.text) Then
    If Len(txtX.text) <= MAX_INTEGER Then
        CustomSprites(EditorIndex).Layers(GetLayerNum).CentersPositions(scrlDir.value).X = CInt(Trim$(txtX.text))
    End If
End If

End Sub

Private Sub txtY_Validate(Cancel As Boolean)
If Not IsLayerLegal Then Exit Sub

If IsNumeric(txtX.text) Then
    If Len(txtX.text) <= MAX_INTEGER Then
        CustomSprites(EditorIndex).Layers(GetLayerNum).CentersPositions(scrlDir.value).Y = CInt(txtY.text)
    End If
End If

End Sub

Private Function IsLayerLegal() As Boolean
    If scrlLayers.value > 0 Then
        If scrlLayers.value <= CustomSprites(EditorIndex).NLayers Then
            IsLayerLegal = True
        End If
    End If
End Function

Private Function GetLayerNum() As Byte
    GetLayerNum = scrlLayers.value
End Function

Private Sub CopyLayer(ByVal layer As Byte, ByVal layer0 As Byte)
    If layer > 0 And layer <= CustomSprites(EditorIndex).NLayers And layer0 > 0 And layer0 <= CustomSprites(EditorIndex).NLayers Then
        CustomSprites(EditorIndex).Layers(layer) = CustomSprites(EditorIndex).Layers(layer0)
        Call UpdateInfo
    End If
End Sub

Public Sub UpdateInfo()
    If Not IsLayerLegal Then Exit Sub
    chkPlayerSprite.value = BTI(IsLayerUsingPlayerSprite(CustomSprites(EditorIndex).Layers(GetLayerNum)))
    chkCenter.value = BTI(IsLayerUsingCenter(CustomSprites(EditorIndex).Layers(GetLayerNum)))
    If GetLayerSprite(CustomSprites(EditorIndex).Layers(GetLayerNum)) <= scrlSprite.Max Then
        scrlSprite.value = GetLayerSprite(CustomSprites(EditorIndex).Layers(GetLayerNum))
    End If
    scrlDir.value = 0
    scrlAnims.value = 0
    
    txtX.text = CStr(GetLayerCenterX(CustomSprites(EditorIndex).Layers(GetLayerNum), scrlDir.value))
    txtY.text = CStr(GetLayerCenterY(CustomSprites(EditorIndex).Layers(GetLayerNum), scrlDir.value))

End Sub


