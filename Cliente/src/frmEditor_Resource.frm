VERSION 5.00
Begin VB.Form frmEditor_Resource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Editor"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12255
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
   Icon            =   "frmEditor_Resource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   817
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   28
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   27
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resource Properties"
      Height          =   7575
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   8175
      Begin VB.Frame fraWalkable 
         Caption         =   "Walkable"
         Height          =   1215
         Left            =   5160
         TabIndex        =   34
         Top             =   1920
         Width           =   1335
         Begin VB.CheckBox chkWalkableExhausted 
            Caption         =   "Exhausted"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chkWalkableNormal 
            Caption         =   "Normal"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame frmRewards 
         Caption         =   "Rewards"
         Height          =   3615
         Left            =   5160
         TabIndex        =   33
         Top             =   3600
         Width           =   1935
         Begin VB.ComboBox cmbRewardType 
            Height          =   300
            ItemData        =   "frmEditor_Resource.frx":0A4E
            Left            =   120
            List            =   "frmEditor_Resource.frx":0A58
            TabIndex        =   43
            Text            =   "Type"
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CheckBox chkMessage 
            Caption         =   "Message: item name?"
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CheckBox chkEqual 
            Caption         =   "Equal %?"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   2400
            Width           =   1095
         End
         Begin VB.HScrollBar scrlPercent 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   39
            Top             =   1440
            Value           =   1
            Width           =   1335
         End
         Begin VB.HScrollBar scrlRewards 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   38
            Top             =   720
            Value           =   1
            Width           =   1335
         End
         Begin VB.Label lblPercent 
            Caption         =   "Percent: 0%"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblRewards 
            Caption         =   "Reward num: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   7080
         Width           =   3975
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   29
         Top             =   6720
         Width           =   4815
      End
      Begin VB.HScrollBar scrlExhaustedPic 
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   1920
         Width           =   2295
      End
      Begin VB.PictureBox picExhaustedPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1680
         Left            =   2640
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   23
         Top             =   2280
         Width           =   2280
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":0A67
         Left            =   960
         List            =   "frmEditor_Resource.frx":0A77
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
         Width           =   3975
      End
      Begin VB.HScrollBar scrlNormalPic 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   2295
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   4320
         Width           =   4815
      End
      Begin VB.HScrollBar scrlTool 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   9
         Top             =   4920
         Width           =   4815
      End
      Begin VB.HScrollBar scrlHealth 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   8
         Top             =   5520
         Width           =   4815
      End
      Begin VB.PictureBox picNormalPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1680
         Left            =   120
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   7
         Top             =   2280
         Width           =   2280
      End
      Begin VB.HScrollBar scrlRespawn 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   2000
         TabIndex        =   6
         Top             =   6120
         Width           =   4815
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtMessage2 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation: None"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   6480
         Width           =   1260
      End
      Begin VB.Label lblExhaustedPic 
         AutoSize        =   -1  'True
         Caption         =   "Exhausted Image: 0"
         Height          =   180
         Left            =   2640
         TabIndex        =   25
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lblNormalPic 
         AutoSize        =   -1  'True
         Caption         =   "Normal Image: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1470
      End
      Begin VB.Label lblReward 
         AutoSize        =   -1  'True
         Caption         =   "Reward: None"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   4080
         Width           =   1035
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Tool Required: None"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   4680
         Width           =   1530
      End
      Begin VB.Label lblHealth 
         AutoSize        =   -1  'True
         Caption         =   "Health: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   5280
         Width           =   705
      End
      Begin VB.Label lblRespawn 
         AutoSize        =   -1  'True
         Caption         =   "Respawn Time (Seconds): 0"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   5880
         Width           =   2100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Success:"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   7800
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resource List"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7260
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkEqual_Click()

If chkEqual.value = False Then Exit Sub

Dim i As Byte, N As Byte
Dim BoolVect(1 To MAX_RESOURCE_REWARDS) As Boolean
    N = 0
    For i = 1 To MAX_RESOURCE_REWARDS
        If RewardsInfo(i).Reward <> 0 Then
            BoolVect(i) = True
            N = N + 1
        End If
    Next
    
    If N = 0 Then Exit Sub
    
    For i = 1 To MAX_RESOURCE_REWARDS
        If BoolVect(i) = True Then
            RewardsInfo(i).Chance = CByte(100 / N)
            Resource(EditorIndex).Rewards(i).Chance = RewardsInfo(i).Chance
        End If
    Next
    
    scrlPercent.value = RewardsInfo(scrlRewards.value).Chance
    
End Sub

Private Sub chkMessage_Click()
    'Personalize missage or not
    Resource(EditorIndex).ItemSuccessMessage = chkMessage.value
End Sub

Private Sub chkWalkableExhausted_Click()
    Resource(EditorIndex).WalkableExhausted = CBool(chkWalkableExhausted.value)
End Sub

Private Sub chkWalkableNormal_Click()
    Resource(EditorIndex).WalkableNormal = CBool(chkWalkableNormal.value)
End Sub

Private Sub cmbRewardType_click()
    Resource(EditorIndex).Rewards(scrlRewards.value).RewardType = cmbRewardType.ListIndex + 1
    SetScrlRewardMax
    DisplaylblReward
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).ResourceType = cmbType.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearResource EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).TranslatedName, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'Testing if chances summation is equal to 100
    
    Call ResourceEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub form_load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlReward.Max = MAX_ITEMS
    
    scrlRewards.Max = MAX_RESOURCE_REWARDS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ResourceEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlAnimation.value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.value).TranslatedName)
    lblAnim.Caption = "Animation: " & sString
    Resource(EditorIndex).Animation = scrlAnimation.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlExhaustedPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblExhaustedPic.Caption = "Exhausted Image: " & scrlExhaustedPic.value
    EditorResource_BltSprite
    Resource(EditorIndex).ExhaustedImage = scrlExhaustedPic.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlExhaustedPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlHealth_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblHealth.Caption = "Health: " & scrlHealth.value
    Resource(EditorIndex).health = scrlHealth.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlHealth_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNormalPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblNormalPic.Caption = "Normal Image: " & scrlNormalPic.value
    EditorResource_BltSprite
    Resource(EditorIndex).ResourceImage = scrlNormalPic.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNormalPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPercent_Change()

    RewardsInfo(scrlRewards.value).Chance = scrlPercent.value
    lblPercent.Caption = "Percent: " & scrlPercent.value & "%"
    
    Resource(EditorIndex).Rewards(scrlRewards.value).Chance = scrlPercent.value
    
End Sub

Private Sub scrlRespawn_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRespawn.Caption = "Respawn Time (Seconds): " & scrlRespawn.value
    Resource(EditorIndex).RespawnTime = scrlRespawn.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRespawn_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlReward_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    DisplaylblReward
    
    'Resource(EditorIndex).ItemReward = scrlReward.Value
    Resource(EditorIndex).Rewards(scrlRewards.value).Reward = scrlReward.value
    RewardsInfo(scrlRewards.value).Reward = scrlReward.value
    

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlReward_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRewards_Change()
    SetCmbReward
    SetScrlRewardMax
    DisplaylblReward
    lblRewards.Caption = "Reward num: " & scrlRewards.value
    
    If RewardsInfo(scrlRewards.value).Reward > GetScrlRewardMax Then
        scrlReward.value = 0
    Else
        scrlReward.value = RewardsInfo(scrlRewards.value).Reward
    End If
    
    scrlPercent.value = RewardsInfo(scrlRewards.value).Chance
    lblPercent.Caption = "Percent: " & RewardsInfo(scrlRewards.value).Chance & "%"
    If Resource(EditorIndex).Rewards(scrlRewards.value).RewardType > 1 And Resource(EditorIndex).Rewards(scrlRewards.value).RewardType <= cmbRewardType.ListCount Then
        cmbRewardType.ListIndex = Resource(EditorIndex).Rewards(scrlRewards.value).RewardType - 1
    End If
    

End Sub

Private Sub scrlTool_Change()
    Dim Name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlTool.value
        Case 0
            Name = "None"
        Case 1
            Name = "Hatchet"
        Case 2
            Name = "Rod"
        Case 3
            Name = "Pickaxe"
        Case 4
            Name = "Fire"
        Case 5
            Name = "Dig"
    End Select

    lblTool.Caption = "Tool Required: " & Name
    
    Resource(EditorIndex).ToolRequired = scrlTool.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTool_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).SuccessMessage = Trim$(txtMessage.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage2_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).EmptyMessage = Trim$(txtMessage2.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage2_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Resource(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).TranslatedName, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Resource(EditorIndex).Sound = cmbSound.list(cmbSound.ListIndex)
    Else
        Resource(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function CheckRewardTable() As Boolean
Dim sum As Byte, i As Byte, N As Byte

sum = 0
N = 0

For i = 1 To MAX_RESOURCE_REWARDS
    If RewardsInfo(i).Reward > 0 Then
        sum = sum + RewardsInfo(i).Chance
        N = N + 1
    End If
Next

If sum = 100 Or N = 0 Then
    CheckRewardTable = True
Else
    CheckRewardTable = False
End If

End Function

Private Sub DisplaylblReward()
Select Case cmbRewardType.ListIndex + 1

Case REWARD_ITEM
    If scrlReward.value > 0 Then
        lblReward.Caption = "Item Reward: " & Trim$(Item(scrlReward.value).TranslatedName) & "," & scrlReward.value
    Else
        lblReward.Caption = "Item Reward: None"
    End If
Case REWARD_SPAWN_NPC
    If scrlReward.value > 0 Then
        lblReward.Caption = "Spawn NPC: " & Trim$(NPC(scrlReward.value).TranslatedName) & "," & scrlReward.value
    Else
        lblReward.Caption = "Spawn NPC: None"
    End If
Case Else
    lblReward.Caption = "Undefined Type"
End Select

End Sub

Private Sub SetScrlRewardMax()
Select Case cmbRewardType.ListIndex + 1

Case REWARD_ITEM
    scrlReward.Max = MAX_ITEMS
Case REWARD_SPAWN_NPC
    scrlReward.Max = MAX_NPCS
Case Else
    scrlReward.Max = 1
End Select
End Sub

Public Function GetScrlRewardMax() As Long
Select Case cmbRewardType.ListIndex + 1

Case REWARD_ITEM
    GetScrlRewardMax = MAX_ITEMS
Case REWARD_SPAWN_NPC
    GetScrlRewardMax = MAX_NPCS
Case Else
    scrlReward.Max = 1
End Select
End Function

Private Sub SetCmbReward()
Select Case Resource(EditorIndex).Rewards(scrlRewards.value).RewardType

Case REWARD_ITEM
    cmbRewardType.ListIndex = REWARD_ITEM - 1
Case REWARD_SPAWN_NPC
    cmbRewardType.ListIndex = REWARD_SPAWN_NPC - 1
End Select
End Sub


