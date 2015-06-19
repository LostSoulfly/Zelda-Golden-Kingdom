VERSION 5.00
Begin VB.Form frmGuildAdmin 
   Caption         =   "Guild Panel"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Opciones del Líder del Clan"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Command3 
         Caption         =   "Opciones"
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Usuarios"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Grados"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frameMainUsers 
      Caption         =   "Editar Usuarios"
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame frameUser 
         BorderStyle     =   0  'None
         Caption         =   "Hide Later"
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   4695
         Begin VB.CommandButton cmduser 
            Caption         =   "Guardar Usuario"
            Height          =   255
            Left            =   1680
            TabIndex        =   18
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox txtcomment 
            Height          =   855
            Left            =   840
            TabIndex        =   15
            Top             =   600
            Width           =   3735
         End
         Begin VB.ComboBox cmbRanks 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   230
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Comentario:"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Grado:"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.ListBox listusers 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame frameMainRanks 
      Caption         =   "Editar Grados"
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ListBox listranks 
         Height          =   1425
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4815
      End
      Begin VB.Frame frameranks 
         BorderStyle     =   0  'None
         Caption         =   "Hide Later"
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   4815
         Begin VB.OptionButton opAccess 
            Caption         =   "No"
            Height          =   375
            Index           =   0
            Left            =   4080
            TabIndex        =   29
            Top             =   1200
            Width           =   615
         End
         Begin VB.OptionButton opAccess 
            Caption         =   "Sí"
            Height          =   255
            Index           =   1
            Left            =   4080
            TabIndex        =   28
            Top             =   840
            Width           =   495
         End
         Begin VB.ListBox listAccess 
            Height          =   1425
            ItemData        =   "frmGuildAdmin.frx":0000
            Left            =   840
            List            =   "frmGuildAdmin.frx":0002
            TabIndex        =   26
            Top             =   480
            Width           =   3015
         End
         Begin VB.CommandButton cmdRankSave 
            Caption         =   "Guardar Grados"
            Height          =   255
            Left            =   1680
            TabIndex        =   17
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   840
            TabIndex        =   8
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label5 
            Caption         =   "¿Puede?"
            Height          =   255
            Left            =   4020
            TabIndex        =   31
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Grado:"
            Height          =   255
            Left            =   360
            TabIndex        =   27
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Width           =   615
         End
      End
   End
   Begin VB.Frame frameMainoptions 
      Caption         =   "Editar Opciones"
      Height          =   4215
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   5055
      Begin VB.ComboBox cmbColor 
         Height          =   315
         ItemData        =   "frmGuildAdmin.frx":0004
         Left            =   1320
         List            =   "frmGuildAdmin.frx":001D
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtMOTD 
         Height          =   975
         Left            =   960
         TabIndex        =   25
         Text            =   "Text3"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton cmdoptions 
         Caption         =   "Guardar Opciones"
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   3840
         Width           =   1455
      End
      Begin VB.HScrollBar scrlRecruits 
         Height          =   255
         Left            =   2760
         Max             =   6
         Min             =   1
         TabIndex        =   21
         Top             =   840
         Value           =   1
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Mensaje del Día:"
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblrecruit 
         Caption         =   "100"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Los reclutas comienzan en Grado:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Color del Clan:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmGuildAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbColor_Change()
    GuildData.Guild_Color = cmbColor.ListIndex
End Sub

Private Sub cmbRanks_Click()
    If listusers.ListIndex > 0 Then
        GuildData.Guild_Members(listusers.ListIndex).Rank = cmbRanks.ListIndex
    End If
End Sub

Private Sub cmdoptions_Click()
    Call GuildSave(1, 1)
End Sub

Private Sub cmdRankSave_Click()
    Call GuildSave(3, listranks.ListIndex)
End Sub

Private Sub cmduser_Click()
     Call GuildSave(2, listusers.ListIndex)
End Sub

Private Sub Command1_Click()
    frameMainRanks.Visible = True
    frameMainUsers.Visible = False
    frameMainoptions.Visible = False
End Sub

Private Sub Command2_Click()
    frameMainRanks.Visible = False
    frameMainUsers.Visible = True
    frameMainoptions.Visible = False
End Sub

Private Sub Command3_Click()
    frameMainRanks.Visible = False
    frameMainUsers.Visible = False
    frameMainoptions.Visible = True
End Sub

Private Sub Form_Load()
 'Load all 3 on load
Call Load_Guild_Admin
 
End Sub
Public Sub Load_Guild_Admin()
 Call Load_Menu_Options
 Call Load_Menu_Ranks
 Call Load_Menu_Users
End Sub
Public Sub Load_Menu_Options()
scrlRecruits.Max = MAX_GUILD_RANKS
scrlRecruits.Value = GuildData.Guild_RecruitRank
cmbColor.ListIndex = GuildData.Guild_Color

txtMOTD.text = GuildData.Guild_MOTD
End Sub
Public Sub Load_Menu_Ranks()
Dim i As Integer

listranks.Clear
listranks.AddItem ("Select rank to edit...")
For i = 1 To MAX_GUILD_RANKS
    listranks.AddItem ("Rank #" & i & ": " & GuildData.Guild_Ranks(i).Name)
Next i

    For i = 0 To 1
        opAccess(i).Visible = False
    Next i

frameranks.Visible = False
listranks.ListIndex = 0


End Sub
Public Sub Load_Menu_Users()
Dim i As Integer

listusers.Clear
listusers.AddItem ("Select user to edit...")

For i = 1 To MAX_GUILD_MEMBERS
    listusers.AddItem ("User #" & i & ": " & GuildData.Guild_Members(i).User_Name)
Next i

cmbRanks.Clear
cmbRanks.AddItem ("Must Select Ranks...")
cmbRanks.ListIndex = 0
For i = 1 To MAX_GUILD_RANKS
    cmbRanks.AddItem (GuildData.Guild_Ranks(i).Name)
Next i

frameUser.Visible = False
listusers.ListIndex = 0
End Sub

Private Sub listAccess_Click()
Dim i As Integer

If listAccess.ListIndex = 0 Then
    For i = 0 To 1
        opAccess(i).Visible = False
    Next i
    Exit Sub
Else
    For i = 0 To 1
        opAccess(i).Visible = True
    Next i
End If

    opAccess(GuildData.Guild_Ranks(listranks.ListIndex).RankPermission(listAccess.ListIndex)).Value = True
End Sub

Private Sub listranks_Click()
Dim i As Integer
Dim HoldString As String

    If listranks.ListIndex = 0 Then
        frameranks.Visible = False
        Exit Sub
    End If
    
    cmdRankSave.Caption = "Save Rank #" & listranks.ListIndex
    txtName.text = GuildData.Guild_Ranks(listranks.ListIndex).Name
    
listAccess.Clear
listAccess.AddItem ("Select permission to edit...")
For i = 1 To MAX_GUILD_RANKS_PERMISSION
    If GuildData.Guild_Ranks(listranks.ListIndex).RankPermission(i) = 1 Then
        HoldString = "Can"
    Else
        HoldString = "Cannot"
    End If
    listAccess.AddItem (GuildData.Guild_Ranks(listranks.ListIndex).RankPermissionName(i) & " (" & HoldString & ")")
Next i
    
    For i = 0 To 1
        opAccess(i).Visible = False
    Next i
    
    frameranks.Visible = True
End Sub

Private Sub listusers_Click()
Dim i As Integer
    
    If listusers.ListIndex = 0 Then
        frameUser.Visible = False
        Exit Sub
    End If
    cmduser.Caption = "Save User #" & listusers.ListIndex
    txtcomment.text = GuildData.Guild_Members(listusers.ListIndex).Comment
    cmbRanks.ListIndex = GuildData.Guild_Members(listusers.ListIndex).Rank

    If Not GuildData.Guild_Members(listusers.ListIndex).User_Name = vbNullString Then
        frameUser.Visible = True
    Else
        frameUser.Visible = False
    End If

End Sub

Private Sub opAccess_Click(Index As Integer)
Dim HoldString As String

 If listranks.ListIndex = 0 Then Exit Sub
 If listAccess.ListIndex = 0 Then Exit Sub
 
 GuildData.Guild_Ranks(listranks.ListIndex).RankPermission(listAccess.ListIndex) = Index
 
    If GuildData.Guild_Ranks(listranks.ListIndex).RankPermission(listAccess.ListIndex) = 1 Then
        HoldString = "Can"
    Else
        HoldString = "Cannot"
    End If
    
    listAccess.List(listAccess.ListIndex) = GuildData.Guild_Ranks(listranks.ListIndex).RankPermissionName(listAccess.ListIndex) & " (" & HoldString & ")"
End Sub

Private Sub scrlRecruits_Change()
    lblrecruit.Caption = scrlRecruits.Value
    GuildData.Guild_RecruitRank = scrlRecruits.Value
End Sub

Private Sub txtcomment_Change()
    If listusers.ListIndex = 0 Then Exit Sub
    
    GuildData.Guild_Members(listusers.ListIndex).Comment = txtcomment.text
End Sub

Private Sub txtMOTD_Change()
    GuildData.Guild_MOTD = txtMOTD.text
End Sub

Private Sub txtName_Change()
If listranks.ListIndex = 0 Then Exit Sub

GuildData.Guild_Ranks(listranks.ListIndex).Name = txtName.text
End Sub
