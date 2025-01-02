VERSION 5.00
Begin VB.Form frmReTranslate 
   Caption         =   "Re-Translate"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3480
      TabIndex        =   4
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Use Selected"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6000
      TabIndex        =   3
      Top             =   5520
      Width           =   2175
   End
   Begin VB.OptionButton optYandex 
      Caption         =   "Option3"
      Height          =   1575
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Width           =   8415
   End
   Begin VB.OptionButton optBing 
      Caption         =   "Option1"
      Height          =   1575
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   8415
   End
   Begin VB.OptionButton optGoogle 
      Caption         =   "Option1"
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8415
   End
End
Attribute VB_Name = "frmReTranslate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private selectedTranslation As String
Private blCanceled As Boolean

Public Function hasCanceled() As Boolean
    hasCanceled = blCanceled
End Function

Public Function ChosenTranslation() As String
    ChosenTranslation = selectedTranslation
End Function

Public Sub TranslateThis(Text As String)
On Error Resume Next

    optBing.Caption = Translate(Text, 11)
    optGoogle.Caption = Translate(Text, 10)
    optYandex.Caption = Translate(Text, 12)

cmdCancel.Enabled = True
cmdSelect.Enabled = True
End Sub

Private Sub cmdCancel_Click()
blCanceled = True
Me.Visible = False
End Sub

Private Sub cmdSelect_Click()
Dim temp As String
blCanceled = False
Me.Visible = False
End Sub

Private Sub Form_Activate()
Form_Resize
End Sub

Private Sub Form_Load()
optGoogle.Font = "Arial"
optBing.Font = optGoogle.Font
optYandex.Font = optBing.Font
End Sub

Private Sub Form_Resize()
On Error Resume Next
'I'm terrible at this. How can I do this better??
optGoogle.Width = frmReTranslate.Width / 1.1
optBing.Width = optGoogle.Width
optYandex.Width = optBing.Width
optGoogle.Height = (frmReTranslate.Height / 3) - (frmReTranslate.Height / 8)
optBing.Height = optGoogle.Height
optYandex.Height = optBing.Height
optGoogle.Top = (frmReTranslate.Height / 2) - ((optGoogle.Height / 0.48))
optBing.Top = optGoogle.Height + 250 + optGoogle.Top
optYandex.Top = optBing.Top + 250 + optBing.Height
optGoogle.Left = (frmReTranslate.Width - optGoogle.Width - 130) / 2
optBing.Left = optGoogle.Left
optYandex.Left = optBing.Left
cmdCancel.Top = optYandex.Top + optYandex.Height + 100
cmdSelect.Top = cmdCancel.Top
cmdCancel.Left = optYandex.Left
cmdSelect.Left = frmReTranslate.Width - (cmdSelect.Width / 0.836)
cmdCancel.Height = frmReTranslate.Height / 7
cmdSelect.Height = cmdCancel.Height
cmdCancel.Width = frmReTranslate.Width / 4
cmdSelect.Width = cmdCancel.Width
'increase fontsize as page gets bigger?
Dim intInt As Integer
intInt = optGoogle.Width / 650
optGoogle.FontSize = intInt
optBing.FontSize = intInt
optYandex.FontSize = intInt

End Sub

Private Sub optBing_Click()
selectedTranslation = optBing.Caption
End Sub

Private Sub optBing_DblClick()
Dim temp As String
With optBing
    temp = InputBox("Edit this translation.", "Edit Translation", .Caption)
    If LenB(temp) > 0 Then .Caption = temp
End With
optBing_Click
End Sub

Private Sub optGoogle_Click()
selectedTranslation = optGoogle.Caption
End Sub

Private Sub optGoogle_DblClick()
Dim temp As String
With optGoogle
    temp = InputBox("Edit this translation.", "Edit Translation", .Caption)
    If LenB(temp) > 0 Then .Caption = temp
End With
optGoogle_Click
End Sub

Private Sub optYandex_Click()
selectedTranslation = optYandex.Caption
End Sub

Private Sub optYandex_DblClick()
Dim temp As String
With optYandex
    temp = InputBox("Edit this translation.", "Edit Translation", .Caption)
    If LenB(temp) > 0 Then .Caption = temp
End With
optYandex_Click
End Sub
