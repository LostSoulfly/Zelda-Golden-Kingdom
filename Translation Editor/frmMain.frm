VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "TranslationEditor"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMassReplace 
      Caption         =   "Mass Replace"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Show All Entries"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find By MD5"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearchOrig 
      Caption         =   "Search Original"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdReTranslate 
      Caption         =   "Re-Translate"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Edit Translation"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Translation"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Translation"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.ListBox lstResults 
      Height          =   5520
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Translation"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Translation"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Translation"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim col As Collection
Dim origCol As Collection
Dim usingOrigCol As Boolean
Dim lastSearch As String

Private Sub cmdAdd_Click()
'multi ways to do this.
'want to add it in a different language and have it translated?
'or directly by MD5? probably not lol..
Dim f As New frmReTranslate
Dim newTemp As String
Dim MD5 As String
Dim original As String

original = InputBox("What would you like to translate? (Please note that this is case and puctuation sensitive as it uses MD5 for storage!)", "Add Translation (Spanish to English")
If LenB(original) <= 0 Then Exit Sub
MD5 = t.GetMD5Hash(original)

Load f
f.TranslateThis original
f.Show vbModal
If f.hasCanceled = True Then: Set f = Nothing
newTemp = f.ChosenTranslation
If LenB(newTemp) <= 0 Then Exit Sub
If newTemp = original Then Exit Sub
        
    AddToCache MD5, original, origCol
    AddToCache MD5, newTemp, col
    
End Sub

Private Sub cmdAll_Click()
EnumCollection col
lastSearch = ""
End Sub

Private Sub setUseTransCol()
    usingOrigCol = False
End Sub

Private Sub SetUseOrigCol()
    usingOrigCol = True
End Sub

Private Sub updateButtons()
If usingOrigCol = True Then
    cmdDelete.Enabled = True
    cmdModify.Enabled = True
Else
cmdReTranslate.Enabled = True
End If
End Sub

Private Sub cmdDelete_Click()
Dim MD5 As String
If lstResults.List(lstResults.ListIndex) = "" Then Exit Sub
If usingOrigCol = True Then
MD5 = origCol.Item(lstResults.ItemData(lstResults.ListIndex))(0)
Else
MD5 = col.Item(lstResults.ItemData(lstResults.ListIndex))(0)
End If

DeleteFromCollection MD5, origCol
DeleteFromCollection MD5, col

lstResults.Clear
If Not lastSearch = "" Then SearchCollection lastSearch, col Else cmdAll_Click
End Sub

Private Sub cmdFind_Click()
Dim temp As String
temp = InputBox("Search for what?", "Search MD5")
If LenB(temp) <= 0 Then Exit Sub
SearchCollection temp, col, True
End Sub

Private Sub cmdLoad_Click()

Set col = Nothing
Set origCol = Nothing

Set col = New Collection
Set origCol = New Collection


cmdSave.Enabled = True
cmdSearch.Enabled = True
cmdSearchOrig.Enabled = True
cmdFind.Enabled = True
cmdAll.Enabled = True
cmdAdd.Enabled = True
cmdMassReplace.Enabled = True
LoadLanguage App.Path & "\en.dat", col
LoadLanguage App.Path & "\es-en.dat", origCol
updateCaption
End Sub

Private Sub cmdModify_Click()
    
If lstResults.List(lstResults.ListIndex) = "" Then Exit Sub
    
    Dim newTemp As String
    Dim temp As String
    Dim MD5 As String
    temp = col.Item(lstResults.ItemData(lstResults.ListIndex))(1)
    MD5 = col.Item(lstResults.ItemData(lstResults.ListIndex))(0)
    newTemp = InputBox("Current Translation: & temp", "Edit Translation", temp)
    
    If LenB(newTemp) <= 0 Then Exit Sub
    If newTemp = temp Then Exit Sub
    
    DeleteFromCollection lstResults.ItemData(lstResults.ListIndex), col
    AddToCache MD5, newTemp, col
    'UpdateList text, index, lstResults
    'col.Add(lstResults.ItemData(lstResults.ListIndex))(1) = newTemp
    lstResults.Clear
    If Not lastSearch = "" Then SearchCollection lastSearch, col Else cmdAll_Click

End Sub

Private Sub updateCaption()
If col.Count = origCol.Count Then
Me.Caption = "TranslationEditor - Records: " & origCol.Count
Else
Me.Caption = "TranslationEditor - Trans:" & col.Count & " - Orig:" & origCol.Count
End If

End Sub

Private Sub DeleteFromList(index As Long, lst As ListBox)
    lst.RemoveItem (index)
End Sub

Private Sub UpdateList(Text As String, index As Long, lst As ListBox)

    lst.List(index) = Text

End Sub

Private Sub cmdReTranslate_Click()
Dim f As New frmReTranslate
Dim newTemp As String
Dim temp As String
Dim MD5 As String

If lstResults.List(lstResults.ListIndex) = "" Then Exit Sub

If usingOrigCol = False Then
    MD5 = col.Item(lstResults.ItemData(lstResults.ListIndex))(0)
Else
    MD5 = origCol.Item(lstResults.ItemData(lstResults.ListIndex))(0)
End If



temp = origCol.Item(MD5)(1)

Load f

f.TranslateThis temp

f.Show vbModal

If f.hasCanceled = True Then: Set f = Nothing

newTemp = f.ChosenTranslation

If LenB(newTemp) <= 0 Then Exit Sub
If newTemp = temp Then Exit Sub
    
    DeleteFromCollection lstResults.ItemData(lstResults.ListIndex), col
    AddToCache MD5, newTemp, col

lstResults.Clear
'UpdateList newTemp, lstResults.ListIndex, lstResults
If Not lastSearch = "" Then SearchCollection lastSearch, col Else cmdAll_Click
Unload f
Set f = Nothing

End Sub

Private Sub cmdSearch_Click()
Dim temp As String
temp = InputBox("Search for what?", "Search Translations")
If LenB(temp) <= 0 Then Exit Sub
setUseTransCol
SearchCollection temp, col
lastSearch = temp
End Sub

Private Sub cmdSearchOrig_Click()
Dim temp As String
temp = InputBox("Search for what?", "Search Original (untranslated)")
If LenB(temp) <= 0 Then Exit Sub
SetUseOrigCol
SearchCollection temp, origCol

End Sub

Public Sub LoadLanguage(Path As String, coll As Collection)
    loadLang Path, coll
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Kill App.Path & "\en.dat.bak"
'FileCopy App.Path & "\en.dat", App.Path & "\en.dat.bak"
'Kill App.Path & "\en.dat"
    saveLang App.Path & "\en.dat", col
    saveLang App.Path & "\es-en.dat", origCol
End Sub

Public Sub SearchCollection(Text As String, coll As Collection, Optional blMD5Search As Boolean)
lstResults.Clear
Dim i As Long
For i = 1 To coll.Count
    
    If blMD5Search = True Then
        If InStr(1, LCase$(coll.Item(i)(0)), LCase$(Text)) <> 0 Then
            lstResults.AddItem coll.Item(i)(0)
            lstResults.ItemData(lstResults.NewIndex) = i
            Exit Sub
        End If
    Else
        If InStr(1, LCase$(coll.Item(i)(1)), LCase$(Text)) <> 0 Then
            lstResults.AddItem coll.Item(i)(1)
            lstResults.ItemData(lstResults.NewIndex) = i
        End If
    End If
Next i

End Sub

Public Sub EnumCollection(coll As Collection)
lstResults.Clear
Dim i As Long
For i = 1 To coll.Count

            lstResults.AddItem coll.Item(i)(1)
            lstResults.ItemData(lstResults.NewIndex) = i

Next i

End Sub

Public Sub DeleteFromCollection(index As Variant, coll As Collection)
On Error Resume Next
    coll.Remove (index)
End Sub

Private Sub Form_Load()
If t Is Nothing Then Set t = New GTranslate.DLL
End Sub

Private Sub Form_Resize()
On Error Resume Next
lstResults.Width = frmMain.Width - 2100
lstResults.Height = frmMain.Height - 600
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstResults_Click()
On Error Resume Next
Dim MD5 As String
cmdDelete.Enabled = True
cmdModify.Enabled = True
cmdReTranslate.Enabled = True

If usingOrigCol = False Then
MD5 = col.Item(lstResults.ItemData(lstResults.ListIndex))(0)
Else
MD5 = origCol.Item(lstResults.ItemData(lstResults.ListIndex))(0)
End If

Debug.Print "Selected: " & col.Item(lstResults.ItemData(lstResults.ListIndex))(1)
Debug.Print "Original: " & origCol.Item(MD5)(1)

lstResults.ToolTipText = origCol.Item(MD5)(1)

End Sub
