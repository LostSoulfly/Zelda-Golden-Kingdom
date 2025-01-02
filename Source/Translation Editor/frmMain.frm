VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
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
   Begin VB.CommandButton Command5 
      Caption         =   "En Cache"
      Height          =   375
      Left            =   600
      TabIndex        =   18
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Dis Cache"
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update MD5s"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Timer tmrTranslation 
      Interval        =   1500
      Left            =   600
      Top             =   4560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Recocnnect"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock sckTranslate 
      Left            =   240
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   6969
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TransTest"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CheckBox chkCompression 
      Caption         =   "Enable Compression"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton cmdClean 
      Caption         =   "Clean Translation"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
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
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Edit Translation"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Translation"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Translation"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
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

original = InputBox("What would you like to translate? (Please note that this is case and punctuation sensitive as it uses MD5 for storage!)", "Add Translation (Spanish to English")
If LenB(original) <= 0 Then Exit Sub
MD5 = BytesToHex(HashStringA(original))

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

Private Sub cmdClean_Click()
Dim MD5 As String
Dim Removed As Long
Dim removed2 As Long
Dim removedTotal As Long
Dim i As Long
Dim ii As Long
Dim Count As Long
If MsgBox("This will clean out the collection of unknown original translations and should bring the original and translated collections to the same number of records.", vbYesNo) = vbYes Then

Me.Caption = "Cleaning collections.."
DoEvents
    For ii = 0 To 5
        Count = col.Count
        For i = 1 To Count
        If i < (Count - Removed) Then
            MD5 = col.Item(i)(0)
                If Exists(origCol, MD5) = False Then
                    Debug.Print "Removing:" & col.Item(i)(1)
                    'DeleteFromCollection MD5, origCol
                    DeleteFromCollection MD5, col
                    Removed = Removed + 1
                End If
        End If
        Next i
        
        Count = origCol.Count
        For i = 1 To Count
        If i < (Count - removed2) Then
            MD5 = origCol.Item(i)(0)
                If Exists(col, MD5) = False Then
                    Debug.Print "Removing:" & origCol.Item(i)(1)
                    DeleteFromCollection MD5, origCol
                    'DeleteFromCollection MD5, col
                    removed2 = removed2 + 1
                End If
        End If
        Next i
        
        removedTotal = removedTotal + Removed + removed2
        removed2 = 0
        Removed = 0
    Next ii
End If
    
    updateCaption
    MsgBox "Removed items from the collections: " & removedTotal, vbOKOnly
    
End Sub

Private Sub UpdateMD5s()
Dim oldMD5 As String
Dim newMD5 As String
Dim Spanish As String
Dim English As String
Dim i As Long
Dim ii As Long
Dim Removed As Long
Dim Count As Long
'If MsgBox("This will recalculate MD5s in both translation databases. Continue?", vbYesNo) = vbYes Then

Me.Caption = "Updating MD5s.."
DoEvents
        
        Count = origCol.Count
        For i = 0 To Count
        If i < (Count) Then
            oldMD5 = origCol.Item(i + 1)(0)
            If InStr(1, oldMD5, "-") <> 0 Then
                If Exists(col, oldMD5) = True Then
                    Debug.Print "Updating:" & origCol.Item(i + 1)(0)
                    Spanish = ReadFromCache(oldMD5, origCol)
                    English = ReadFromCache(oldMD5, col)
                    newMD5 = BytesToHex(HashStringA(Spanish))
                    
                    Debug.Print ReadFromCache(oldMD5, origCol)
                    origCol.Remove oldMD5
                    col.Remove oldMD5
                    Debug.Print ReadFromCache(oldMD5, origCol)
                    'DeleteFromCollection oldMD5, origCol
                    'DeleteFromCollection oldMD5, col
                    
                    AddToCache newMD5, Spanish, origCol
                    AddToCache newMD5, English, col
                    
                Else
                    Debug.Print "Updating:" & origCol.Item(i + 1)(0)
                    Spanish = ReadFromCache(oldMD5, origCol)
                    English = ReadFromCache(oldMD5, col)
                    newMD5 = BytesToHex(HashStringA(Spanish))
                    
                    DeleteFromCollection oldMD5, origCol
                    DeleteFromCollection oldMD5, col
                    
                    AddToCache newMD5, Spanish, origCol
                    AddToCache newMD5, English, col
                    
                End If
            End If
        End If
        Next i

    updateCaption
    MsgBox "Done", vbOKOnly
    
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
cmdClean.Enabled = True
Command3.Enabled = True
cmdMassReplace.Enabled = True
LoadLanguage App.Path & "\en.dat", col
LoadLanguage App.Path & "\es-en.dat", origCol
updateCaption
End Sub

Private Sub cmdMassReplace_Click()

Dim strFind As String
Dim strReplace As String
Dim strTemp As String
Dim MD5 As String
Dim Count As Long
Dim Removed As Long
Dim TotalRemoved As Long

MsgBox "This will only replace text in the Translated collection!"

strFind = InputBox("What text are we looking for? CASE SENSITIVE!", "Find Text")
strReplace = InputBox("What text are we Replacing it with? CASE SENSITIVE!", "Replace Text")

If LenB(strFind) = 0 Then Exit Sub
If LenB(strReplace) = 0 Then Exit Sub

Dim i As Long

Removed = 1

Do While Removed > 0
    Removed = 0
    
    Count = col.Count
    For i = 1 To Count
        If i < (Count - Removed) Then
            If InStr(1, col.Item(i)(1), strFind) <> 0 Then
                MD5 = col.Item(i)(0)
                strTemp = Replace(col.Item(i)(1), strFind, strReplace)
                DeleteFromCollection MD5, col
                DoEvents
                AddToCache MD5, strTemp, col
                Removed = Removed + 1
            End If
        End If
    Next i
TotalRemoved = TotalRemoved + Removed
Loop
    
    MsgBox "Removed items from the collections: " & TotalRemoved, vbOKOnly
    
End Sub

Private Sub cmdModify_Click()
    
If lstResults.List(lstResults.ListIndex) = "" Then Exit Sub
    
    Dim newTemp As String
    Dim temp As String
    Dim strOriginal As String
    Dim MD5 As String
    
    temp = col.Item(lstResults.ItemData(lstResults.ListIndex))(1)
    temp = Replace(temp, vbCr, "\r")
    temp = Replace(temp, vbLf, "\n")
    temp = Replace(temp, vbNewLine, "\r\n")
    
    MD5 = col.Item(lstResults.ItemData(lstResults.ListIndex))(0)
    strOriginal = origCol.Item(MD5)(1)
    newTemp = InputBox("Current Translation: " & strOriginal, "Edit Translation", temp)
    
    temp = Replace(temp, "\r", vbCr)
    temp = Replace(temp, "\n", vbLf)
    temp = Replace(temp, "\r\n", vbNewLine)
    
    
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
            lstResults.AddItem Replace(coll.Item(i)(1), vbNewLine, "\r\n")
            lstResults.ItemData(lstResults.NewIndex) = i
        End If
    End If
Next i

End Sub

Public Sub EnumCollection(coll As Collection)
lstResults.Clear
Dim i As Long
For i = 1 To coll.Count

            lstResults.AddItem Replace(coll.Item(i)(1), vbNewLine, "\r\n")
            lstResults.ItemData(lstResults.NewIndex) = i

Next i

End Sub

Public Sub DeleteFromCollection(index As Variant, coll As Collection)
On Error Resume Next
    coll.Remove (index)
End Sub

Private Sub Command1_Click()
MsgBox Translate("Hola my nombre Brian", 12)
End Sub

Private Sub Command2_Click()
sckTranslate.Close
sckTranslate.RemoteHost = "127.0.0.1"
sckTranslate.RemotePort = "6969"
sckTranslate.Connect
End Sub

Private Sub Command3_Click()
UpdateMD5s
End Sub

Private Sub Command4_Click()
sckTranslate.SendData "14" + vbCrLf
End Sub

Private Sub Command5_Click()
sckTranslate.SendData "15" + vbCrLf
End Sub

Private Sub Form_Load()
'If t Is Nothing Then Set t = New GTranslateDLL.DLL
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
Dim temp As String
cmdDelete.Enabled = True
cmdModify.Enabled = True
cmdReTranslate.Enabled = True

If usingOrigCol = False Then
MD5 = col.Item(lstResults.ItemData(lstResults.ListIndex))(0)
Else
MD5 = origCol.Item(lstResults.ItemData(lstResults.ListIndex))(0)
End If

'example of the different types of issues that could crop up when editing the translation..
'I probably should have enforced a stricter control on these initially.
'Oops.
temp = col.Item(lstResults.ItemData(lstResults.ListIndex))(1)
temp = Replace(temp, vbCr, "\r")
temp = Replace(temp, vbLf, "\n")
temp = Replace(temp, vbNewLine, "\r\n")

Debug.Print "Selected: " & temp
If usingOrigCol Then
Debug.Print "Original: " & Replace(col.Item(MD5)(1), vbNewLine, "\r\n")
Else
Debug.Print "Original: " & Replace(origCol.Item(MD5)(1), vbNewLine, "\r\n")
End If
Debug.Print MD5

lstResults.ToolTipText = Replace(origCol.Item(MD5)(1), vbNewLine, "\r\n")

End Sub

Private Sub sckTranslate_Connect()
'sckTranslate.SendData "14" + vbCrLf
End Sub

Private Sub sckTranslate_DataArrival(ByVal bytesTotal As Long)
Static buffer As String
Dim NewData As String
Dim Msgs() As String
Dim MD5 As String
Dim i As Integer

sckTranslate.GetData NewData
Msgs = Split(buffer & NewData, vbNewLine)
buffer = Msgs(UBound(Msgs))
For i = 0 To UBound(Msgs) - 1
If Mid(Msgs(i), 1, 2) = 99 Then
    MsgBox Mid(Msgs(i), 3)
Else
    MD5 = Mid(Msgs(i), 1, 32)
    Msgs(i) = Mid(Msgs(i), 33)
    Msgs(i) = Replace(Msgs(i), "\r", vbCr)
    Msgs(i) = Replace(Msgs(i), "\n", vbLf)
    Msgs(i) = Replace(Msgs(i), "\r\n", vbNewLine)
    modTranslate.AddToCache MD5, Msgs(i), modTranslate.transCol
End If
    'MsgBox Msgs(I)
Next
End Sub

Private Sub tmrTranslation_Timer()
If sckTranslate.State = sckClosed Then
    sckTranslate.Connect
    Exit Sub
End If

If sckTranslate.State = sckConnecting Then
    sckTranslate.Close
    sckTranslate.Connect
    Exit Sub
End If

If sckTranslate.State = sckConnected Then
    frmMain.Caption = "Translate Server Connected - TransEdit"
Else
    frmMain.Caption = "Translate Server Not Connected - TransEdit"
End If


End Sub
