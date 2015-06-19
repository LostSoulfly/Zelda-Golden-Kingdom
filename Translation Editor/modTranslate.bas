Attribute VB_Name = "modTranslate"
Option Explicit
Private Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)

'Save the collection every 50 translations
Private Const saveEvery = 50
'Max of 6 translations per second. Good luck hitting that very often!
Private Const intTransPerSec As Integer = 10
'Sleep the server every 1ms until it's able to translate again.
Public Const blWaitToTranslate As Boolean = True

'these could be const, but I didn't want them to be.
Public LangTo As String
Public LangFrom As String
Public strTransPath As String
Public strOrigPath As String

'last number of total translations saved at
Private lastSave As Long
'timer for the translations
Private tmrTrans As Long
'number of translations done in the last second
Private TransCount As Integer
'the collection for the current language.
'You can make multiples of these and pass them to each of the functions/subs that use them.
Public langCol As Collection
Public origCol As Collection

Public t As GTranslate.DLL

'I try to use byref wherever possible to prevent VB from having to
'make copies of strings continuously just to process them.
'it should actually lead to a small speed increase,
'but you have to be careful to not modify the original string

Public Function ReadFromCache(ByRef strHash As String, ByRef col As Collection) As String
Dim temp() As String

If col.Count = 1 Then ReadFromCache = "": Exit Function

If Exists(col, strHash) = True Then ReadFromCache = col.Item(strHash)(1)

End Function

Public Sub AddToCache(ByRef strHash As String, ByRef Translate As String, ByRef col As Collection)
Dim temp(1) As String

buildArray strHash, Translate, temp
col.Add temp, temp(0)

End Sub

Public Sub loadLang(Path As String, ByRef col As Collection)
'Dim strTemp As String
Dim temp() As Byte
Dim tempArray(1) As String
Dim buffer As New clsBuffer
Dim lngBufferCount As Long
Dim NF As Integer
Dim NotNull As Boolean
Dim bfFail As Boolean
NF = FreeFile

' check exists
    Open Path For Binary As NF
    Close NF

temp = ReadFile(Path, NotNull)

If NotNull = False Then Exit Sub

    temp = Decompress(temp, bfFail)
    If bfFail = True Then GoTo skip
    buffer.WriteBytes temp
    
lngBufferCount = buffer.ReadLong
lastSave = lngBufferCount
Dim i As Long
For i = 1 To lngBufferCount
    buildArray buffer.ReadString, buffer.ReadString, tempArray
    col.Add tempArray, tempArray(0)
Next

'set our lastSave variable so we can save again in 50 translations
lastSave = col.Count

skip:
Set buffer = Nothing
End Sub

Public Sub saveLang(Path As String, ByRef col As Collection)
Dim NF As Integer
Dim tempOut() As Byte
Dim buffer As New clsBuffer
Dim i As Long
NF = FreeFile

'If (lastSave) = (langCol.Count) Then Exit Sub
Debug.Print "Saving lang to: " & Path
buffer.WriteLong col.Count

For i = 1 To col.Count
'write the key first
    buffer.WriteString (col.Item(i)(0))
'write the actual translation
    buffer.WriteString (col.Item(i)(1))
Next

'write buffer to temp out
tempOut = Compress(buffer.ReadBytes(buffer.length))

    Open Path For Binary As #NF
    Put #NF, , tempOut
    Close #NF
    
lastSave = langCol.Count
End Sub

Private Sub buildArray(ByRef key As String, ByRef text As String, ByRef myArr() As String)

If LenB(text) <= 0 Then
    Debug.Print
End If

myArr(0) = key
myArr(1) = text

End Sub

'use a buffer class to write the
'key length, then the key. use long - string.
'the variable length, then the variable. use long - string
'write them into a collection. How long can the keys be?

Public Function ReadFile(sFile As String, Optional ByRef NotNull As Boolean) As Byte()
    Dim nFile As Integer

    nFile = FreeFile
    Open sFile For Binary Access Read As #nFile
    If LOF(nFile) > 0 Then
        ReDim ReadFile(0 To LOF(nFile) - 1)
        Get nFile, , ReadFile
        NotNull = True
    Else
        NotNull = False
    End If
    Close #nFile
End Function

Public Sub debugLangFile(Path As String)

Dim NF As Integer
Dim tempOut() As Byte
Dim buffer As New clsBuffer
NF = FreeFile

Dim col As Collection
Set col = New Collection

Dim myArr(1) As String
col.Add Array("0a-6d-d0-dd-a2-ee-52-6b-57-55-6b-68-97-33-4a-b1", "Level 40-50"), "0a-6d-d0-dd-a2-ee-52-6b-57-55-6b-68-97-33-4a-b1"
col.Add Array("da-f8-9b-9c-2d-b8-51-d6-91-84-f0-95-6a-44-a0-d3", "Level 5-10"), "da-f8-9b-9c-2d-b8-51-d6-91-84-f0-95-6a-44-a0-d3"
col.Add Array("7e-54-5b-21-c5-a5-49-23-4b-ca-43-23-32-cf-54-82", "Level 10-20"), "7e-54-5b-21-c5-a5-49-23-4b-ca-43-23-32-cf-54-82"
buffer.WriteLong col.Count

Dim i As Long


For i = 1 To col.Count
'write the key first
    buffer.WriteString (col.Item(i)(0))
'write the actual translation
    buffer.WriteString (col.Item(i)(1))
Next i

'write buffer to temp out
tempOut = Compress(buffer.ReadBytes(buffer.length))

    Open Path For Binary As #NF
    Put #NF, , tempOut
    Close #NF

Set col = Nothing

End Sub

Public Function Exists(ByRef col As Collection, ByRef index) As Boolean
On Error GoTo ExistsTryNonObject
    Dim O As Object

    Set O = col(index)
    Exists = True
    Exit Function

ExistsTryNonObject:
    Exists = ExistsNonObject(col, index)
End Function

Private Function ExistsNonObject(ByRef col As Collection, ByRef index) As Boolean
On Error GoTo ExistsNonObjectErrorHandler
    Dim v As Variant

    v = col(index)
    ExistsNonObject = True
    Exit Function

ExistsNonObjectErrorHandler:
    ExistsNonObject = False
End Function
