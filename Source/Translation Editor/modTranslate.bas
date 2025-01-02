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
Public transCol As Collection

Private Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
 
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByRef pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByRef pByte As Any, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptBinaryToString Lib "Crypt32.dll" Alias "CryptBinaryToStringA" (ByRef pbBinary As Any, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As String, ByRef pcchString As Long) As Long
 
Private Const PROV_RSA_AES As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
 
Public Enum HashAlgo
    HALG_MD2 = &H8001&
    HALG_MD4 = &H8002&
    HALG_MD5 = &H8003&
    HALG_SHA1 = &H8004&
    HALG_SHA2_256 = &H800C&
    HALG_SHA2_384 = &H800D&
    HALG_SHA2_512 = &H800E&
End Enum

Public Enum TransEngine
    Google = 10
    Bing = 11
    Yandex = 12
End Enum

Private Const HP_HASHSIZE As Long = &H4&
Private Const HP_HASHVAL As Long = &H2&

Public LastTranslation As String

'I try to use byref wherever possible to prevent VB from having to
'make copies of strings continuously just to process them.
'it should actually lead to a small speed increase,
'but you have to be careful to not modify the original string

''' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long
    
' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Catch empty or null input string
    Utf8BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function

Public Function Translate(ByRef Text As String, Optional Engine As String) As String
Dim temp As String
Dim MD5 As String

If transCol Is Nothing Then Set transCol = New Collection

If frmMain.sckTranslate.State <> sckConnected Then
MsgBox "Not connected to Translation server"
Translate = Text
Exit Function
End If

temp = Text
temp = Replace(temp, vbCr, "\r")
temp = Replace(temp, vbLf, "\n")
temp = Replace(temp, vbNewLine, "\r\n")
MD5 = BytesToHex(HashStringA(Text))
temp = Engine + MD5 + temp + vbCrLf
frmMain.sckTranslate.SendData Utf8BytesFromString(temp)

Do While Exists(transCol, MD5) = False
    DoEvents
    Sleep 10
    If frmMain.sckTranslate.State <> sckConnected Then
        Translate = Text
        Exit Do
    End If
Loop

Translate = ReadFromCache(MD5, transCol)
transCol.Remove MD5

End Function

Public Function ReadFromCache(ByRef strHash As String, ByRef col As Collection) As String
Dim temp() As String

If col.Count = 0 Then ReadFromCache = "": Exit Function

If Exists(col, strHash) = True Then ReadFromCache = col.Item(strHash)(1)

End Function

Public Sub AddToCache(ByRef strHash As String, ByRef Translate As String, ByRef transCol As Collection)
On Error Resume Next
Dim temp(1) As String

buildArray strHash, Translate, temp
transCol.Add temp, temp(0)

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
Debug.Print "Char: " & Left(temp, 1)
Debug.Print "Ascw: " & AscW(Left(temp, 1))

If AscW(Left(temp, 1)) = 17233 Then 'Compressed data
    If frmMain.chkCompression.Value = vbUnchecked Then
        If MsgBox("The data appears to be compressed but you've deselected compression!" & _
        vbNewLine & "Would you like to re-enable compression?", vbYesNo) = vbYes Then
            frmMain.chkCompression.Value = vbChecked
        End If
    End If
ElseIf AscW(Left(temp, 1)) = 1940 Then 'uncompressed

    If frmMain.chkCompression.Value = vbChecked Then
        If MsgBox("The data appears to be uncompressed but you've selected compression!" & _
        vbNewLine & "Would you like to disable compression?", vbYesNo) = vbYes Then
            frmMain.chkCompression.Value = vbUnchecked
        End If
    End If
End If

If NotNull = False Then Exit Sub

If frmMain.chkCompression.Value = vbChecked Then temp = Decompress(temp, bfFail)

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
If frmMain.chkCompression.Value = vbChecked Then
    tempOut = Compress(buffer.ReadBytes(buffer.length))
Else
    tempOut = buffer.ReadBytes(buffer.length)
End If

    Open Path For Binary As #NF
    Put #NF, , tempOut
    Close #NF
    
lastSave = langCol.Count
End Sub

Private Sub buildArray(ByRef key As String, ByRef Text As String, ByRef myArr() As String)

If LenB(Text) <= 0 Then
    Debug.Print
End If

myArr(0) = key
myArr(1) = Text

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


Public Function HashBytes(ByRef Data() As Byte, Optional ByVal HashAlgorithm As HashAlgo = HALG_MD5) As Byte()
Dim hProv As Long
Dim hHash As Long
Dim Hash() As Byte
Dim HashSize As Long
 
CryptAcquireContext hProv, vbNullString, vbNullString, 24, CRYPT_VERIFYCONTEXT
CryptCreateHash hProv, HashAlgorithm, 0, 0, hHash
CryptHashData hHash, Data(0), UBound(Data) + 1, 0
CryptGetHashParam hHash, HP_HASHSIZE, HashSize, 4, 0
ReDim Hash(HashSize - 1)
CryptGetHashParam hHash, HP_HASHVAL, Hash(0), HashSize, 0
CryptDestroyHash hHash
CryptReleaseContext hProv, 0
 
HashBytes = Hash()
End Function
 
Public Function HashStringA(ByVal Text As String, Optional ByVal LocaleID As Long, Optional ByVal HashAlgorithm As HashAlgo = HALG_MD5) As Byte()
Dim Data() As Byte
Data() = StrConv(Text, vbFromUnicode, LocaleID)
HashStringA = HashBytes(Data, HashAlgorithm)
End Function
 
Public Function HashStringU(ByVal Text As String, Optional ByVal HashAlgorithm As HashAlgo = HALG_MD5) As Byte()
Dim Data() As Byte
Data() = Text
HashStringU = HashBytes(Data, HashAlgorithm)
End Function
 
Public Function HashArbitraryData(ByVal MemAddress As Long, ByVal ByteCount As Long, Optional ByVal HashAlgorithm As HashAlgo = HALG_MD5) As Byte()
Dim Data() As Byte
ReDim Data(ByteCount - 1)
CopyMemory Data(0), ByVal MemAddress, ByteCount
HashArbitraryData = HashBytes(Data, HashAlgorithm)
End Function

Public Function BytesToHex(ByRef Bytes() As Byte) As String
Dim HexStringLen As Long
Dim HexString As String
 
CryptBinaryToString Bytes(0), UBound(Bytes) + 1, 12, vbNullString, HexStringLen
HexString = String$(HexStringLen - 1, vbNullChar)
CryptBinaryToString Bytes(0), UBound(Bytes) + 1, 12, HexString, HexStringLen
 
BytesToHex = Replace(UCase$(HexString), vbNewLine, "")
End Function
 
Public Function BytesToB64(ByRef Bytes() As Byte) As String
Dim B64StringLen As Long
Dim B64String As String
 
CryptBinaryToString Bytes(0), UBound(Bytes) + 1, 1, vbNullString, B64StringLen
B64String = String$(B64StringLen, vbNullChar)
CryptBinaryToString Bytes(0), UBound(Bytes) + 1, 1, B64String, B64StringLen
 
BytesToB64 = B64String
End Function

