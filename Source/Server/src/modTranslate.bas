Attribute VB_Name = "modTranslate"
Option Explicit
Private Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)

'Save the collection every 50 translations
Private Const saveEvery = 50
Private Const intTransPerSec As Integer = 50
'Sleep the server every 1ms until it's able to translate again.
Public Const blWaitToTranslate As Boolean = True

'these could be const, but I didn't want them to be.
Public LangTo As String
Public LangFrom As String
Public strTransPath As String
Public strOrigPath As String

Public Enum UnTrimType

    UnTrimFront = 1
    UnTrimBack = 2
    UnTrimBoth = 3

End Enum

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

Public isLocked As Boolean

Private currentTranslation As String
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
    Google = 0
    Bing = 1
    Yandex = 2
End Enum

Private Const HP_HASHSIZE As Long = &H4&
Private Const HP_HASHVAL As Long = &H2&

'I try to use byref wherever possible here to prevent VB from having to
'make copies of strings continuously just to process them.
'However, I realized that I was trimming the text, which might not always
'be what the server wants to happen!
'it should actually lead to a small speed increase,
'but you have to be careful to not modify the original string.

Public Function GetTranslation(ByRef Text As String, Optional transLock As Boolean, Optional UnTrim As UnTrimType) As String
'weed out easy stuff
Dim txtTemp As String
txtTemp = Trim$(Text) 'if there are spaces on the ends, trim them
If LenB(txtTemp) <= 1 Then GetTranslation = txtTemp: Exit Function 'if the length of the string is <=1 then we aren't translating it.
If IsNumeric(txtTemp) = True Then GetTranslation = Text: Exit Function 'if it's a number.. we aren't translating it.


'I didn't feel like creating a translation queue or something of the sort,
'so this is the simple method to prevent translating the same thing multiple times while
'a translation is pending. Especially an issue when showing item descriptions/names.
If isLocked = True Then
    AddTransLog "Trying to translate when transLock is active!"
    Exit Function
End If

If transLock = True Then isLocked = True

    If txtTemp = currentTranslation Then
    
        AddTransLog "Trying to translate the same thing multiple times!"
        Exit Function
    End If
    
currentTranslation = txtTemp

'some loaded data will contain nullchars, which can waste time translating.
If InStr(1, txtTemp, vbNullChar, vbBinaryCompare) <> 0 Then
    txtTemp = Replace(txtTemp, vbNullChar, "") 'replace the nullchars.
    currentTranslation = vbNullString
    isLocked = False
    If LenB(txtTemp) <= 1 Then GetTranslation = txtTemp: Exit Function 'if the new length is too short, we're not translating it.
End If

'get the actual translation (either from cache, or from a translation service)
GetTranslation = Translate(txtTemp)

'currentTranslation = vbNullString
'isLocked = False

'checking for a new line in the text first is faster than
'simply running the replace on every translation
If InStr(1, GetTranslation, "\r\n", vbBinaryCompare) <> 0 Then
'Some lines in this game require this to look right!
    GetTranslation = Replace(GetTranslation, "\r\n", vbNewLine)
End If

Select Case UnTrim

    Case UnTrimType.UnTrimFront
        GetTranslation = " " & GetTranslation
    
    Case UnTrimType.UnTrimBack
        GetTranslation = GetTranslation & " "
    
    Case UnTrimType.UnTrimBoth
        GetTranslation = " " & GetTranslation & " "

End Select

'release the lock and reset current translation
currentTranslation = vbNullString
isLocked = False

End Function

Public Sub AddTransLog(Text As String)
With frmTransLog.txtLog

If frmTransLog.txtLog.Visible = False Then Exit Sub

    .SelText = vbCrLf & Time & ": " & Text
    '.Text = .Text & vbCrLf & Time & ": " & Text
End With
End Sub


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
Public Function TranslateNew(ByRef Text As String, Engine As String) As String
Dim Temp As String
Dim MD5 As String

If transCol Is Nothing Then Set transCol = New Collection

If frmTransLog.sckTranslate.state <> sckConnected Then
    AddTransLog "Not connected to Translation server"
    TranslateNew = Text
    frmTransLog.sckTranslate.Close
    frmTransLog.sckTranslate.Connect
    Sleep 200
End If

If IsNumeric(Text) Then TranslateNew = Text
If IsNumeric(Replace(Text, "/", "")) Then TranslateNew = Text: Exit Function
If IsNumeric(Replace(Text, "\", "")) Then TranslateNew = Text: Exit Function
If Len(Text) = 1 Then TranslateNew = Text: Exit Function

TranslateNew = Text
Temp = Text
Temp = Replace(Temp, vbCr, "\r")
Temp = Replace(Temp, vbLf, "\n")
Temp = Replace(Temp, vbNewLine, "\r\n")
'Temp = StripAccents(Temp)
MD5 = BytesToHex(HashStringA(Text))
Temp = Engine + MD5 + Temp + vbCrLf
If (frmTransLog.sckTranslate.state <> sckConnected) Then Exit Function
frmTransLog.sckTranslate.SendData Utf8BytesFromString(Temp)

Dim timeout As Integer
Do While Exists(transCol, MD5) = False
    If timeout > 500 Then
        frmTransLog.sckTranslate.SendData Temp
        timeout = 0
    End If
    DoEvents
    Sleep 10
    If frmTransLog.sckTranslate.state <> sckConnected Then
        TranslateNew = Text
        Exit Do
    End If
    timeout = timeout + 1
Loop

TranslateNew = ReadFromCache(MD5, transCol)
'transCol.Remove MD5

End Function

Private Function StripAccents(str)

Dim accent As String, noaccent As String, k As Integer, o As Integer, currentChar As String, result As String

accent = "àèìòùÀÈÌÒÙäëïöüÄËÏÖÜâêîôûÂÊÎÔÛáéíóúÁÉÍÓÚðÐýÝãñõÃÑÕšŠžŽçÇåÅøØ"
noaccent = "aeiouAEIOUaeiouAEIOUaeiouAEIOUaeiouAEIOUdDyYanoANOsSzZcCaAoO"

currentChar = ""
result = ""
k = 0
o = 0

For k = 1 To Len(str)
    currentChar = Mid(str, k, 1)
    o = InStr(accent, currentChar)
    If o > 0 Then
        result = result & Mid(noaccent, 0, 1)
    Else
        result = result & currentChar
    End If
Next

StripAccents = result

End Function

Private Function Translate(ByRef Text As String) As String
Dim strTranslation As String
Dim strHash As String
Dim SleepTime As Long
Dim i As Long


'init the collection for the lang if it doesn't exist (and load our current language)
If langCol Is Nothing Then Set langCol = New Collection: loadLang strTransPath, langCol
If origCol Is Nothing Then Set origCol = New Collection: loadLang strOrigPath, origCol

'get the md5 of our current string
strHash = BytesToHex(HashStringA(Text))
'read from the file to see if it's already been translated
strTranslation = ReadFromCache(strHash, langCol)
'If the length of it is 0, translate it.
If LenB(strTranslation$) = 0 Then

StartOver:
    'Check to see if we're over the timer, if we are, reset it.
    If GetRealTickCount > tmrTrans Then
        tmrTrans = GetRealTickCount + 1000
        TransCount = 0
    End If
    
    'check to see if we can do more translations this second
        If TransCount + 1 > intTransPerSec Then
            'too many translations. bail. or wait.
            If blWaitToTranslate = True Then
                AddTransLog "Over translate quota. Sleeping.. "
                'calculate the sleeptime from now to when we can do another translation
                SleepTime = (tmrTrans) - (GetRealTickCount + 1)
                'sleep 1ms every iteration between 0 and sleeptime
                For i = 0 To SleepTime
                    'this should prevent problems for players from small lag maybe?
                    Sleep 1
                    DoEvents
                Next i
                'go back up a bit and try again.
                GoTo StartOver
            Else
                'blWaitToTranslate is false. Return untranslated text.
                'This is the best setting for a populated server, as otherwise
                'it would slow down a bit and lag for people.
                AddTransLog "Skipping translation; over quota.."
                Translate = Text
                Exit Function
            End If
        Else
        
        'russian roulette! WHOEVER WINS, GETS TO TRANSLATE FOR US! (not truly random :o)
        Select Case RAND(0, 2)

        Case Is = 0
            Translate = TranslateNew(Text, 11)
            AddTransLog "Translated(Bing): [" & Text & "] to [" & Translate & "]"
        Case Is = 1
            'Translate = TranslateNew(text, 0)
            'AddTranslog "Translated(Google): [" & text & "] to [" & Translate & "]"
            Translate = TranslateNew(Text, 12)
            AddTransLog "Translated(Yandex): [" & Text & "] to [" & Translate & "]"
        Case Is = 2
            Translate = TranslateNew(Text, 12)
            AddTransLog "Translated(Yandex): [" & Text & "] to [" & Translate & "]"
        End Select
        
        'for now, if it's blank, just return the original text. However, this means that a translation error happened most likely.
        If LenB(Translate) <= 1 Then
        Translate = Text
        Exit Function
        End If
        
        'check that an error didn't occur.. log it to server log?
            AddToCache strHash, Translate, langCol ', Text
            AddToCache strHash, Text, origCol
        'uncomment this to save the collection every time a new translation is made, but be careful as it could get slow..
            'saveLang strTransPath, langCol
            'saveLang strOrigPath, origCol
        'increase the number of translations currently
            TransCount = TransCount + 1
        
        End If

Else
Translate = strTranslation
If Exists(origCol, strHash) = False Then
'if it's not in the cache, let's add it.
'This shouldn't happen, but I didn't have a separate collection
'for the original untranslated text.
    AddToCache strHash, Text, origCol
End If

AddTransLog "Cached: [" & strTranslation & "] original: [" & Text & "]"

End If

If (lastSave + saveEvery) < (langCol.Count) And (lastSave + saveEvery) < (origCol.Count) Then
    saveLang strTransPath, langCol
    saveLang strOrigPath, origCol
End If
'Set T = Nothing

End Function

Public Function ReadFromCache(ByRef strHash As String, ByRef col As Collection) As String
Dim Temp() As String

If col.Count < 1 Then ReadFromCache = "": Exit Function

If Exists(col, strHash) = True Then ReadFromCache = col.item(strHash)(1)

End Function

Public Sub AddToCache(ByRef strHash As String, ByRef Translate As String, ByRef col As Collection)
On Error Resume Next
Dim Temp(1) As String

buildArray strHash, Translate, Temp
col.Add Temp, Temp(0)

End Sub

Public Sub loadLang(Path As String, ByRef col As Collection)
'Dim strTemp As String
Dim Temp() As Byte
Dim tempArray(1) As String
Dim Buffer As New clsBuffer
Dim lngBufferCount As Long
Dim NF As Integer
Dim NotNull As Boolean
Dim bfFail As Boolean
NF = FreeFile

AddTransLog "Loading Lang file: " & Path

' check exists
    Open Path For Binary As NF
    Close NF

Temp = ReadFile(Path, NotNull)

If NotNull = False Then Exit Sub

    Temp = Decompress(Temp, bfFail)
    If bfFail = True Then GoTo skip
    Buffer.WriteBytes Temp
    
lngBufferCount = Buffer.ReadLong
lastSave = lngBufferCount
Dim i As Long
For i = 1 To lngBufferCount
    buildArray Buffer.ReadString, Buffer.ReadString, tempArray
    col.Add tempArray, tempArray(0)
Next

'set our lastSave variable so we can save again in 50 translations
lastSave = col.Count

skip:
Set Buffer = Nothing
End Sub

Public Sub saveLang(Path As String, ByRef col As Collection, Optional blForceSave As Boolean = False)
Dim NF As Integer
Dim tempOut() As Byte
Dim Buffer As New clsBuffer
Dim i As Long
NF = FreeFile

If col Is Nothing Then Exit Sub

If blForceSave = False Then If (lastSave) = (langCol.Count) Then Exit Sub
AddTransLog "Saving lang to: " & Path
Buffer.WriteLong col.Count

For i = 1 To col.Count
'write the key first
    Buffer.WriteString (col.item(i)(0))
'write the actual translation
    Buffer.WriteString (col.item(i)(1))
Next

'write buffer to temp out
tempOut = Compress(Buffer.ReadBytes(Buffer.length))

    Open Path For Binary As #NF
    Put #NF, , tempOut
    Close #NF
    
lastSave = langCol.Count
End Sub

Private Sub buildArray(ByRef key As String, ByRef Text As String, ByRef myArr() As String)

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
Dim Buffer As New clsBuffer
NF = FreeFile

Dim col As Collection
Set col = New Collection

Dim myArr(1) As String
col.Add Array("0a-6d-d0-dd-a2-ee-52-6b-57-55-6b-68-97-33-4a-b1", "Level 40-50"), "0a-6d-d0-dd-a2-ee-52-6b-57-55-6b-68-97-33-4a-b1"
col.Add Array("da-f8-9b-9c-2d-b8-51-d6-91-84-f0-95-6a-44-a0-d3", "Level 5-10"), "da-f8-9b-9c-2d-b8-51-d6-91-84-f0-95-6a-44-a0-d3"
col.Add Array("7e-54-5b-21-c5-a5-49-23-4b-ca-43-23-32-cf-54-82", "Level 10-20"), "7e-54-5b-21-c5-a5-49-23-4b-ca-43-23-32-cf-54-82"
Buffer.WriteLong col.Count

Dim i As Long


For i = 1 To col.Count
'write the key first
    Buffer.WriteString (col.item(i)(0))
'write the actual translation
    Buffer.WriteString (col.item(i)(1))
Next i

'write buffer to temp out
tempOut = Compress(Buffer.ReadBytes(Buffer.length))

    Open Path For Binary As #NF
    Put #NF, , tempOut
    Close #NF

Set col = Nothing

End Sub

Public Function Exists(ByRef col As Collection, ByRef index) As Boolean
On Error GoTo ExistsTryNonObject
    Dim o As Object

    Set o = col(index)
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


