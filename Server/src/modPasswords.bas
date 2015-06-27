Attribute VB_Name = "modPasswords"
Option Explicit


Public Function EncriptatePassword(ByVal password As String) As String

Dim i As Byte
Dim PasLen As Byte
Dim newpass() As String
Dim med As Byte

PasLen = Len(Trim$(password))

If PasLen <= 0 Then Exit Function

newpass = SubCadenas(password, 1)


For i = 0 To PasLen - 1
    If i = PasLen - 1 Then
        newpass(i) = CharSummation(newpass(i), newpass(0))
    Else
        newpass(i) = CharSummation(newpass(i), newpass(i + 1))
    End If
    
Next


EncriptatePassword = Join(newpass, "")

End Function

Public Function DesEncriptatePassword(ByVal password As String) As String

Dim i As Integer
Dim PasLen As Byte
Dim newpass() As String

PasLen = Len(Trim$(password))

If PasLen <= 0 Then Exit Function

newpass = SubCadenas(password, 1)

i = CInt(PasLen) - 1
Do While i <> -1
            If i = PasLen - 1 Then
                newpass(i) = CharSubstract(newpass(i), newpass(0))
            Else
                newpass(i) = CharSubstract(newpass(i), newpass(i + 1))
            End If

    
i = i - 1
Loop


DesEncriptatePassword = Join(newpass, "")

End Function

Private Function SubCadenas(ByVal Cadena As String, Bytes As Integer) As String()
  Dim Temp() As String
  Dim i As Integer
  Do While Cadena <> ""
    ReDim Preserve Temp(i)
    Temp(i) = left$(Cadena, Bytes)
    i = i + 1
    Cadena = Mid$(Cadena, Bytes + 1)
  Loop
  SubCadenas = Temp
End Function

Private Function CharSummation(ByVal char1 As String, ByVal char2 As String) As String
Dim bits As Long

bits = 255

If Asc(char1) + Asc(char2) > bits Then
    CharSummation = Chr$(Asc(char1) + Asc(char2) - bits)
Else
    CharSummation = Chr$(Asc(char1) + Asc(char2))
End If

End Function

Private Function CharSubstract(ByVal char1 As String, ByVal char2 As String) As String
Dim bits As Long

bits = 255


If Asc(char1) - Asc(char2) < 0 Then
    CharSubstract = Chr$(Asc(char1) - Asc(char2) + bits)
Else
    CharSubstract = Chr$(Asc(char1) - Asc(char2))
End If

End Function
