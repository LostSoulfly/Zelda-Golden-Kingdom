Attribute VB_Name = "modCodes"
Option Explicit
Public IPTries As Collection



Public Const CURRENCY_NAME As String = "Tingles"

Private Const CODE_PATH As String = "/Data/Codes"
Private Const CODE_FILE_NAME As String = "/Codes.txt"
Private Const CODE_LOG As String = "/Log.txt"
Private Const USEDCODES_FILE_NAME As String = "/UsedCodes.txt"

'NO TOCAR
Private Const FIRST_PURSHASE As Long = 100
Private Const SECOND_PURSHASE As Long = 700
Private Const THIRD_PURSHASE As Long = 3000
Private Const MAX_PURSHASE_TYPES As Byte = 3

Private Const key As String = "&H329BA92D"

Public Sub ClearIPTries()
    Set IPTries = New Collection
End Sub

Public Function IPExists(ByVal IP As String) As Long
    Dim i As Long
    For i = 1 To IPTries.Count
        If IPTries.item(i) = IP Then
            IPExists = i
            Exit Function
        End If
    Next
End Function


Public Function GetPlayerBonusPoints(ByVal index As Long) As Long
    GetPlayerBonusPoints = player(index).BonusPoints
End Function

Public Sub SetPlayerBonusPoints(ByVal index As Long, ByVal points As Long)
    player(index).BonusPoints = points
End Sub

Public Sub CheckCode(ByVal index As Long, ByVal code As String)
    Dim parsedcode As String
    If Not Encriptate(code, parsedcode) Then Exit Sub
    
    If IPExists(GetPlayerIP(index)) Then
        PlayerMsg index, "Error", BrightRed
    Else
        Dim points As Long
        If VerifyCode(parsedcode, points) Then
            AddCodeLog index, parsedcode, points
            AddPlayerBonusPoints index, points
            SavePlayer index
            AddInFile "Success", App.Path & CODE_PATH & CODE_LOG
        Else
            IPTries.Add GetPlayerIP(index)
            PlayerMsg index, "Error", BrightRed
        End If
    End If
End Sub


Function FindInFile(ByVal Line As String, ByVal search_file As String)
    Dim F As Long
    Dim S As String
    F = FreeFile
    Open search_file For Input As #F

    Do While Not EOF(F)
        Input #F, S
        If S = Line Then
            FindInFile = True
            Close #F
            Exit Function
        End If
    Loop

    Close #F
End Function


Sub AddInFile(ByVal Line As String, ByVal search_file As String)
    Dim F As Long
    F = FreeFile
    Open search_file For Append As #F
    Print #F, Line
    Close #F
End Sub

Function PurshaseTypeToPoints(ByVal PurshaseType As Byte) As Long
    Select Case PurshaseType
    Case 0
        PurshaseTypeToPoints = FIRST_PURSHASE
    Case 1
        PurshaseTypeToPoints = SECOND_PURSHASE
    Case 2
        PurshaseTypeToPoints = THIRD_PURSHASE
    End Select
End Function

Function IsHexChar(ByVal C As String) As Long
    If IsNumeric(C) Then
        IsHexChar = CLng(C)
    Else
        Select Case C
        Case "a"
            IsHexChar = 10
        Case "b"
            IsHexChar = 11
        Case "c"
            IsHexChar = 12
        Case "d"
            IsHexChar = 13
        Case "e"
            IsHexChar = 14
        Case "f"
            IsHexChar = 15
        Case Else
            IsHexChar = -1
        End Select
    End If
End Function

Function HexStringToLong(ByVal S As String) As Long
    If Not IsNumeric("&H" & S) Then Exit Function
    HexStringToLong = CLng("&H" & S)
End Function

Function VerifyCode(ByVal parsedcode As String, ByRef points As Long) As Boolean

    
    
    Dim search_file_used As String
    search_file_used = App.Path & CODE_PATH & USEDCODES_FILE_NAME
    
    
    If Not FileExist(search_file_used, True) Then Exit Function
    If FindInFile(parsedcode, search_file_used) Then Exit Function
    
    
    
    Dim N As Long
    N = HexStringToLong(parsedcode)
    If N = -1 Then Exit Function
    
    Dim ct As Byte
    ct = Abs(N) Mod MAX_PURSHASE_TYPES
    
    If ct > MAX_PURSHASE_TYPES Then Exit Function
    
    points = PurshaseTypeToPoints(ct) 'TODO
     
    Dim search_file As String
    If Not FileExist(search_file, True) Then Exit Function
    search_file = App.Path & CODE_PATH & CODE_FILE_NAME
    
    If FindInFile(parsedcode, search_file) Then
        Call AddInFile(parsedcode, search_file_used)
        'Delete?
        VerifyCode = True
    End If
    
    
End Function

Sub AddCodeLog(ByVal index As Long, ByVal code As String, ByVal points As Long)
    Dim CurDate As Date
    CurDate = DateValue(Now)
    Dim DateString As String
    DateString = Format(CurDate, "yyyy:mm:dd")
    AddInFile DateString & ", " & GetPlayerLogin(index) & ", " & code & ", " & points, App.Path & CODE_PATH & CODE_LOG
End Sub

Private Function Encriptate(ByVal X As String, ByRef Out As String) As Boolean
    If Not IsNumeric("&H" & X) Then Exit Function
    If Len(X) > 8 Then Exit Function
    Out = Hex(CLng("&H" & X) Xor CLng(key))
    Encriptate = True
End Function

Sub AddPlayerBonusPoints(ByVal index As Long, ByVal points As Long)
    Dim curpoints As Long
    curpoints = GetPlayerBonusPoints(index)
    
    SetPlayerBonusPoints index, curpoints + points
    PlayerMsg index, GetTranslation("Has ganado:", , UnTrimBack) & points & " " & CURRENCY_NAME & "!", BrightGreen, , False
    SendPlayerBonusPoints index
End Sub

Sub SendPlayerBonusPoints(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SBonusPoints
    buffer.WriteLong GetPlayerBonusPoints(index)
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub



