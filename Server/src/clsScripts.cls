VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NO USADO
Option Explicit

Private Type Define
    sVari As String
    sValue As String
End Type

Public WithEvents SControl      As ScriptControl
Attribute SControl.VB_VarHelpID = -1
Private sAllCode()              As String
Private sGlobals()              As String
Private sSubs()                 As String
Private sFunctions()            As String
Public p_colSubs                As Collection
Public p_colFuncs               As Collection

' Reads code into memory and initalizes global settings.
Public Function ReadInCode(sFile As String, sModTitle As String, msc As ScriptControl)
    Dim sTemp As String
    Dim iTemp As Integer
    Dim sTotalTemp As String
    Dim sTempCode() As String
    Dim iFileNum As Integer
    Dim sDefines() As Define
    Dim Path As String

    Path = App.Path & "\Scripts"

    Erase sAllCode
    ReDim sDefines(0)
    ReDim sGlobals(0)

    iFileNum = FreeFile

    Open sFile For Input As iFileNum
    Do Until EOF(iFileNum)
        Line Input #iFileNum, sTemp

        sTemp = Trim$(Replace(sTemp, vbTab, vbNullString))

        If LenB(sTemp) <> 0 Then
            If left$(sTemp, 1) <> "#" And LCase$(left$(sTemp, 6)) <> "public" And LCase$(left$(sTemp, 7)) <> "private" Then
                sTotalTemp = sTotalTemp & sTemp & vbNewLine
            Else
                If LCase$(left$(sTemp, 8)) = "#include" Then
                    sTemp = Mid$(sTemp, InStr(sTemp, "<") + 1, Len(sTemp) - InStr(sTemp, "<") - 1)
                    sTemp = ReturnStringFromFile(Path & "\" & sTemp)
                    sTotalTemp = sTemp & vbNewLine & sTotalTemp
                ElseIf LCase$(left$(sTemp, 7)) = "#define" Then
                    sTemp = right$(sTemp, Len(sTemp) - 8)
                    sDefines(UBound(sDefines)).sVari = Mid$(sTemp, 2, InStr(sTemp, "> <") - 2)
                    sDefines(UBound(sDefines)).sValue = Mid$(sTemp, InStr(sTemp, "> <") + 3, Len(sTemp) - InStr(sTemp, "> <") - 3)
                    ReDim Preserve sDefines(UBound(sDefines) + 1)
                ElseIf LCase$(left$(sTemp, 6)) = "public" Or LCase$(left$(sTemp, 7)) = "private" Then
                    sGlobals(UBound(sGlobals)) = RTrim$(sTemp)
                    ReDim Preserve sGlobals(UBound(sGlobals) + 1)
                End If
            End If
        End If
    Loop

    Close iFileNum

    If UBound(sDefines) <> 0 Then
        ReDim Preserve sDefines(UBound(sDefines) - 1)
    End If
        
    If UBound(sGlobals) <> 0 Then
        ReDim Preserve sGlobals(UBound(sGlobals) - 1)
    End If

    For iTemp = 0 To UBound(sDefines)
        sTotalTemp = Replace(sTotalTemp, sDefines(iTemp).sVari, sDefines(iTemp).sValue)
    Next

    sAllCode = Split(sTotalTemp, vbNewLine)

    ' Get rid of last array element (which is blank).
    ReDim Preserve sAllCode(UBound(sAllCode) - 1)

    ' Split string into collection.
    GetSubs sAllCode
    GetFunctions sAllCode

    ' Put collection into the script control.
    
    msc.Modules.Add sModTitle
    AddGlobalsToCode msc, sModTitle
    AddSubsToCode msc, sModTitle
    AddFuncsToCode msc, sModTitle
End Function

' Returns a string from an #include file.
Private Function ReturnStringFromFile(sFile As String) As String
    Dim sTotalTemp As String
    Dim iFileNum As Integer
    Dim sTemp As String
    Dim Path As String

    On Local Error GoTo FileNotFound

    Path = App.Path & "\Scripts"

    iFileNum = FreeFile

    Open sFile For Input As iFileNum
    Err.Clear

    Do Until EOF(iFileNum)
        Line Input #iFileNum, sTemp

        sTemp = Trim$(Replace(sTemp, vbTab, vbNullString))

        If LenB(sTemp) <> 0 Then
            If left$(sTemp, 1) <> "#" Then
                sTotalTemp = sTotalTemp & sTemp & vbNewLine
            Else
                If LCase$(left$(sTemp, 8)) = "#include" Then
                    sTemp = Mid$(sTemp, InStr(sTemp, "<") + 1, Len(sTemp) - InStr(sTemp, "<") - 1)
                    sTemp = ReturnStringFromFile(Path & "\" & sTemp)
                    sTotalTemp = sTemp & vbNewLine & sTotalTemp
                End If
            End If
        End If
    Loop

    Close iFileNum

    ReturnStringFromFile = sTotalTemp

    Exit Function

FileNotFound:
End Function

Private Sub AddGlobalsToCode(mscControl As ScriptControl, sModName As String)
    Dim iCount As Integer

    On Error Resume Next

    For iCount = 0 To UBound(sGlobals)
        mscControl.Modules(sModName).AddCode sGlobals(iCount)
    Next
End Sub

Private Function GetSubs(sCode() As String)
    Dim iCount As Integer
    Dim iTemp As Integer
    Dim sTitle As String
    Dim sSub As String

    Set p_colSubs = New Collection

    For iCount = 0 To UBound(sCode)
        sSub = vbNullString

        If LCase$(left$(sCode(iCount), 3)) = "sub" Then
            For iTemp = 5 To Len(sCode(iCount))
                If Mid$(sCode(iCount), iTemp, 1) = "(" Then
                    sTitle = Mid$(sCode(iCount), 5, iTemp - 5)
                    Exit For
                End If
            Next

            Do Until LCase$(sCode(iCount)) = "end sub"
                sSub = sSub & sCode(iCount) & vbNewLine
                iCount = iCount + 1
            Loop

            sSub = sSub & sCode(iCount)

            On Error Resume Next

            p_colSubs.Add sSub, sTitle
            Err.Clear
        End If
    Next
End Function

Private Function AddSubsToCode(mscControl As ScriptControl, sModName As String)
    Dim iCount As Integer

    On Error Resume Next

    For iCount = 1 To p_colSubs.Count
        mscControl.Modules(sModName).AddCode p_colSubs(iCount)
    Next
End Function

Private Function GetFunctions(sCode() As String)
    Dim iCount As Integer
    Dim iTemp As Integer
    Dim sTitle As String
    Dim sFunc As String

    Set p_colFuncs = New Collection

    For iCount = 0 To UBound(sCode)
        sFunc = vbNullString

        If LCase$(left$(sCode(iCount), 8)) = "function" Then
            For iTemp = 10 To Len(sCode(iCount))
                If Mid$(sCode(iCount), iTemp, 1) = "(" Then
                    sTitle = Mid$(sCode(iCount), 10, iTemp - 10)
                    Exit For
                End If
            Next

            Do Until LCase$(sCode(iCount)) = "end function"
                sFunc = sFunc & sCode(iCount) & vbNewLine
                iCount = iCount + 1
            Loop

            sFunc = sFunc & sCode(iCount)

            On Error Resume Next

            p_colFuncs.Add sFunc, sTitle
            Err.Clear
        End If
    Next
End Function

Private Function AddFuncsToCode(mscControl As ScriptControl, sModName As String)
    Dim iCount As Integer

    On Error Resume Next

    For iCount = 1 To p_colFuncs.Count
        mscControl.Modules(sModName).AddCode p_colFuncs(iCount)
    Next
End Function

Private Sub Class_initialize()
    Set SControl = New ScriptControl
    SControl.Language = "vbScript"
End Sub

Private Sub Class_Terminate()
    Set SControl = Nothing
    Set p_colFuncs = Nothing
    Set p_colSubs = Nothing
End Sub

'PRIVATE BECAUSE I CANT GET IT TO WORK
Private Function RunCodeReturn(sModule As String, sCode As String, ParamArray abc() As Variant) As Variant
    Dim sError As String
    Dim aTemp() As String
    Dim bTemp As Variant

    'How the heck to pass a paramarray array to another function using a
    'paramarray array?? if ya know email me andyd@vuetura.com

    If IsMissing(abc) Then
        bTemp = SControl.Modules(sModule).Run(sCode)
    Else
        bTemp = SControl.Modules(sModule).Run(sCode, abc)
    End If

        RunCodeReturn = bTemp
End Function

Public Function ExecuteStatement(sModule As String, sCode As String)
    On Error Resume Next

    SControl.Modules(sModule).ExecuteStatement sCode
    
End Function

Private Sub SControl_Error()
    Dim sError As String

    If val(GetVar(App.Path & "\Data.ini", "CONFIG", "ScriptErrors")) = 1 Then
        MsgBox "TYPE: " & SControl.error.Description & vbNewLine & "LINE: " & SControl.error.Line & vbNewLine & "COLUMN: " & SControl.error.Column & vbNewLine & "CODE: " & vbNewLine & "    " & SControl.error.Text
    End If

    Err.Clear
End Sub

