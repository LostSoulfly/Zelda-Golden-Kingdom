Attribute VB_Name = "modQuestions"
Option Explicit

Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Public Const MAX_QUESTIONS As Byte = 15
Public Const QUESTION_WAIT_TIME As Byte = 30
Public Enum QuestionType
    WarpMeTo = 1
    WarpToMe = 2
End Enum

Type QuestionRec
    InUse As Boolean
    EndTime As Long
    Questioner As Long
    Respondent As Long
    QuestionType As QuestionType
End Type

Public Questions(1 To MAX_QUESTIONS) As QuestionRec


Public Sub AddQuestion(ByVal Questioner As Long, ByVal Respondent As Long, ByVal qt As QuestionType)

    Dim i As Byte
    i = FindOpenQuestionSlot
    If i > 0 Then
        With Questions(i)
            .InUse = True
            .Questioner = Questioner
            .QuestionType = qt
            .Respondent = Respondent
            .EndTime = GetRealTickCount + QUESTION_WAIT_TIME * 1000
        End With
        SendQuestion i
    End If
End Sub

Public Sub SendQuestion(ByVal question As Byte)
    If Not QuestionInUse(question) Then Exit Sub
    Dim N As Long, m As Long
    
    N = GetQuestionQuestioner(question)
    m = GetQuestionRespondent(question)
    
    Select Case GetQuestionType(question)
    Case WarpMeTo
        Call SendQuestionData(m, "teleport", GetPlayerName(N) & " se quiere teletransportar hacia ti, le dejas?")
    Case WarpToMe
        Call SendQuestionData(m, "teleport", GetPlayerName(N) & " quiere teletransportarte, le dejas?")
    End Select
End Sub

Sub SendQuestionData(ByVal index As Long, ByVal header As String, ByVal question As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SQuestion
    Buffer.WriteString header
    Buffer.WriteString question
    
    SendDataTo index, Buffer.ToArray
    Set Buffer = Nothing
End Sub

Public Function FindOpenQuestionSlot() As Byte
    Dim i As Byte
    i = 0
    For i = 1 To MAX_QUESTIONS
        If Not Questions(i).InUse Then
            FindOpenQuestionSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub ClearQuestion(ByVal question As Byte)
    If question < 1 Or question > MAX_QUESTIONS Then Exit Sub
    ZeroMemory Questions(question), Len(Questions(question))
End Sub

Public Function FindQuestionByRespondent(ByVal index As Long) As Byte
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    Dim i As Byte
    For i = 1 To MAX_QUESTIONS
        If QuestionInUse(i) Then
            If GetQuestionRespondent(i) = index Then
                FindQuestionByRespondent = i
                Exit Function
            End If
        End If
    Next
End Function

Public Sub SolveQuestion(ByVal question As Byte, ByVal Response As Boolean)
    If question < 1 Or question > MAX_QUESTIONS Then Exit Sub
    
    If Not Response Then: ClearQuestion (question)
    
    If Not QuestionInUse(question) Then Exit Sub
    
    Dim N As Long, m As Long
    
    m = GetQuestionQuestioner(question)
    N = GetQuestionRespondent(question)
    Select Case GetQuestionType(question)
    Case WarpMeTo
        Call WarpXtoY(m, N, False)
    Case WarpToMe
        Call WarpXtoY(N, m, True)
    End Select
    
    ClearQuestion question
End Sub

Public Function GetQuestionQuestioner(ByVal question As Byte) As Long
    GetQuestionQuestioner = Questions(question).Questioner
End Function

Public Function GetQuestionRespondent(ByVal question As Byte) As Long
    GetQuestionRespondent = Questions(question).Respondent
End Function

Public Function GetQuestionType(ByVal question As Byte) As QuestionType
    If question < 1 Or question > MAX_QUESTIONS Then Exit Function
    
    GetQuestionType = Questions(question).QuestionType
End Function

Public Function QuestionInUse(ByVal question As Byte) As Boolean

    If question < 1 Or question > MAX_QUESTIONS Then Exit Function
    QuestionInUse = Questions(question).InUse
End Function

Public Sub ClearQuestions()
Dim i As Byte
For i = 1 To MAX_QUESTIONS
    If QuestionInUse(i) Then
        If GetRealTickCount > Questions(i).EndTime Then
            ClearQuestion i
        End If
    End If
Next
End Sub
