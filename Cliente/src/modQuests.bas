Attribute VB_Name = "modQuests"
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////
Option Explicit

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const MAX_QUESTS As Byte = 70
Public Const MAX_QUESTS_ITEMS As Byte = 10 'Alatar v1.2
Public Const EDITOR_TASKS As Byte = 7

Public Const QUEST_TYPE_GOSLAY As Byte = 1
Public Const QUEST_TYPE_GOGATHER As Byte = 2
Public Const QUEST_TYPE_GOTALK As Byte = 3
Public Const QUEST_TYPE_GOREACH As Byte = 4
Public Const QUEST_TYPE_GOGIVE As Byte = 5
Public Const QUEST_TYPE_GOKILL As Byte = 6
Public Const QUEST_TYPE_GOTRAIN As Byte = 7
Public Const QUEST_TYPE_GOGET As Byte = 8

Public Const QUEST_NOT_STARTED As Byte = 0
Public Const QUEST_STARTED As Byte = 1
Public Const QUEST_COMPLETED As Byte = 2
Public Const QUEST_COMPLETED_BUT As Byte = 3

Public Quest_Changed(1 To MAX_QUESTS) As Boolean

'Types
Public Quest(1 To MAX_QUESTS) As QuestRec

Public Type PlayerQuestRec
    Status As Long
    ActualTask As Long
    CurrentCount As Long 'Used to handle the Amount property
End Type

'Alatar v1.2
Private Type QuestRequiredItemRec
    Item As Long
    value As Long
End Type

Private Type QuestGiveItemRec
    Item As Long
    value As Long
End Type

Private Type QuestTakeItemRec
    Item As Long
    value As Long
End Type

Private Type QuestRewardItemRec
    Item As Long
    value As Long
End Type
'/Alatar v1.2

Public Type TaskRec
    Order As Long
    NPC As Long
    Item As Long
    map As Long
    Resource As Long
    amount As Long
    Speech As String
    TaskLog As String
    QuestEnd As Boolean
End Type

Public Type QuestRec
    'Alatar v1.2
    Name As String * 30
    
    Repeat As Long
    QuestLog As String
    Speech(1 To 3) As String
    GiveItem(1 To MAX_QUESTS_ITEMS) As QuestGiveItemRec
    TakeItem(1 To MAX_QUESTS_ITEMS) As QuestTakeItemRec
    
    RequiredLevel As Long
    RequiredQuest As Long
    RequiredClass(1 To 5) As Long
    RequiredItem(1 To MAX_QUESTS_ITEMS) As QuestRequiredItemRec
    
    RewardExp As Long
    RewardItem(1 To MAX_QUESTS_ITEMS) As QuestRewardItemRec
    
    Task(1 To MAX_TASKS) As TaskRec
    
    Level As Long
    '/Alatar v1.2
 
 TranslatedName As String * 30
End Type

' ////////////
' // Editor //
' ////////////

Public Sub QuestEditorInit()
Dim i As Long
    
    If frmEditor_Quest.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Quest.lstIndex.ListIndex + 1

    With frmEditor_Quest
        'Alatar v1.2
        .txtName = Trim$(Quest(EditorIndex).Name)
        If Quest(EditorIndex).Repeat = 1 Then
            .chkRepeat.value = 1
        Else
            .chkRepeat.value = 0
        End If
        .txtQuestLog = Trim$(Quest(EditorIndex).QuestLog)
        For i = 1 To 3
            '.scrlReq(i).Value = Quest(EditorIndex).Requirement(i)
            .txtSpeech(i).text = Trim$(Quest(EditorIndex).Speech(i))
        Next
        'For i = 1 To MAX_QUESTS_ITEMS
        '    .scrlGiveItem.Value = Quest(EditorIndex).GiveItem(i).Item
        '    If Not Quest(EditorIndex).GiveItem(i).Value = 0 Then
        '        .scrlGiveItemValue.Value = Quest(EditorIndex).GiveItem(i).Value
        '    Else
        '        .scrlGiveItemValue.Value = 1
        '    End If
        '
        '    .scrlTakeItem.Value = Quest(EditorIndex).TakeItem(i).Item
        '    If Not Quest(EditorIndex).TakeItem(i).Value = 0 Then
        '        .scrlTakeItemValue.Value = Quest(EditorIndex).TakeItem(i).Value
        '    Else
        '        .scrlTakeItemValue.Value = 1
        '    End If
        '
        '    .scrlItemRew.Value = Quest(EditorIndex).RewardItem(i).Item
        '    If Not Quest(EditorIndex).RewardItem(i).Value = 0 Then
        '        .scrlItemRewValue.Value = Quest(EditorIndex).RewardItem(i).Value
        '    Else
        '        .scrlItemRewValue.Value = 1
        '    End If
        'Next
        
        .scrlReqLevel.value = Quest(EditorIndex).RequiredLevel
        .scrlReqQuest.value = Quest(EditorIndex).RequiredQuest
        For i = 1 To 5
            .scrlReqClass.value = Quest(EditorIndex).RequiredClass(i)
        Next
        
        .scrlExp.value = Quest(EditorIndex).RewardExp
        
        UpdateQuestOptimalLevel
        
        'Update the lists
        UpdateQuestGiveItems
        UpdateQuestTakeItems
        UpdateQuestRewardItems
        UpdateQuestRequirementItems
        UpdateQuestClass
        
        '/Alatar v1.2
        
        'load task nº1
        .scrlTotalTasks.value = 1
        LoadTask EditorIndex, 1
        
    End With

    Quest_Changed(EditorIndex) = True
    
End Sub

Sub UpdateQuestOptimalLevel()
    With frmEditor_Quest
    Dim lvl As Long
    lvl = Quest(EditorIndex).Level
    
    .scrlOptimalLevel.Min = 0
    .scrlOptimalLevel.Max = MAX_LEVELS
    
    If lvl >= 0 And lvl <= MAX_LEVELS Then
        .scrlOptimalLevel.value = lvl
    Else
        .scrlOptimalLevel.value = 0
    End If
    End With
End Sub

'Alatar v1.2
Public Sub UpdateQuestGiveItems()
    Dim i As Long
    
    frmEditor_Quest.lstGiveItem.Clear
    
    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).GiveItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstGiveItem.AddItem "-"
            Else
                frmEditor_Quest.lstGiveItem.AddItem Trim$(Trim$(Item(.Item).TranslatedName) & ":" & .value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestTakeItems()
    Dim i As Long
    
    frmEditor_Quest.lstTakeItem.Clear
    
    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).TakeItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstTakeItem.AddItem "-"
            Else
                frmEditor_Quest.lstTakeItem.AddItem Trim$(Trim$(Item(.Item).TranslatedName) & ":" & .value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestRewardItems()
    Dim i As Long
    
    frmEditor_Quest.lstItemRew.Clear
    
    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).RewardItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstItemRew.AddItem "-"
            Else
                frmEditor_Quest.lstItemRew.AddItem Trim$(Trim$(Item(.Item).TranslatedName) & ":" & .value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestRequirementItems()
    Dim i As Long
    
    frmEditor_Quest.lstReqItem.Clear
    
    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).RequiredItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstReqItem.AddItem "-"
            Else
                frmEditor_Quest.lstReqItem.AddItem Trim$(Trim$(Item(.Item).TranslatedName) & ":" & .value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestClass()
    Dim i As Long
    
    frmEditor_Quest.lstReqClass.Clear
    
    For i = 1 To 5
        If Quest(EditorIndex).RequiredClass(i) = 0 Then
            frmEditor_Quest.lstReqClass.AddItem "-"
        Else
            frmEditor_Quest.lstReqClass.AddItem Trim$(Trim$(Class(Quest(EditorIndex).RequiredClass(i)).TranslatedName))
        End If
    Next
End Sub
'/Alatar v1.2

Public Sub QuestEditorOk()
Dim i As Long

    For i = 1 To MAX_QUESTS
        If Quest_Changed(i) Then
            Call SendSaveQuest(i)
        End If
    Next
    
    Unload frmEditor_Quest
    Editor = 0
    ClearChanged_Quest
    
End Sub

Public Sub QuestEditorCancel()
    Editor = 0
    Unload frmEditor_Quest
    ClearChanged_Quest
    ClearQuests
    SendRequestQuests
End Sub

Public Sub ClearChanged_Quest()
    ZeroMemory Quest_Changed(1), MAX_QUESTS * 2 ' 2 = boolean length
End Sub

' //////////////
' // DATABASE //
' //////////////

Sub ClearQuest(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(index)), LenB(Quest(index)))
    Quest(index).Name = vbNullString
End Sub

Sub ClearQuests()
Dim i As Long

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next
End Sub

' ////////////////////
' // C&S PROCEDURES //
' ////////////////////

Public Sub SendRequestEditQuest()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditQuest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Public Sub SendSaveQuest(ByVal QuestNum As Long)
Dim Buffer As clsBuffer
'Dim QuestSize As Long
Dim QuestData() As Byte

    Set Buffer = New clsBuffer
    'QuestSize = LenB(Quest(QuestNum))
    'ReDim QuestData(QuestSize - 1)
    'CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    QuestData = GetQuestData(QuestNum)
    Buffer.WriteLong CSaveQuest
    Buffer.WriteLong QuestNum
    'Buffer.WriteLong QuestSize
    Buffer.WriteBytes QuestData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Sub SendRequestQuests()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestQuests
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Public Sub UpdateQuestLog()
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CQuestLogUpdate
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Public Sub PlayerHandleQuest(ByVal QuestNum As Long, ByVal Order As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CPlayerHandleQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteLong Order '1=accept quest, 2=cancel quest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ///////////////
' // Functions //
' ///////////////

'Tells if the quest is in progress or not
Public Function QuestInProgress(ByVal QuestNum As Long) As Boolean
    QuestInProgress = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    
    If Player(MyIndex).PlayerQuest(QuestNum).Status = QUEST_STARTED Then 'Status=1 means started
        QuestInProgress = True
    End If
End Function

Public Function QuestCompleted(ByVal QuestNum As Long) As Boolean
    QuestCompleted = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    
    If Player(MyIndex).PlayerQuest(QuestNum).Status = QUEST_COMPLETED Or Player(MyIndex).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        QuestCompleted = True
    End If
End Function

Public Function GetQuestNum(ByVal QuestName As String) As Long
    Dim i As Long
    GetQuestNum = 0
    
    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).Name) = Trim$(QuestName) Then
            GetQuestNum = i
            Exit For
        End If
        
        If Trim$(Quest(i).TranslatedName) = Trim$(QuestName) Then
            GetQuestNum = i
            Exit For
        End If
        
    Next
End Function

' /////////////////////
' // General Purpose //
' /////////////////////

'Subroutine that load the desired task in the form
Public Sub LoadTask(ByVal QuestNum As Long, ByVal TaskNum As Long)
    Dim TaskToLoad As TaskRec
    TaskToLoad = Quest(QuestNum).Task(TaskNum)
    
    With frmEditor_Quest
        'Load the task type
        .optTask(TaskToLoad.Order).value = True
        'Load textboxes
        .txtTaskSpeech.text = vbNullString
        .txtTaskLog.text = "" & Trim$(TaskToLoad.TaskLog)
        'Set scrolls to 0 and disable them so they can be enabled when needed
        .scrlNPC.value = 0
        .scrlItem.value = 0
        .scrlMap.value = 0
        .scrlResource.value = 0
        .scrlAmount.value = 0
        .txtTaskSpeech.enabled = False
        .scrlNPC.enabled = False
        .scrlItem.enabled = False
        .scrlMap.enabled = False
        .scrlResource.enabled = False
        .scrlAmount.enabled = False
        
        If TaskToLoad.QuestEnd = True Then
            .chkEnd.value = 1
        Else
            .chkEnd.value = 0
        End If
        
        Select Case TaskToLoad.Order
            Case 0 'Nothing
                
            Case QUEST_TYPE_GOSLAY '1
                .scrlNPC.enabled = True
                .scrlNPC.value = TaskToLoad.NPC
                .scrlAmount.enabled = True
                .scrlAmount.value = TaskToLoad.amount
                
            Case QUEST_TYPE_GOGATHER '2
                .scrlItem.enabled = True
                .scrlItem.value = TaskToLoad.Item
                .scrlAmount.enabled = True
                .scrlAmount.value = TaskToLoad.amount
                
            Case QUEST_TYPE_GOTALK '3
                .scrlNPC.enabled = True
                .scrlNPC.value = TaskToLoad.NPC
                .txtTaskSpeech.enabled = True
                .txtTaskSpeech.text = "" & Trim$(TaskToLoad.Speech)
                
            Case QUEST_TYPE_GOREACH '4
                .scrlMap.enabled = True
                .scrlMap.value = TaskToLoad.map
            
            Case QUEST_TYPE_GOGIVE '5
                .scrlItem.enabled = True
                .scrlItem.value = TaskToLoad.Item
                .scrlAmount.enabled = True
                .scrlAmount.value = TaskToLoad.amount
                .scrlNPC.enabled = True
                .scrlNPC.value = TaskToLoad.NPC
                .txtTaskSpeech.enabled = True
                .txtTaskSpeech.text = "" & Trim$(TaskToLoad.Speech)
            
            Case QUEST_TYPE_GOKILL '6
                .scrlAmount.enabled = True
                .scrlAmount.value = TaskToLoad.amount
                
            Case QUEST_TYPE_GOTRAIN '7
                .scrlResource.enabled = True
                .scrlResource.value = TaskToLoad.Resource
                .scrlAmount.enabled = True
                .scrlAmount.value = TaskToLoad.amount
            
            Case QUEST_TYPE_GOGET '8
                .scrlNPC.enabled = True
                .scrlNPC.value = TaskToLoad.NPC
                .scrlItem.enabled = True
                .scrlItem.value = TaskToLoad.Item
                .scrlAmount.enabled = True
                .scrlAmount.value = TaskToLoad.amount
                .txtTaskSpeech.enabled = True
                .txtTaskSpeech.text = "" & Trim$(TaskToLoad.Speech)
            
        End Select
    
    End With
    
    
End Sub

Public Sub RefreshQuestLog()
    Dim i As Long
    
    frmMain.lstQuestLog.Clear
    For i = 1 To MAX_QUESTS
        If QuestInProgress(i) Or QuestCompleted(i) Then
            frmMain.lstQuestLog.AddItem Trim$(Quest(i).TranslatedName)
        End If
    Next
    
End Sub

' ////////////////////////
' // Visual Interaction //
' ////////////////////////

Public Sub LoadQuestlogBox(ByVal ButtonPressed As Integer)
    Dim QuestNum As Long, i As Long
    Dim QuestSay As String
    
    With frmMain
        If Trim$(.lstQuestLog.text) = vbNullString Then Exit Sub
        
        QuestNum = GetQuestNum(Trim$(.lstQuestLog.text))
        
        Select Case ButtonPressed
            Case 1 'Actual Task
                .lblQuestSubtitle = GetTranslation("Tarea Actual ") & " [" + Trim$(Player(MyIndex).PlayerQuest(QuestNum).ActualTask) + "]"
                If QuestCompleted(QuestNum) = False Then
                    .lblQuestSay = Trim$(Quest(QuestNum).Task(Player(MyIndex).PlayerQuest(QuestNum).ActualTask).TaskLog)
                Else
                    .lblQuestSay = "."
                End If
                
            Case 2 'Last Speech
                .lblQuestSubtitle = GetTranslation("Último discurso")
                If Player(MyIndex).PlayerQuest(QuestNum).ActualTask > 1 Then
                    .lblQuestSay = Trim$(Quest(QuestNum).Task(Player(MyIndex).PlayerQuest(QuestNum).ActualTask - 1).Speech)
                    If .lblQuestSay = "" Then
                        .lblQuestSay = Trim$(Quest(QuestNum).Speech(1))
                    End If
                Else
                    .lblQuestSay = Trim$(Quest(QuestNum).Speech(1))
                End If
            
            Case 3 'Quest Status
                .lblQuestSubtitle = GetTranslation("Estado de la Quest")
                If Player(MyIndex).PlayerQuest(QuestNum).Status = QUEST_STARTED Then
                    .lblQuestSay = GetTranslation("Misión en progreso. Paso " & Player(MyIndex).PlayerQuest(QuestNum).ActualTask & ".")
                    .lblQuestExtra = GetTranslation("Cancelar Quest")
                    .lblQuestExtra.Visible = True
                ElseIf QuestCompleted(QuestNum) Then
                    .lblQuestSay = GetTranslation("Completada")
                End If
                
            Case 4 'Quest Log (Main Task)
                .lblQuestSubtitle = GetTranslation("Tarea principal")
                .lblQuestSay = Trim$(Quest(QuestNum).QuestLog)
            
            Case 5 'Requirements
                .lblQuestSubtitle = GetTranslation("Requerimientos")
                QuestSay = "Nivel: "
                If Quest(QuestNum).RequiredLevel > 0 Then
                    QuestSay = QuestSay & "" & Quest(QuestNum).RequiredLevel & vbNewLine & "Quest: "
                Else
                    QuestSay = QuestSay & " Nada." & vbNewLine & "Quest: "
                End If
                If Quest(QuestNum).RequiredQuest > 0 Then
                    QuestSay = QuestSay & "" & Trim$(Quest(Quest(QuestNum).RequiredQuest).TranslatedName) & vbNewLine & "Clase: "
                Else
                    QuestSay = QuestSay & " Nada." & vbNewLine & "Clase: "
                End If
                For i = 1 To 5
                    If Quest(QuestNum).RequiredClass(i) > 0 Then
                        QuestSay = QuestSay & Trim$(Class(Quest(QuestNum).RequiredClass(i)).Name) & ". "
                    End If
                Next
                QuestSay = QuestSay & vbNewLine & "Items:"
                For i = 1 To MAX_QUESTS_ITEMS
                    If Quest(QuestNum).RequiredItem(i).Item > 0 Then
                        QuestSay = QuestSay & " " & Trim$(Item(Quest(QuestNum).RequiredItem(i).Item).TranslatedName) & "(" & Trim$(Quest(QuestNum).RequiredItem(i).value) & ")"
                    End If
                Next
                .lblQuestSay = QuestSay
            
            Case 6 'Rewards
                .lblQuestSubtitle = GetTranslation("Recompensa")
                QuestSay = "Experiencia: " & Quest(QuestNum).RewardExp & vbNewLine & "Items:"
                For i = 1 To MAX_QUESTS_ITEMS
                    If Quest(QuestNum).RewardItem(i).Item > 0 Then
                        QuestSay = QuestSay & " " & Trim$(Item(Quest(QuestNum).RewardItem(i).Item).TranslatedName) & "(" & Trim$(Quest(QuestNum).RewardItem(i).value) & ")"
                    End If
                Next
                .lblQuestSay = QuestSay
            
            Case Else
                Exit Sub
        End Select
        
        .lblQuestName = Trim$(Quest(QuestNum).TranslatedName)
        .picQuestDialogue.Visible = True
        
    End With
End Sub

Public Sub RunQuestDialogueExtraLabel()
    If frmMain.lblQuestExtra = GetTranslation("Cancelar Quest") Then
        PlayerHandleQuest GetQuestNum(Trim$(frmMain.lblQuestName.Caption)), 2
        frmMain.lblQuestExtra = "Extra"
        frmMain.lblQuestExtra.Visible = False
        frmMain.picQuestDialogue.Visible = False
        RefreshQuestLog
    End If
End Sub


Function GetQuestData(ByVal QuestNum As Long) As Byte()
Dim Buffer As clsBuffer, i As Long
Set Buffer = New clsBuffer

With Quest(QuestNum)
     Buffer.WriteString .Name
     Buffer.WriteLong .Repeat
     Buffer.WriteString .QuestLog
    
    For i = 1 To 3
         Buffer.WriteString .Speech(i)
    Next
    
    For i = 1 To MAX_QUESTS_ITEMS
         Buffer.WriteLong .GiveItem(i).Item
         Buffer.WriteLong .GiveItem(i).value
        
         Buffer.WriteLong .TakeItem(i).Item
         Buffer.WriteLong .TakeItem(i).value
        
         Buffer.WriteLong .RequiredItem(i).Item
         Buffer.WriteLong .RequiredItem(i).value
        
         Buffer.WriteLong .RewardItem(i).Item
         Buffer.WriteLong .RewardItem(i).value
    Next
    
     Buffer.WriteLong .RequiredLevel
     Buffer.WriteLong .RequiredQuest
    For i = 1 To 5
         Buffer.WriteLong .RequiredClass(i)
    Next
    
     Buffer.WriteLong .RewardExp
    
    For i = 1 To MAX_TASKS
         Buffer.WriteLong .Task(i).Order
         Buffer.WriteLong .Task(i).NPC
         Buffer.WriteLong .Task(i).Item
        Buffer.WriteLong .Task(i).map
         Buffer.WriteLong .Task(i).Resource
        Buffer.WriteLong .Task(i).amount
         Buffer.WriteString .Task(i).Speech
         Buffer.WriteString .Task(i).TaskLog
         Buffer.WriteByte .Task(i).QuestEnd
    Next
    
    Buffer.WriteLong .Level
End With
GetQuestData = Buffer.ToArray()
Set Buffer = Nothing
End Function

Sub SetQuestData(ByRef Data() As Byte, ByVal QuestNum As Long)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteBytes Data
    
    
    With Quest(QuestNum)
    .Name = Buffer.ReadString
    .TranslatedName = Buffer.ReadString
    .Repeat = Buffer.ReadLong
    .QuestLog = Buffer.ReadString
    
    Dim i As Long
    
    For i = 1 To 3
       .Speech(i) = Buffer.ReadString
    Next
    
    For i = 1 To MAX_QUESTS_ITEMS
        .GiveItem(i).Item = Buffer.ReadLong
        .GiveItem(i).value = Buffer.ReadLong
        
        .TakeItem(i).Item = Buffer.ReadLong
        .TakeItem(i).value = Buffer.ReadLong
        
        .RequiredItem(i).Item = Buffer.ReadLong
        .RequiredItem(i).value = Buffer.ReadLong
        
        .RewardItem(i).Item = Buffer.ReadLong
        .RewardItem(i).value = Buffer.ReadLong
    Next
    
    .RequiredLevel = Buffer.ReadLong
    .RequiredQuest = Buffer.ReadLong
    For i = 1 To 5
        .RequiredClass(i) = Buffer.ReadLong
    Next
    
    .RewardExp = Buffer.ReadLong
    
    For i = 1 To MAX_TASKS
        .Task(i).Order = Buffer.ReadLong
        .Task(i).NPC = Buffer.ReadLong
        .Task(i).Item = Buffer.ReadLong
        .Task(i).map = Buffer.ReadLong
        .Task(i).Resource = Buffer.ReadLong
        .Task(i).amount = Buffer.ReadLong
        .Task(i).Speech = Buffer.ReadString
        .Task(i).TaskLog = Buffer.ReadString
        .Task(i).QuestEnd = Buffer.ReadByte
    Next
    
    .Level = Buffer.ReadLong
    End With
End Sub
