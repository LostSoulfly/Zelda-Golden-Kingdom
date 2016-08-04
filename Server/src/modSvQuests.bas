Attribute VB_Name = "modSvQuests"
Option Explicit
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const MAX_QUESTS As Byte = 70
Public Const MAX_QUESTS_ITEMS As Byte = 10 'Alatar v1.2

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

'Types
Public Quest(1 To MAX_QUESTS) As QuestRec
'Public Quest2(1 To MAX_QUESTS) As QuestRec2

Public Type PlayerQuestRec
    Status As Long
    ActualTask As Long
    CurrentCount As Long 'Used to handle the Amount property
End Type

'Alatar v1.2
Private Type QuestRequiredItemRec
    item As Long
    Value As Long
End Type

Private Type QuestGiveItemRec
    item As Long
    Value As Long
End Type

Private Type QuestTakeItemRec
    item As Long
    Value As Long
End Type

Private Type QuestRewardItemRec
    item As Long
    Value As Long
End Type
'/Alatar v1.2
Public Type TaskRec
    Order As Long
    NPC As Long
    item As Long
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
    level As Long
    '/Alatar v1.2
    TranslatedName As String * 30
End Type

'Public Type QuestRec2
'    'Alatar v1.2
'    Name As String * 30
'
'    Repeat As Long
'    QuestLog As String
'    Speech(1 To 3) As String
'    GiveItem(1 To MAX_QUESTS_ITEMS) As QuestGiveItemRec
'    TakeItem(1 To MAX_QUESTS_ITEMS) As QuestTakeItemRec
'
'    RequiredLevel As Long
'    RequiredQuest As Long
'    RequiredClass(1 To 5) As Long
'    RequiredItem(1 To MAX_QUESTS_ITEMS) As QuestRequiredItemRec
'
'    RewardExp As Long
'    RewardItem(1 To MAX_QUESTS_ITEMS) As QuestRewardItemRec
'
'    Task(1 To MAX_TASKS) As TaskRec
'    level As Long
'    '/Alatar v1.2
'    TranslatedName As String * 30
'End Type


' //////////////
' // DATABASE //
' //////////////

Sub SaveQuests()
    Dim i As Long
    For i = 1 To MAX_QUESTS
        Call SaveQuest(i)
    Next
End Sub

Sub SaveQuest(ByVal questnum As Long)
    Dim FileName As String
    Dim F As Long, i As Long
    FileName = App.Path & "\data\quests\quest" & questnum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    
        Put #F, , Quest(questnum)
        Close #F
        Exit Sub
        'Alatar v1.2
        Put #F, , Quest(questnum).Name
        Put #F, , Quest(questnum).Repeat
        Put #F, , Quest(questnum).QuestLog
        For i = 1 To 3
            Put #F, , Quest(questnum).Speech(i)
        Next
        For i = 1 To MAX_QUESTS_ITEMS
            Put #F, , Quest(questnum).GiveItem(i)
        Next
        For i = 1 To MAX_QUESTS_ITEMS
            Put #F, , Quest(questnum).TakeItem(i)
        Next
        Put #F, , Quest(questnum).RequiredLevel
        Put #F, , Quest(questnum).RequiredQuest
        For i = 1 To 5
            Put #F, , Quest(questnum).RequiredClass(i)
        Next
        For i = 1 To MAX_QUESTS_ITEMS
            Put #F, , Quest(questnum).RequiredItem(i)
        Next
        Put #F, , Quest(questnum).RewardExp
        For i = 1 To MAX_QUESTS_ITEMS
            Put #F, , Quest(questnum).RewardItem(i)
        Next
        For i = 1 To MAX_TASKS
            Put #F, , Quest(questnum).Task(i)
        Next
        Put #F, , Quest(questnum).level
        '/Alatar v1.2
    Close #F
End Sub

Sub LoadQuests()
    Dim FileName As String
    Dim i As Integer
    Dim F As Long, N As Long
    Dim sLen As Long
    
    Call CheckQuests

    For i = 1 To MAX_QUESTS
        FileName = App.Path & "\data\quests\quest2.0-" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Quest(i)
        Close #F
        'Quest(i).Name = GetTranslation(Quest(i).Name)
        'Quest(i).Speech(1) = GetTranslation(Quest(i).Speech(1))
        'Quest(i).Speech(2) = GetTranslation(Quest(i).Speech(2))
        'Quest(i).Speech(3) = GetTranslation(Quest(i).Speech(3))
        'If Trim$(Replace(Quest(i).TranslatedName, vbNullChar, "")) = "" Then Quest(i).TranslatedName = GetTranslation(Quest(i).Name)
        Next
        Exit Sub
        'Alatar v1.2
        Get #F, , Quest(i).Name
        Get #F, , Quest(i).Repeat
        Get #F, , Quest(i).QuestLog
        For N = 1 To 3
            Get #F, , Quest(i).Speech(N)
        Next
        For N = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(i).GiveItem(N)
        Next
        For N = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(i).TakeItem(N)
        Next
        Get #F, , Quest(i).RequiredLevel
        Get #F, , Quest(i).RequiredQuest
        For N = 1 To 5
            Get #F, , Quest(i).RequiredClass(N)
        Next
        For N = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(i).RequiredItem(N)
        Next
        Get #F, , Quest(i).RewardExp
        For N = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(i).RewardItem(N)
        Next
        For N = 1 To MAX_TASKS
            Get #F, , Quest(i).Task(N)
        'Next
        
        Get #F, , Quest(i).level
        '/Alatar v1.2
        Close #F

    Next
        
End Sub

'Sub migrateQuests()
'Dim FileName As String
'Dim i As Long
'Dim F As Long

'For i = 1 To MAX_QUESTS
'    UpdateQuest Quest(i), Quest2(i)
'    FileName = App.Path & "\data\quests\quest2.0-" & i & ".dat"
'    F = FreeFile
'    Open FileName For Binary As #F
'    Put #F, , Quest2(i)
'    Close #F
    'Exit Sub
'Next

'End Sub

'Sub UpdateQuest(oldQuest As QuestRec, newQuest As QuestRec2)
'Dim n As Integer

'With newQuest
'.Name = oldQuest.Name
'.TranslatedName = GetTranslation(.Name)
'.QuestLog = oldQuest.QuestLog

'For n = 1 To 3
'    .Speech(n) = oldQuest.Speech(n)
'Next
'For n = 1 To 3
'    .Speech(n) = oldQuest.Speech(n)
'Next
'For n = 1 To MAX_QUESTS_ITEMS
'    .GiveItem(n) = oldQuest.GiveItem(n)
'Next
'For n = 1 To MAX_QUESTS_ITEMS
'    .TakeItem(n) = oldQuest.TakeItem(n)
'Next
'.RequiredLevel = oldQuest.RequiredLevel
'.RequiredQuest = oldQuest.RequiredQuest
'For n = 1 To 5
'   .RequiredClass(n) = oldQuest.RequiredClass(n)
'Next
'For n = 1 To MAX_QUESTS_ITEMS
'    .RequiredItem(n) = oldQuest.RequiredItem(n)
'Next
'.RewardExp = oldQuest.RewardExp
'For n = 1 To MAX_QUESTS_ITEMS
'    .RewardItem(n) = oldQuest.RewardItem(n)
'Next
'For n = 1 To MAX_TASKS
'    .Task(n) = oldQuest.Task(n)
'Next
'.level = oldQuest.level
'End With

'End Sub

Sub CheckQuests()
    Dim i As Long
    For i = 1 To MAX_QUESTS
        If Not FileExist("\Data\quests\quest" & i & ".dat") Then
            Call SaveQuest(i)
        End If
    Next
End Sub

Sub ClearQuest(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(index)), LenB(Quest(index)))
    Quest(index).Name = vbNullString
    'Quest(index).TranslatedName = vbNullString
    Quest(index).QuestLog = vbNullString
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

Sub SendQuests(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If QuestExists(i) Then
            Call SendUpdateQuestTo(index, i)
        End If
    Next
End Sub

Function QuestData(ByVal questnum As Long) As Byte()
Dim buffer As clsBuffer, i As Long
Set buffer = New clsBuffer

With Quest(questnum)
     buffer.WriteString .Name
     buffer.WriteString .TranslatedName
     buffer.WriteLong .Repeat
     buffer.WriteString GetTranslation(.QuestLog)
    
    For i = 1 To 3
         buffer.WriteString GetTranslation(.Speech(i))
    Next
    
    For i = 1 To MAX_QUESTS_ITEMS
         buffer.WriteLong .GiveItem(i).item
         buffer.WriteLong .GiveItem(i).Value
        
         buffer.WriteLong .TakeItem(i).item
         buffer.WriteLong .TakeItem(i).Value
        
         buffer.WriteLong .RequiredItem(i).item
         buffer.WriteLong .RequiredItem(i).Value
        
         buffer.WriteLong .RewardItem(i).item
         buffer.WriteLong .RewardItem(i).Value
    Next
    
     buffer.WriteLong .RequiredLevel
     buffer.WriteLong .RequiredQuest
    For i = 1 To 5
         buffer.WriteLong .RequiredClass(i)
    Next
    
     buffer.WriteLong .RewardExp
    
    For i = 1 To MAX_TASKS
         buffer.WriteLong .Task(i).Order
         buffer.WriteLong .Task(i).NPC
         buffer.WriteLong .Task(i).item
        buffer.WriteLong .Task(i).map
         buffer.WriteLong .Task(i).Resource
        buffer.WriteLong .Task(i).amount
         buffer.WriteString GetTranslation(.Task(i).Speech)
         buffer.WriteString GetTranslation(.Task(i).TaskLog)
         buffer.WriteByte .Task(i).QuestEnd
    Next
    
    buffer.WriteLong .level
End With
QuestData = buffer.ToArray()
Set buffer = Nothing
End Function

Sub SendUpdateQuestToAll(ByVal questnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    'Dim QuestSize As Long
   ' Dim QuestData() As Byte
    Set buffer = New clsBuffer
    'QuestSize = LenB(Quest(QuestNum))
    'ReDim QuestData(QuestSize - 1)
    'CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    buffer.WriteLong SUpdateQuest
    buffer.WriteLong questnum
    buffer.WriteBytes CompressData(QuestData(questnum), 2)
    
    'buffer.WriteBytes QuestData(QuestNum)
    SendDataToAll buffer.ToArray()

    Set buffer = Nothing
End Sub

Sub SendUpdateQuestTo(ByVal index As Long, ByVal questnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    'Dim QuestSize As Long
    'Dim QuestData() As Byte
    Set buffer = New clsBuffer
    'QuestSize = LenB(Quest(QuestNum))
    'ReDim QuestData(QuestSize - 1)
    'CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    buffer.WriteLong SUpdateQuest
    buffer.WriteLong questnum
    'buffer.WriteBytes QuestData
    buffer.WriteBytes CompressData(QuestData(questnum), 2)
    
    
    'buffer.WriteBytes QuestData(QuestNum)
    SendDataTo index, buffer.ToArray()
    
    ByteCounter = ByteCounter + buffer.length
    Set buffer = Nothing
End Sub

Public Sub SendPlayerQuests(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerQuest
        For i = 1 To MAX_QUESTS
            buffer.WriteLong player(index).PlayerQuest(i).Status
            buffer.WriteLong player(index).PlayerQuest(i).ActualTask
            buffer.WriteLong player(index).PlayerQuest(i).CurrentCount
        Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing

End Sub

Public Sub SendPlayerQuest(ByVal index As Long, ByVal questnum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerQuest
    buffer.WriteLong player(index).PlayerQuest(questnum).Status
    buffer.WriteLong player(index).PlayerQuest(questnum).ActualTask
    buffer.WriteLong player(index).PlayerQuest(questnum).CurrentCount
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

'sends a message to the client that is shown on the screen
Public Sub QuestMessage(ByVal index As Long, ByVal questnum As Long, ByVal message As String, ByVal QuestNumForStart As Long, Optional blForceTranslate As Boolean = False)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    If blForceTranslate = True Then message = GetTranslation(message)
    
    buffer.WriteLong SQuestMessage
    buffer.WriteLong questnum
    buffer.WriteString Trim$(message)
    buffer.WriteLong QuestNumForStart
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
End Sub

' ///////////////
' // Functions //
' ///////////////

Public Function CanStartQuest(ByVal index As Long, ByVal questnum As Long) As Boolean
    Dim i As Long, N As Long
    CanStartQuest = False
    If questnum < 1 Or questnum > MAX_QUESTS Then Exit Function
    If QuestInProgress(index, questnum) Then Exit Function
    
    'check if now a completed quest can be repeated
    If player(index).PlayerQuest(questnum).Status = QUEST_COMPLETED Then
        If Quest(questnum).Repeat = YES Then
            player(index).PlayerQuest(questnum).Status = QUEST_COMPLETED_BUT
            Exit Function
        End If
    End If
    
    'Check if player has the quest 0 (not started) or 3 (completed but it can be started again)
    If player(index).PlayerQuest(questnum).Status = QUEST_NOT_STARTED Or player(index).PlayerQuest(questnum).Status = QUEST_COMPLETED_BUT Then
        'Check if player's level is right
        If Quest(questnum).RequiredLevel <= player(index).level Then
            
            'Check if item is needed
            For i = 1 To MAX_QUESTS_ITEMS
                If Quest(questnum).RequiredItem(i).item > 0 Then
                    'if we don't have it at all then
                    If HasItem(index, Quest(questnum).RequiredItem(i).item) = 0 Then
                        PlayerMsg index, GetTranslation("¡Necesitas", , UnTrimBack) & Trim$(item(Quest(questnum).RequiredItem(i).item).TranslatedName) & GetTranslation("para aceptar ésta quest!", , UnTrimFront), BrightRed, , False
                    ' send the sound
                    SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
                        Exit Function
                    End If
                End If
            Next
            
            'Check if previous quest is needed
            If Quest(questnum).RequiredQuest > 0 And Quest(questnum).RequiredQuest <= MAX_QUESTS Then
                If player(index).PlayerQuest(Quest(questnum).RequiredQuest).Status = QUEST_NOT_STARTED Or player(index).PlayerQuest(Quest(questnum).RequiredQuest).Status = QUEST_STARTED Then
                    PlayerMsg index, GetTranslation("¡Necesitas completar la quest", , UnTrimBack) & Trim$(Quest(Quest(questnum).RequiredQuest).TranslatedName) & GetTranslation("para aceptar ésta quest!", , UnTrimFront), BrightRed, , False
                    ' send the sound
                    SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
                    Exit Function
                End If
            End If
            'Go on :)
            CanStartQuest = True
        Else
            PlayerMsg index, "Necesitas tener más nivel para aceptar ésta quest!", BrightRed
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
        End If
    Else
        PlayerMsg index, "¡No puedes aceptar ésta quest otra vez!", BrightRed
    End If
End Function

Public Function CanEndQuest(ByVal index As Long, questnum As Long) As Boolean
    CanEndQuest = False
    If Quest(questnum).Task(player(index).PlayerQuest(questnum).ActualTask).QuestEnd = True Then
        CanEndQuest = True
    End If
End Function

'Tells if the quest is in progress or not
Public Function QuestInProgress(ByVal index As Long, ByVal questnum As Long) As Boolean
    QuestInProgress = False
    If questnum < 1 Or questnum > MAX_QUESTS Then Exit Function
    
    If player(index).PlayerQuest(questnum).Status = QUEST_STARTED Then
        QuestInProgress = True
    End If
End Function

Public Function QuestCompleted(ByVal index As Long, ByVal questnum As Long) As Boolean
    QuestCompleted = False
    If questnum < 1 Or questnum > MAX_QUESTS Then Exit Function
    
    If player(index).PlayerQuest(questnum).Status = QUEST_COMPLETED Or player(index).PlayerQuest(questnum).Status = QUEST_COMPLETED_BUT Then
        QuestCompleted = True
    End If
End Function

'Gets the quest reference num (id) from the quest name (it shall be unique)
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

Public Function GetItemNum(ByVal ItemName As String) As Long
    Dim i As Long
    GetItemNum = 0
    
    For i = 1 To MAX_ITEMS
        If Trim$(item(i).Name) = Trim$(ItemName) Then
            GetItemNum = i
            Exit For
        End If
        
        If Trim$(item(i).TranslatedName) = Trim$(ItemName) Then
            GetItemNum = i
            Exit For
        End If
    
    Next
End Function

' /////////////////////
' // General Purpose //
' /////////////////////

Public Sub CheckTasks(ByVal index As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim i As Long
    
    For i = 1 To MAX_QUESTS
        If QuestInProgress(index, i) Then
            If TaskType = Quest(i).Task(player(index).PlayerQuest(i).ActualTask).Order Then
                Call CheckTask(index, i, TaskType, TargetIndex)
            End If
        End If
    Next
End Sub

Public Sub CheckTask(ByVal index As Long, ByVal questnum As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim ActualTask As Long, i As Long
    ActualTask = player(index).PlayerQuest(questnum).ActualTask
    
    Select Case TaskType
        Case QUEST_TYPE_GOSLAY 'Kill X amount of X npc's.
        
            'is npc's defeated id is the same as the npc i have to kill?
            If TargetIndex = Quest(questnum).Task(ActualTask).NPC Then
                'Count +1
                player(index).PlayerQuest(questnum).CurrentCount = player(index).PlayerQuest(questnum).CurrentCount + 1
                'show msg
                PlayerMsg index, GetTranslation("Quest:", , UnTrimBack) & Trim$(Quest(questnum).TranslatedName) & " - " & Trim$(player(index).PlayerQuest(questnum).CurrentCount) & "/" & Trim$(Quest(questnum).Task(ActualTask).amount) & " " & Trim$(NPC(TargetIndex).TranslatedName) & GetTranslation("matados.", , UnTrimFront), Yellow, , False
                'did i finish the work?
                If player(index).PlayerQuest(questnum).CurrentCount >= Quest(questnum).Task(ActualTask).amount Then
                    QuestMessage index, questnum, "Tarea completada", 0, True
                    'is the quest's end?
                    If CanEndQuest(index, questnum) Then
                        EndQuest index, questnum
                    Else
                        'otherwise continue to the next task
                        player(index).PlayerQuest(questnum).CurrentCount = 0
                        player(index).PlayerQuest(questnum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
                        
        Case QUEST_TYPE_GOGATHER 'Gather X amount of X item.
            If TargetIndex = Quest(questnum).Task(ActualTask).item Then
                
                'reset the count first if we started
                If player(index).PlayerQuest(questnum).CurrentCount < 1 Then
                    player(index).PlayerQuest(questnum).CurrentCount = 0
                End If
                
                'Check inventory for the items
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(index, i) = TargetIndex Then
                        If isItemStackable(i) Then
                            player(index).PlayerQuest(questnum).CurrentCount = GetPlayerInvItemValue(index, i)
                        Else
                            'If is the correct item add it to the count
                            player(index).PlayerQuest(questnum).CurrentCount = player(index).PlayerQuest(questnum).CurrentCount + 1
                        End If
                    End If
                Next
                
                PlayerMsg index, GetTranslation("Quest:", , UnTrimBack) & Trim$(Quest(questnum).TranslatedName) & " - " & GetTranslation("Tienes", , UnTrimBack) & Trim$(player(index).PlayerQuest(questnum).CurrentCount) & "/" & Trim$(Quest(questnum).Task(ActualTask).amount) & " " & Trim$(item(TargetIndex).TranslatedName), Yellow, , False
                
                If player(index).PlayerQuest(questnum).CurrentCount >= Quest(questnum).Task(ActualTask).amount Then
                    QuestMessage index, questnum, "Tarea completada", 0, True
                    
                    If CanEndQuest(index, questnum) Then
                        EndQuest index, questnum
                    Else
                        player(index).PlayerQuest(questnum).CurrentCount = 0
                        player(index).PlayerQuest(questnum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
            
        Case QUEST_TYPE_GOTALK 'Interact with X npc.
            If TargetIndex = Quest(questnum).Task(ActualTask).NPC Then
                QuestMessage index, questnum, GetTranslation(Quest(questnum).Task(ActualTask).Speech), 0
                If CanEndQuest(index, questnum) Then
                    EndQuest index, questnum
                Else
                    player(index).PlayerQuest(questnum).ActualTask = ActualTask + 1
                End If
            End If
        
        Case QUEST_TYPE_GOREACH 'Reach X map.
            If TargetIndex = Quest(questnum).Task(ActualTask).map Then
                QuestMessage index, questnum, "Tarea completada", 0, True
                If CanEndQuest(index, questnum) Then
                    EndQuest index, questnum
                Else
                    player(index).PlayerQuest(questnum).ActualTask = ActualTask + 1
                End If
            End If
        
        Case QUEST_TYPE_GOGIVE 'Give X amount of X item to X npc.
            If TargetIndex = Quest(questnum).Task(ActualTask).NPC Then
                
                player(index).PlayerQuest(questnum).CurrentCount = 0
                
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(index, i) = Quest(questnum).Task(ActualTask).item Then
                        If isItemStackable(i) Then
                            If GetPlayerInvItemValue(index, i) >= Quest(questnum).Task(ActualTask).amount Then
                                player(index).PlayerQuest(questnum).CurrentCount = GetPlayerInvItemValue(index, i)
                            End If
                        Else
                            'If is the correct item add it to the count
                            player(index).PlayerQuest(questnum).CurrentCount = player(index).PlayerQuest(questnum).CurrentCount + 1
                        End If
                    End If
                Next
                
                If player(index).PlayerQuest(questnum).CurrentCount >= Quest(questnum).Task(ActualTask).amount Then
                    'if we have enough items, then remove them and finish the task
                    If isItemStackable(Quest(questnum).Task(ActualTask).item) Then
                        TakeInvItem index, Quest(questnum).Task(ActualTask).item, Quest(questnum).Task(ActualTask).amount
                    Else
                        'If it's not a currency then remove all the items
                        For i = 1 To Quest(questnum).Task(ActualTask).amount
                            TakeInvItem index, Quest(questnum).Task(ActualTask).item, 1
                        Next
                    End If
                    
                    PlayerMsg index, GetTranslation("Quest:", , UnTrimBack) & Trim$(Quest(questnum).TranslatedName) & " - " & GetTranslation("Has dado", , UnTrimBack) & Trim$(Quest(questnum).Task(ActualTask).amount) & " " & Trim$(item(TargetIndex).TranslatedName), Yellow, , False
                    QuestMessage index, questnum, GetTranslation(Quest(questnum).Task(ActualTask).Speech), 0
                    
                    If CanEndQuest(index, questnum) Then
                        EndQuest index, questnum
                    Else
                        player(index).PlayerQuest(questnum).CurrentCount = 0
                        player(index).PlayerQuest(questnum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
                    
        Case QUEST_TYPE_GOKILL 'Kill X amount of players.
            player(index).PlayerQuest(questnum).CurrentCount = player(index).PlayerQuest(questnum).CurrentCount + 1
            PlayerMsg index, GetTranslation("Quest:", , UnTrimBack) & Trim$(Quest(questnum).TranslatedName) & " - " & Trim$(player(index).PlayerQuest(questnum).CurrentCount) & "/" & Trim$(Quest(questnum).Task(ActualTask).amount) & GetTranslation("jugadores matados.", , UnTrimFront), Yellow, , False
            If player(index).PlayerQuest(questnum).CurrentCount >= Quest(questnum).Task(ActualTask).amount Then
                QuestMessage index, questnum, "Tarea completada", 0, True
                If CanEndQuest(index, questnum) Then
                    EndQuest index, questnum
                Else
                    player(index).PlayerQuest(questnum).CurrentCount = 0
                    player(index).PlayerQuest(questnum).ActualTask = ActualTask + 1
                End If
            End If
            
        Case QUEST_TYPE_GOTRAIN 'Hit X amount of times X resource.
            If TargetIndex = Quest(questnum).Task(ActualTask).Resource Then
                player(index).PlayerQuest(questnum).CurrentCount = player(index).PlayerQuest(questnum).CurrentCount + 1
                PlayerMsg index, GetTranslation("Quest:", , UnTrimBack) & Trim$(Quest(questnum).TranslatedName) & " - " & Trim$(player(index).PlayerQuest(questnum).CurrentCount) & "/" & Trim$(Quest(questnum).Task(ActualTask).amount) & GetTranslation("golpes.", , UnTrimFront), Yellow, , False
                If player(index).PlayerQuest(questnum).CurrentCount >= Quest(questnum).Task(ActualTask).amount Then
                    QuestMessage index, questnum, "Tarea completada", 0, True
                    If CanEndQuest(index, questnum) Then
                        EndQuest index, questnum
                    Else
                        player(index).PlayerQuest(questnum).CurrentCount = 0
                        player(index).PlayerQuest(questnum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
                      
        Case QUEST_TYPE_GOGET 'Get X amount of X item from X npc.
            If TargetIndex = Quest(questnum).Task(ActualTask).NPC Then
                GiveInvItem index, Quest(questnum).Task(ActualTask).item, Quest(questnum).Task(ActualTask).amount
                QuestMessage index, questnum, GetTranslation(Quest(questnum).Task(ActualTask).Speech), 0
                If CanEndQuest(index, questnum) Then
                    EndQuest index, questnum
                Else
                    player(index).PlayerQuest(questnum).ActualTask = ActualTask + 1
                End If
            End If
        
    End Select
    SendPlayerQuests index
End Sub

Public Sub EndQuest(ByVal index As Long, ByVal questnum As Long)
    Dim i As Long, N As Long
    
    'remove items on the end
    For i = 1 To MAX_QUESTS_ITEMS
        If Quest(questnum).TakeItem(i).item > 0 Then
        If HasItem(index, Quest(questnum).TakeItem(i).item) <= 0 Then
        Exit Sub
            ElseIf HasItem(index, Quest(questnum).TakeItem(i).item) > 0 Then
                If isItemStackable(Quest(questnum).TakeItem(i).item) Then
                    TakeInvItem index, Quest(questnum).TakeItem(i).item, Quest(questnum).TakeItem(i).Value
                Else
                    For N = 1 To Quest(questnum).TakeItem(i).Value
                        TakeInvItem index, Quest(questnum).TakeItem(i).item, 1
                    Next
                End If
            End If
        End If
    Next
    
    player(index).PlayerQuest(questnum).Status = QUEST_COMPLETED
    
    'reset counters to 0
    player(index).PlayerQuest(questnum).ActualTask = 0
    player(index).PlayerQuest(questnum).CurrentCount = 0
    
    'give experience
    CheckQuestExp index, questnum
    
    'give rewards
    For i = 1 To MAX_QUESTS_ITEMS
        If Quest(questnum).RewardItem(i).item <> 0 Then
            'check if we have space
            If FindOpenInvSlot(index, Quest(questnum).RewardItem(i).item) = 0 Then
                PlayerMsg index, "No tienes espacio en el inventario.", BrightRed
                Exit For
            Else
                'if so, check if it's a currency stack the item in one slot
                If isItemStackable(Quest(questnum).RewardItem(i).item) Then
                    GiveInvItem index, Quest(questnum).RewardItem(i).item, Quest(questnum).RewardItem(i).Value
                Else
                'if not, create a new loop and store the item in a new slot if is possible
                    For N = 1 To Quest(questnum).RewardItem(i).Value
                        If FindOpenInvSlot(index, Quest(questnum).RewardItem(i).item) = 0 Then
                            PlayerMsg index, "No tienes espacio en el inventario.", BrightRed
                            Exit For
                        Else
                            GiveInvItem index, Quest(questnum).RewardItem(i).item, 1
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    'show ending message
    QuestMessage index, questnum, GetTranslation(Quest(questnum).Speech(3)), 0
    
    'mark quest as completed in chat
    PlayerMsg index, Trim$(Quest(questnum).TranslatedName) & ": " & GetTranslation("quest completada"), Green, , False
    
    SendPlayerQuests index
End Sub

Function QuestExists(ByVal questnum As Long) As Boolean
If LenB(Trim$(Quest(questnum).Name)) > 0 And Asc(Quest(questnum).Name) <> 0 Then
    QuestExists = True
End If
End Function

Sub CheckQuestExp(ByVal index As Long, ByVal questnum As Long)
    Dim level As Long, optimallevel As Long, GivenExp As Long
    level = GetPlayerLevel(index)
    optimallevel = Quest(questnum).level
    
    If optimallevel > 0 And optimallevel <= MAX_LEVELS Then
        Dim PercentReduction As Single
        PercentReduction = Line(MAX_LEVELS / 2, 0, 100, 0, 0, Abs(level - optimallevel))
        
        GivenExp = Quest(questnum).RewardExp - Quest(questnum).RewardExp * (PercentReduction / 100)
    Else
        GivenExp = Quest(questnum).RewardExp
    End If
    
    GivePlayerEXP index, GivenExp
End Sub


Sub SetQuestData(ByRef Data() As Byte, ByVal questnum As Long)
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteBytes Data
    
    
    With Quest(questnum)
    .Name = buffer.ReadString
    .TranslatedName = GetTranslation(.Name)
    .Repeat = buffer.ReadLong
    .QuestLog = buffer.ReadString
    Dim i As Long
    
    For i = 1 To 3
       .Speech(i) = buffer.ReadString
    Next
    
    For i = 1 To MAX_QUESTS_ITEMS
        .GiveItem(i).item = buffer.ReadLong
        .GiveItem(i).Value = buffer.ReadLong
        
        .TakeItem(i).item = buffer.ReadLong
        .TakeItem(i).Value = buffer.ReadLong
        
        .RequiredItem(i).item = buffer.ReadLong
        .RequiredItem(i).Value = buffer.ReadLong
        
        .RewardItem(i).item = buffer.ReadLong
        .RewardItem(i).Value = buffer.ReadLong
    Next
    
    .RequiredLevel = buffer.ReadLong
    .RequiredQuest = buffer.ReadLong
    For i = 1 To 5
        .RequiredClass(i) = buffer.ReadLong
    Next
    
    .RewardExp = buffer.ReadLong
    
    For i = 1 To MAX_TASKS
        .Task(i).Order = buffer.ReadLong
        .Task(i).NPC = buffer.ReadLong
        .Task(i).item = buffer.ReadLong
        .Task(i).map = buffer.ReadLong
        .Task(i).Resource = buffer.ReadLong
        .Task(i).amount = buffer.ReadLong
        .Task(i).Speech = buffer.ReadString
        .Task(i).TaskLog = buffer.ReadString
        .Task(i).QuestEnd = buffer.ReadByte
    Next
    
    .level = buffer.ReadLong
    End With
End Sub
