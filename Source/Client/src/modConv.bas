Attribute VB_Name = "modConv"
Option Explicit

Public Sub InitChat(ByVal index As Long, ByVal mapNum As Long, ByVal mapNpcNum As Long, Optional ByVal remoteChat As Boolean = False)
    Dim npcNum As Long
    npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
    ' check if we can chat
    If NPC(npcNum).Conv = 0 Then Exit Sub
    If NPC(npcNum).Conv > MAX_CONVS Then Exit Sub
    If Len(Trim$(Conv(NPC(npcNum).Conv).Name)) = 0 Then Exit Sub
    
    If Not remoteChat Then
        With MapNpc(mapNum).NPC(mapNpcNum)
            .c_inChatWith = index
            .c_lastDir = .Dir
            If GetPlayerY(index) = .y - 1 Then
                .Dir = DIR_UP
            ElseIf GetPlayerY(index) = .y + 1 Then
                .Dir = DIR_DOWN
            ElseIf GetPlayerX(index) = .x - 1 Then
                .Dir = DIR_LEFT
            ElseIf GetPlayerX(index) = .x + 1 Then
                .Dir = DIR_RIGHT
            End If
            ' send NPC's dir to the map
            NpcDir mapNum, mapNpcNum, .Dir
        End With
    End If
    
    ' Set chat value to Npc
    TempPlayer(index).inChatWith = npcNum
    TempPlayer(index).c_mapNpcNum = mapNpcNum
    TempPlayer(index).c_mapNum = mapNum
    ' set to the root chat
    TempPlayer(index).curChat = 1
    ' send the root chat
    sendChat index
End Sub

Public Sub chatOption(ByVal index As Long, ByVal chatOption As Long)
    Dim exitChat As Boolean
    Dim convNum As Long
    Dim curChat As Long
    
    convNum = NPC(TempPlayer(index).inChatWith).Conv
    curChat = TempPlayer(index).curChat
    
    exitChat = False
    
    ' follow route
    If Conv(convNum).Conv(curChat).rTarget(chatOption) = 0 Then
        exitChat = True
    Else
        TempPlayer(index).curChat = Conv(convNum).Conv(curChat).rTarget(chatOption)
    End If
    
    ' if exiting chat, clear temp values
    If exitChat Then
        TempPlayer(index).inChatWith = 0
        TempPlayer(index).curChat = 0
        ' send chat update
        sendChat index
        ' send npc dir
        With MapNpc(TempPlayer(index).c_mapNum).NPC(TempPlayer(index).c_mapNpcNum)
            If .c_inChatWith = index Then
                .c_inChatWith = 0
                .Dir = .c_lastDir
                NpcDir TempPlayer(index).c_mapNum, TempPlayer(index).c_mapNpcNum, .Dir
            End If
        End With
        ' clear last of data
        TempPlayer(index).c_mapNpcNum = 0
        TempPlayer(index).c_mapNum = 0
        ' exit out early so we don't send chat update twice
        Exit Sub
    End If
    
    ' send update to the client
    sendChat index
End Sub

Public Sub sendChat(ByVal index As Long)
    Dim convNum As Long
    Dim curChat As Long
    Dim mainText As String
    Dim optText(1 To 4) As String
    Dim P_GENDER As String
    Dim P_NAME As String
    Dim P_CLASS As String
    Dim i As Long
    
    If TempPlayer(index).inChatWith > 0 Then
        convNum = NPC(TempPlayer(index).inChatWith).Conv
        curChat = TempPlayer(index).curChat
        
        ' check for unique events and trigger them early
        If Conv(convNum).Conv(curChat).Event > 0 Then
            Select Case Conv(convNum).Conv(curChat).Event
                Case 1 ' Open Shop
                    If Conv(convNum).Conv(curChat).Data1 > 0 Then ' shop exists?
                        SendOpenShop index, Conv(convNum).Conv(curChat).Data1
                        TempPlayer(index).InShop = Conv(convNum).Conv(curChat).Data1 ' stops movement and the like
                    End If
                    ' exit the chat
                    TempPlayer(index).inChatWith = 0
                    TempPlayer(index).curChat = 0
                    ' send chat update
                    sendChat index
                    ' send npc dir
                    With MapNpc(TempPlayer(index).c_mapNum).NPC(TempPlayer(index).c_mapNpcNum)
                        If .c_inChatWith = index Then
                            .c_inChatWith = 0
                            .Dir = .c_lastDir
                            NpcDir TempPlayer(index).c_mapNum, TempPlayer(index).c_mapNpcNum, .Dir
                        End If
                    End With
                    ' clear last of data
                    TempPlayer(index).c_mapNpcNum = 0
                    TempPlayer(index).c_mapNum = 0
                    ' exit out early so we don't send chat update twice
                    Exit Sub
                Case 2 ' Open Bank
                    SendBank index
                Case 3 ' Give Item
                    Dim b As Long
                    b = FindOpenInvSlot(index, Conv(convNum).Conv(curChat).Data1)
                    Call SetPlayerInvItemNum(index, b, Conv(convNum).Conv(curChat).Data1)
                    Call SetPlayerInvItemValue(index, b, Conv(convNum).Conv(curChat).Data2)
                    SendInventory index
            End Select
        End If

continue:
        ' cache player's details
        If Player(index).Sex = SEX_MALE Then
            P_GENDER = "man"
        Else
            P_GENDER = "woman"
        End If
        P_NAME = Trim$(Player(index).Name)
        P_CLASS = Trim$(Class(Player(index).Class).Name)
        
        mainText = Conv(convNum).Conv(curChat).Conv
        For i = 1 To 4
            optText(i) = Conv(convNum).Conv(curChat).rText(i)
        Next
    End If
    
    SendChatUpdate index, TempPlayer(index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
End Sub
