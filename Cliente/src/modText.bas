Attribute VB_Name = "modText"
Option Explicit

' Text declares
Private Declare Function CreateFont Lib "GDI32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal e As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function SetBkMode Lib "GDI32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "GDI32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "GDI32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long

Public HideIncomingMessages As Boolean


' Used to set a font for GDI text drawing
Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "Verdana")
    frmMain.Font = "Verdana"
    frmMain.FontSize = Size - 5

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetFont", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
' GDI text drawing onto buffer
Public Sub DrawText(ByVal hDC As Long, ByVal X, ByVal y, ByVal text As String, Color As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, 0)
    Call TextOut(hDC, X + 1, y + 1, text, Len(text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, X, y, text, Len(text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayerName(ByVal index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim gColor As Long
Dim Name As String
Dim Text2X As Long
Dim Text2Y As Long
Dim GuildString As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if player is killer
    Color = GetPlayerNameColorByJustice(index)
    
    ' Check access level
    If GetPlayerAccess(index) > 0 Then
        Select Case GetPlayerAccess(index)
            Case 1
                Color = RGB(25, 200, 180)
            Case 2
                Color = RGB(100, 255, 0)
            Case 3
                Color = RGB(0, 155, 255)
            Case 4
                Color = RGB(100, 50, 255)
        End Select
    End If

    If GetPlayerAccess(index) = 0 Then
            Name = (Player(index).Name)
        ElseIf GetPlayerAccess(index) = 1 Then
            Name = "[Epic] " & (Player(index).Name)
        ElseIf GetPlayerAccess(index) = 2 Then
            Name = "[MOD] " & (Player(index).Name)
        ElseIf GetPlayerAccess(index) = 3 Then
            Name = "[MAP] " & (Player(index).Name)
        ElseIf GetPlayerAccess(index) = 4 Then
            Name = "[GM] " & (Player(index).Name)
    End If

    ' calc pos
    TextX = ConvertMapX(GetPlayerX(index) * PIC_X) + Player(index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, ((Name))) + Len(Name) ^ 1.2
    GuildString = Player(index).GuildName
    Text2X = ConvertMapX(GetPlayerX(index) * PIC_X) + Player(index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, ((GuildString))) + Len(GuildString) ^ 1.2
        
    If GetPlayerSprite(index) < 1 Or GetPlayerSprite(index) > NumCharacters Then
        TextY = ConvertMapY(GetPlayerY(index) * PIC_Y) + Player(index).YOffset - 16
        Text2Y = ConvertMapY(GetPlayerY(index) * PIC_Y) + Player(index).YOffset
    Else
        ' Determine location for text
        TextY = ConvertMapY(GetPlayerY(index) * PIC_Y) + Player(index).YOffset - (DDSD_Character(GetPlayerCenterSprite(index)).lHeight / 4)
        Text2Y = ConvertMapY(GetPlayerY(index) * PIC_Y) + Player(index).YOffset - (DDSD_Character(GetPlayerCenterSprite(index)).lHeight / 4) + 16
    End If

    ' Draw name
    If Not GuildString = vbNullString Then
        Call DrawText(TexthDC, Text2X, Text2Y, GuildString, Color)
        Call DrawText(TexthDC, TextX, TextY + 1, Name, Color)
    Else
        Call DrawText(TexthDC, TextX, TextY + 12, Name, Color)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub DrawPlayerLevel(ByVal index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim Level As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if player is killer
        Color = GetPlayerNameColorByJustice(index)
        
    ' Check access level
    If GetPlayerAccess(index) > 0 Then
        Select Case GetPlayerAccess(index)
            Case 1
                Color = RGB(25, 200, 180)
            Case 2
                Color = RGB(100, 255, 0)
            Case 3
                Color = RGB(0, 155, 255)
            Case 4
                Color = RGB(100, 50, 255)
        End Select
    End If

    If GetPlayerAccess(index) = 0 Then
         Level = "Lvl." & GetPlayerLevel(index)
        ElseIf GetPlayerAccess(index) = 1 Then
            Level = "Lvl." & GetPlayerLevel(index)
        ElseIf GetPlayerAccess(index) = 2 Then
         Level = "Lvl." & GetPlayerLevel(index)
        ElseIf GetPlayerAccess(index) = 3 Then
            Level = "Lvl." & GetPlayerLevel(index)
        ElseIf GetPlayerAccess(index) = 4 Then
         Level = "Lvl." & GetPlayerLevel(index)
    End If

    ' calc pos
    TextX = ConvertMapX(GetPlayerX(index) * PIC_X) + Player(index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, ((Level))) + Len(Level) ^ 1.2
    If GetPlayerSprite(index) < 1 Or GetPlayerSprite(index) > NumCharacters Then
        TextY = ConvertMapY(GetPlayerY(index) * PIC_Y) + Player(index).YOffset - 4
    Else
        ' Determine location for text
        TextY = ConvertMapY(GetPlayerY(index) * PIC_Y) + Player(index).YOffset - (DDSD_Character(GetPlayerCenterSprite(index)).lHeight / 4) - 12
    End If

    ' Draw level
    If Not Player(index).GuildName = vbNullString Then
        Call DrawText(TexthDC, TextX, TextY, Level, Color)
    Else
        Call DrawText(TexthDC, TextX, TextY + 12, Level, Color)
    End If
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerLevel", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcName(ByVal index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim Name As String
Dim NPCNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NPCNum = MapNpc(index).num
    Name = Trim$(NPC(NPCNum).TranslatedName)
    
    Select Case NPC(NPCNum).Behaviour
        Case NPC_BEHAVIOUR_ATTACKONSIGHT
            Color = QBColor(BrightRed)
        Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
            Color = QBColor(Yellow)
        Case NPC_BEHAVIOUR_GUARD
            Color = QBColor(Grey)
        Case NPC_BEHAVIOUR_BLADE
            Color = QBColor(BrightRed)
            Name = ""
        Case NPC_BEHAVIOUR_SLIDE
            Color = QBColor(Black)
            Name = ""
        Case Else
            Color = QBColor(BrightGreen)
    End Select
    
    Select Case IsMapNPCaPet(index)
        Case True
            Color = QBColor(BrightGreen)
    End Select


    
    TextX = ConvertMapX(MapNpc(index).X * PIC_X) + MapNpc(index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, ((Name))) + LenB(Name)
    If NPC(NPCNum).sprite < 1 Or NPC(NPCNum).sprite > NumCharacters Then
        TextY = ConvertMapY(MapNpc(index).y * PIC_Y) + MapNpc(index).YOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(MapNpc(index).y * PIC_Y) + MapNpc(index).YOffset - (DDSD_Character(NPC(NPCNum).sprite).lHeight / 4) + 28
    End If

    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, Name, Color)
        'Alatar v1.2
    Dim i As Long
    
    For i = 1 To MAX_QUESTS
        'check if the npc is the next task to any quest: [?] symbol
        If Quest(i).Name <> "" Then
            If Player(MyIndex).PlayerQuest(i).Status = QUEST_STARTED Then
                If Quest(i).Task(Player(MyIndex).PlayerQuest(i).ActualTask).NPC = NPCNum Then
                    Name = "[?]"
                    TextX = ConvertMapX(MapNpc(index).X * PIC_X) + MapNpc(index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, ((Name))) + 6
                    If NPC(NPCNum).sprite < 1 Or NPC(NPCNum).sprite > NumCharacters Then
                        TextY = ConvertMapY(MapNpc(index).y * PIC_Y) + MapNpc(index).YOffset - 16
                    Else
                        TextY = ConvertMapY(MapNpc(index).y * PIC_Y) + MapNpc(index).YOffset - (DDSD_Character(NPC(NPCNum).sprite).lHeight / 4)
                    End If
                    Call DrawText(TexthDC, TextX, TextY, Name, QBColor(Yellow))
                    Exit For
                End If
            End If
            
            'check if the npc is the starter to any quest: [!] symbol
            'can accept the quest as a new one?
            If Player(MyIndex).PlayerQuest(i).Status = QUEST_NOT_STARTED Or Player(MyIndex).PlayerQuest(i).Status = QUEST_COMPLETED_BUT Then
                'the npc gives this quest?
                If NPC(NPCNum).QuestNum = i Then
                    Name = "[!]"
                    TextX = ConvertMapX(MapNpc(index).X * PIC_X) + MapNpc(index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, ((Name))) + 6
                    If NPC(NPCNum).sprite < 1 Or NPC(NPCNum).sprite > NumCharacters Then
                        TextY = ConvertMapY(MapNpc(index).y * PIC_Y) + MapNpc(index).YOffset - 16
                    Else
                        TextY = ConvertMapY(MapNpc(index).y * PIC_Y) + MapNpc(index).YOffset - (DDSD_Character(NPC(NPCNum).sprite).lHeight / 4) + 12
                    End If
                    Call DrawText(TexthDC, TextX, TextY, Name, QBColor(Yellow))
                    Exit For
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpcName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcLevel(ByVal index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim Level As String
Dim NPCNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NPCNum = MapNpc(index).num
    Level = (NPC(NPCNum).Level)
    
    Select Case NPC(NPCNum).Behaviour
        Case NPC_BEHAVIOUR_ATTACKONSIGHT
            Color = QBColor(BrightRed)
        Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
            Color = QBColor(Yellow)
        Case Else
            Exit Sub
    End Select
    
    TextX = ConvertMapX(MapNpc(index).X * PIC_X) + MapNpc(index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, ((Level)))
    If NPC(NPCNum).sprite < 1 Or NPC(NPCNum).sprite > NumCharacters Then
        TextY = ConvertMapY(MapNpc(index).y * PIC_Y) + MapNpc(index).YOffset
    Else
        ' Determine location for text
        TextY = ConvertMapY(MapNpc(index).y * PIC_Y) + MapNpc(index).YOffset - (DDSD_Character(NPC(NPCNum).sprite).lHeight / 4) + 8
    End If

    ' Draw Level
    Call DrawText(TexthDC, TextX, TextY, Level, Color)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpcLevel", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function BltMapAttributes()
    Dim X As Long
    Dim y As Long
    Dim tX As Long
    Dim tY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optAttribs.value Then
        For X = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, y) Then
                    With map.Tile(X, y)
                        tX = ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5)
                        tY = ((ConvertMapY(y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                DrawText TexthDC, tX, tY, "B", QBColor(BrightRed)
                            Case TILE_TYPE_WARP
                                DrawText TexthDC, tX, tY, "W", QBColor(BrightBlue)
                            Case TILE_TYPE_ITEM
                                DrawText TexthDC, tX, tY, "I", QBColor(White)
                            Case TILE_TYPE_NPCAVOID
                                DrawText TexthDC, tX, tY, "N", QBColor(White)
                            Case TILE_TYPE_KEY
                                DrawText TexthDC, tX, tY, "K", QBColor(White)
                            Case TILE_TYPE_KEYOPEN
                                DrawText TexthDC, tX, tY, "O", QBColor(White)
                            Case TILE_TYPE_RESOURCE
                                DrawText TexthDC, tX, tY, "O", QBColor(Green)
                            Case TILE_TYPE_DOOR
                                DrawText TexthDC, tX, tY, "D", QBColor(Brown)
                            Case TILE_TYPE_NPCSPAWN
                                DrawText TexthDC, tX, tY, "S", QBColor(Yellow)
                            Case TILE_TYPE_SHOP
                                DrawText TexthDC, tX, tY, "S", QBColor(BrightBlue)
                            Case TILE_TYPE_BANK
                                DrawText TexthDC, tX, tY, "B", QBColor(Blue)
                            Case TILE_TYPE_HEAL
                                DrawText TexthDC, tX, tY, "H", QBColor(BrightGreen)
                            Case TILE_TYPE_TRAP
                                DrawText TexthDC, tX, tY, "T", QBColor(BrightRed)
                            Case TILE_TYPE_SLIDE
                                DrawText TexthDC, tX, tY, "S", QBColor(BrightCyan)
                            Case TILE_TYPE_SCRIPT
                                DrawText TexthDC, tX, tY, "Sc", QBColor(Yellow)
                            Case TILE_TYPE_ICE
                                DrawText TexthDC, tX, tY, "Ic", QBColor(BrightCyan)
                        End Select
                    End With
                End If
            Next
        Next
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "BltMapAttributes", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub BltActionMsg(ByVal index As Long)
    Dim X As Long, y As Long, i As Long, Time As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' does it exist
    If ActionMsg(index).Created = 0 Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(index).Type
        Case ACTIONMSG_STATIC
            Time = 1500

            If ActionMsg(index).y > 0 Then
                X = ActionMsg(index).X + Int(PIC_X \ 2) - ((Len((ActionMsg(index).message)) \ 2) * 8)
                y = ActionMsg(index).y - Int(PIC_Y \ 2) - 2
            Else
                X = ActionMsg(index).X + Int(PIC_X \ 2) - ((Len((ActionMsg(index).message)) \ 2) * 8)
                y = ActionMsg(index).y - Int(PIC_Y \ 2) + 8
            End If

        Case ACTIONMSG_SCROLL
            Time = 1500
        
            If ActionMsg(index).y > 0 Then
                X = ActionMsg(index).X + Int(PIC_X \ 2) - ((Len((ActionMsg(index).message)) \ 2) * 8)
                y = ActionMsg(index).y - Int(PIC_Y \ 2) - 2 - (ActionMsg(index).Scroll * 0.6)
                ActionMsg(index).Scroll = ActionMsg(index).Scroll + 1
            Else
                X = ActionMsg(index).X + Int(PIC_X \ 2) - ((Len((ActionMsg(index).message)) \ 2) * 8)
                y = ActionMsg(index).y - Int(PIC_Y \ 2) + 18 + (ActionMsg(index).Scroll * 0.6)
                ActionMsg(index).Scroll = ActionMsg(index).Scroll + 1
            End If

        Case ACTIONMSG_SCREEN
            Time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                    If i <> index Then
                        ClearActionMsg index
                        index = i
                    End If
                End If
            Next
            X = (frmMain.picScreen.Width \ 2) - ((Len((ActionMsg(index).message)) \ 2) * 8)
            y = 425

    End Select
    
    X = ConvertMapX(X)
    y = ConvertMapY(y)

    If GetTickCount < ActionMsg(index).Created + Time Then
        Call DrawText(TexthDC, X, y, ActionMsg(index).message, QBColor(ActionMsg(index).Color))
    Else
        ClearActionMsg index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltActionMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function getWidth(ByVal DC As Long, ByVal text As String) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    getWidth = frmMain.TextWidth(text) \ 2
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "getWidth", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub AddText(ByVal msg As String, ByVal Color As Integer, Optional blTranslate As Boolean = False)
Dim s As String

    If blTranslate = True Then msg = GetTranslation(msg)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    s = vbNewLine & msg
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    frmMain.txtChat.SelColor = QBColor(Color)
    frmMain.txtChat.SelText = s
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    
    'LoadOptions
    
    If Options.ChatToScreen = 1 Then
        frmMain.txtChat.Visible = True
    ElseIf Options.ChatToScreen = 2 Then
        Call ReOrderChat(msg, QBColor(Color))
        frmMain.txtChat.Visible = False
    Else
        frmMain.txtChat.Visible = False
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WordWrap_Array(ByVal text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
Dim lineCount As Long, i As Long, Size As Long, lastSpace As Long, b As Long
    
    'Too small of text
    If Len(text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = text
        Exit Sub
    End If
    
    ' default values
    b = 1
    lastSpace = 1
    Size = 0
    
    For i = 1 To Len(text)
        ' if it's a space, store it
        Select Case Mid$(text, i, 1)
            Case " ": lastSpace = i
            Case "_": lastSpace = i
            Case "-": lastSpace = i
        End Select
        
        'Add up the size
        Size = Size + getWidth(TexthDC, Mid$(text, i, 1))
        
        'Check for too large of a size
        If Size > MaxLineLen Then
            'Check if the last space was too far back
            If i - lastSpace > 12 Then
                'Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, b, (i - 1) - b))
                b = i - 1
                Size = 0
            Else
                'Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, b, lastSpace - b))
                b = lastSpace + 1
                
                'Count all the words we ignored (the ones that weren't printed, but are before "i")
                Size = getWidth(TexthDC, Mid$(text, lastSpace, i - lastSpace))
            End If
        End If
        
        ' Remainder
        If i = Len(text) Then
            If b <> i Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(text, b, i)
            End If
        End If
    Next
End Sub

Public Function WordWrap(ByVal text As String, ByVal MaxLineLen As Integer) As String
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim Size As Long
Dim i As Long
Dim b As Long

    'Too small of text
    If Len(text) < 2 Then
        WordWrap = text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(text, vbNewLine)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        Size = 0
        b = 1
        lastSpace = 1
        
        'Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine
        
        'Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Or InStr(1, TempSplit(TSLoop), "-") Or InStr(1, TempSplit(TSLoop), "_") Then
            
            'Loop through all the characters
            For i = 1 To Len(TempSplit(TSLoop))
            
                'If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), i, 1)
                    Case " ": lastSpace = i
                    Case "_": lastSpace = i
                    Case "-": lastSpace = i
                End Select
    
                'Add up the size
                Size = Size + getWidth(TexthDC, Mid$(TempSplit(TSLoop), i, 1))
 
                'Check for too large of a size
                If Size > MaxLineLen Then
                    'Check if the last space was too far back
                    If i - lastSpace > 12 Then
                        'Too far away to the last space, so break at the last character
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)) & vbNewLine
                        b = i - 1
                        Size = 0
                    Else
                        'Break at the last space to preserve the word
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, lastSpace - b)) & vbNewLine
                        b = lastSpace + 1
                        
                        'Count all the words we ignored (the ones that weren't printed, but are before "i")
                        Size = getWidth(TexthDC, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                    End If
                End If
                
                'This handles the remainder
                If i = Len(TempSplit(TSLoop)) Then
                    If b <> i Then
                        WordWrap = WordWrap & Mid$(TempSplit(TSLoop), b, i)
                    End If
                End If
            Next i
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop
End Function

' CHANGE FONT FIX
' PLEASE NOTE THIS WILL FAIL MISERABLY IF YOU DIDN'T APPLY THE FONT MEMORY LEAK FIX FIRST
' CHAT BUBBLE HACK
' I ONLY DID THIS COZ THE CHATBUBBLE TEXT LOOKS BETTER WITHOUT SHADOW OVER WHITE BUBBLES!
Public Sub DrawTextNoShadow(ByVal hDC As Long, ByVal X, ByVal y, ByVal text As String, Color As Long)
    ' If debug mode, handle error then exit out
    Dim OldFont As Long ' HFONT
    
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call SetFont(FONT_NAME, FONT_SIZE)
    'OldFont = SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, X, y, text, Len(text))
    'Call SelectObject(hDC, OldFont)
    Call DeleteObject(GameFont)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTextNoShadow", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearChat()
    If Not InGame Then Exit Sub
    frmMain.txtChat.text = vbNullString
    DisplayChat
End Sub

Sub DisplayChat()
    Dim chatroom As Byte
    For chatroom = 1 To ChatType_Count - 1
        If IsChatActivated(chatroom) Then
            Call ListBegin(ChatRooms(chatroom))
        Else
            Call ListEnd(ChatRooms(chatroom))
        End If
    Next
    
    Dim alldone As Boolean
    alldone = False
    While Not alldone
        chatroom = CompareChatRooms
        If chatroom = 0 Then
            alldone = True 'we reached all list's end or no chat activated
        Else
            Call DisplayChatRoomMsg(chatroom)
            Call ListNext(ChatRooms(chatroom))
        End If
    Wend
    
    
End Sub

Sub DisplayChatRoomMsg(ByVal chatroom As Byte)
    If chatroom > 0 And chatroom < ChatType_Count Then
        Dim msg As ChatMsgRec
        msg = ListActual(ChatRooms(chatroom))
        frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
        frmMain.txtChat.SelColor = msg.colour
        frmMain.txtChat.SelText = vbNewLine & msg.header
        frmMain.txtChat.SelColor = msg.saycolour
        frmMain.txtChat.SelText = msg.text
        frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    End If
End Sub

Function CompareChatRooms() As Byte
    Dim BestOption As Byte
    BestOption = 0
    Dim i As Byte
    For i = 1 To ChatType_Count - 1
        If IsChatActivated(i) And Not ListEnd(ChatRooms(i)) Then
            If BestOption = 0 Then
                BestOption = i
            Else
                BestOption = Compare2ChatRooms(BestOption, i)
            End If
        End If
    Next
    CompareChatRooms = BestOption
End Function

Function Compare2ChatRooms(ByVal firstroom As Byte, ByVal secondroom As Byte) As Byte
    Dim msg1 As ChatMsgRec
    Dim msg2 As ChatMsgRec
    
    
    msg1 = ListActual(ChatRooms(firstroom))
    msg2 = ListActual(ChatRooms(secondroom))
    
    If msg1.ArrivedAt < msg2.ArrivedAt Then
        Compare2ChatRooms = firstroom
    Else
        Compare2ChatRooms = secondroom
    End If
End Function


Function IsChatActivated(ByVal chatroom As Byte) As Boolean
    If chatroom > 0 And chatroom < ChatType_Count Then
        IsChatActivated = Options.ActivatedChats(chatroom)
    End If
End Function

Sub ChatOptionsInit()
    With frmChatDisplay
    Dim i As Byte
    For i = .chkChat.LBound To .chkChat.UBound
        .chkChat(i).value = BTI(Options.ActivatedChats(i + 1))
        .chkChat(i).Caption = ChatTypeToStr(i + 1)
    Next
    End With
End Sub

Function ChatTypeToStr(ByVal chatroom As ChatType) As String
    Select Case chatroom
        Case MapChat
            ChatTypeToStr = "Mapa"
        Case GlobalChat
            ChatTypeToStr = "Global"
        Case PartyChat
            ChatTypeToStr = "Party"
        Case ClanChat
            ChatTypeToStr = "Clan"
        Case WhisperChat
            ChatTypeToStr = "Susurros"
        Case SystemChat
            ChatTypeToStr = "Sistema"
    End Select
End Function


Function CanMsgBeDisplayed(ByVal ChatType As Byte) As Boolean
    If IsChatActivated(ChatType) And Not HideIncomingMessages Then
        CanMsgBeDisplayed = True
    End If
End Function
Public Sub DrawChat()
Dim i As Integer
    For i = 1 To 8
        Call DrawText(TexthDC, Camera.Left + 10, (Camera.Bottom - 20) - (i * 20), Chat(i).text, Chat(i).colour)
    Next
End Sub

Public Sub ReOrderChat(ByVal nText As String, nColour As Long)
Dim i As Integer
   
    For i = 19 To 1 Step -1
        Chat(i + 1).text = Chat(i).text
        Chat(i + 1).colour = Chat(i).colour
    Next
   
    Chat(1).text = nText
    Chat(1).colour = QBColor(White)

End Sub
