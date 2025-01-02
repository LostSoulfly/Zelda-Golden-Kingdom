Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer

Public Sub CheckKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckInputKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If

'space to pickup items
If Options.WASD = 1 Then
    If ChatFocus = False Then
        If GetKeyState(vbKeySpace) < 0 Then
            CheckMapGetItem
            ProcessAttack True
        End If
        'E key to attack
        If GetKeyState(vbKeyE) < 0 Then
            ControlDown = True
        End If
    End If
End If

    If GetKeyState(vbKeyReturn) < 0 Then
        CheckMapGetItem
    End If
    
    If Player(MyIndex).onIce Then
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
    Else

        'Move Up
        If GetKeyState(vbKeyUp) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirUp = False
        End If

        'Move Right
        If GetKeyState(vbKeyRight) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
            Exit Sub
        Else
            DirRight = False
        End If

        'Move down
        If GetKeyState(vbKeyDown) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirDown = False
        End If

        'Move left
        If GetKeyState(vbKeyLeft) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
            Exit Sub
        Else
            DirLeft = False
        End If
        
    End If
    ' Error handler
    
    'If we aren't focused on the chat, let's see if we use W/A/S/D to move
    If ChatFocus = False Then
        'Move Up (W)
        If GetKeyState(vbKeyW) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirUp = False
        End If
    
        'Move Right (D)
        If GetKeyState(vbKeyD) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
            Exit Sub
        Else
            DirRight = False
        End If
    
        'Move down (S)
        If GetKeyState(vbKeyS) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirDown = False
        End If
    
        'Move left (A)
        If GetKeyState(vbKeyA) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
            Exit Sub
        Else
            DirLeft = False
        End If

    End If
    
    Exit Sub
errorhandler:
    HandleError "CheckInputKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim Name As String
Dim i As Long
Dim N As Long
Dim Command() As String
Dim buffer As clsBuffer
'Kill Counter
Dim totalkills As Long
Dim totaldeaths As Long
Dim combatdeaths As Long
Dim alldeaths As Long
'SendRequestPlayerData

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ChatText = Trim$(MyText)

    If LenB(ChatText) = 0 Then Exit Sub
    MyText = LCase$(ChatText)

    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
    
    

        ' Emote message
        If Left$(ChatText, 1) = "-" Then
            MyText = Mid$(ChatText, 2, Len(ChatText) - 1)

            If Len(ChatText) > 0 Then
                Call EmoteMsg(ChatText)
            End If

            MyText = vbNullString
            frmMain.txtMyChat.text = vbNullString
            Exit Sub
        End If

        ' Player message
        If Left$(ChatText, 1) = "!" Then
            'Exit Sub
            MyText = Mid$(ChatText, 2, Len(ChatText) - 1)
            Name = vbNullString

            ' Get the desired player from the user text
            For i = 1 To Len(MyText)

                If Mid$(MyText, i, 1) <> Space(1) Then
                    Name = Name & Mid$(MyText, i, 1)
                Else
                    Exit For
                End If

            Next
            
            If Not Len(MyText) > 0 Then Exit Sub

            MyText = Trim$(Mid$(MyText, i, Len(MyText) - 1))

            ' Make sure they are actually sending something
            If Len(MyText) > 0 Then
                'MyText = Mid$(ChatText, i + 1, Len(MyText) - i)
                ' Send the message to the player
                Call PlayerMsg(MyText, Name)
            Else
                Call AddText("Use !playername (message)", AlertColor, True)
            End If

            MyText = vbNullString
            frmMain.txtMyChat.text = vbNullString
            Exit Sub
        End If

        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))

            Select Case Command(0)
            
                Case "/ss"
                    ScreenshotMap
            
                Case "/wasd"
                    If Options.WASD = 0 Then Options.WASD = 1 Else Options.WASD = 0
                    If Options.WASD = 1 Then AddText "WASD controls enabled! Use E to attack, Spacebar to pickup items.", White
                    If Options.WASD = 0 Then AddText "WASD controls disabled.", White
                    
                Case "/help"
                    Call AddText("Social Commands:", HelpColor, True)
                    Call AddText("-msghere = message in color", HelpColor, True)
                    Call AddText("!nickname message = message to Player", HelpColor, True)
                    Call AddText("Available Commands: /info, /online, /fps, /fpslock, /deaths, /counter", HelpColor)
                    Call AddText("/wasd, /guild, /stats, /translation", HelpColor)
                    Call AddText("/minimap to toggle minimap mode.", HelpColor)
                    Call AddText("/minimap on/off to enable/disable minimap", HelpColor)
                
                Case "/guild"
                    If UBound(Command) < 1 Then
                            Call AddText("Guild Commands:", HelpColor, True)
                            Call AddText("Creat Guild: /guild create (GuildName)", HelpColor, True)
                            Call AddText("To transfer founder data usa /guild founder (Name)", HelpColor, True)
                            Call AddText("Invite to the Guild: /guild invite (name)", HelpColor, True)
                            Call AddText("Quit Guild: /guild quit", HelpColor, True)
                            Call AddText("Open Guild Dashboard: /guild admin", HelpColor, True)
                            Call AddText("Expel from Guild: /guild kick (name)", HelpColor, True)
                            Call AddText("Disband Guild: /guild disband yes", HelpColor)
                            Call AddText("View Guild Members: /guild online (online/all/offline)", HelpColor, True)
                            GoTo Continue
                    End If
                    
                    Select Case Command(1)
                        Case "create"
                        If (Len(ChatText) - Len("/guild create ")) < 30 And (Len(ChatText) - Len("/guild create ")) > 3 Then
                            'If UBound(Command) = 2 Then
                                Call GuildCommand(1, Right(ChatText, Len(ChatText) - Len("/guild create ")))
                            Else
                                Call AddText("You must use /guild create <name>. Must be less than 30 characters!", BrightRed)
                            End If

                        Case "invite"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(2, Command(2))
                            Else
                                Call AddText("You must select a user, use /guild invite (Name)", BrightRed, True)
                            End If

                        Case "quit"
                            Call GuildCommand(3, "")

                        Case "admin"
                            Call GuildCommand(4, "")

                        Case "online"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(5, Command(2))
                            Else
                                Call GuildCommand(5, "")
                            End If

                        Case "accept"
                            Call GuildCommand(6, "")

                        Case "decline"
                            Call GuildCommand(7, "")

                        Case "founder"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(8, Command(2))
                            Else
                                Call AddText("You must select a user, usa /guild founder (Name)", BrightRed, True)
                            End If
                        Case "kick"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(9, Command(2))
                            Else
                                Call AddText("You must select a user, use /guild kick (Name)", BrightRed, True)
                            End If
                        Case "disband"
                            If UBound(Command) = 2 Then
                                If LCase(Command(2)) = LCase("yes") Then
                                    Call GuildCommand(10, "")
                                Else
                                    Call AddText("Type something like /guild disband yes (This is to avoid an accident!)", BrightRed, True)
                                End If
                            Else
                                Call AddText("Type something like /guild disband yes (This is to avoid an accident!)", BrightRed, True)
                            End If

                        End Select
                        
                Case "/translation"
				Case "/translog"
                    frmTransLog.Visible = Not (frmTransLog.Visible)
                        
                Case "/translate"
                    'possibly translate your message and send it? easy enough.
                    
                Case "/minimap"
                If UBound(Command) = 0 Then
                    If Options.MiniMapBltElse = 0 Then Options.MiniMapBltElse = 1 Else Options.MiniMapBltElse = 0
                    If Options.MiniMapBltElse = 0 Then AddText "Minimap performance mode enabled!", White
                    If Options.MiniMapBltElse = 1 Then AddText "Minimap performance mode disabled!", White
                    'GoTo Continue
                Else
                    Select Case Trim$(StringIntersection(MyText, "/minimap"))
                    
                    Case Is = "off"
                    Options.MiniMap = 0
                    Case Is = 0
                    Options.MiniMap = 0
                    Case Is = "on"
                    Options.MiniMap = 1
                    Case Is = 1
                    Options.MiniMap = 1
                    End Select
                End If
             
                Case "/info"

                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        AddText "Uso: /info (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Uso: /info (name)", AlertColor
                        GoTo Continue
                    End If

                    Set buffer = New clsBuffer
                    buffer.WriteLong CPlayerInfoRequest
                    buffer.WriteString Trim$(StringIntersection(MyText, "/info"))
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                    ' Whos Online
                Case "/online"
                    SendWhosOnline
                    ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS
                    ' toggle fps lock
                Case "/fpslock"
                    FPS_Lock = Not FPS_Lock
                    ' Request stats
                Case "/stats"
                    Set buffer = New clsBuffer
                    buffer.WriteLong CGetStats
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                    'Kill Counter
                Case "/deaths"
                    totalkills = Player(MyIndex).Kill + Player(MyIndex).NpcKill
                    Call AddText("-Counter of Committed Death-", DarkGrey, True)
                    Call AddText("Player Deaths: " + Str(Player(MyIndex).Kill), White)
                    Call AddText("Creature Kills: " + Str(Player(MyIndex).NpcKill), White)
                    Call AddText("Total Kills: " & " " + Str(totalkills), White)
                Case "/counter"
                    combatdeaths = Player(MyIndex).Dead + Player(MyIndex).NpcDead
                    alldeaths = combatdeaths + Player(MyIndex).EnviroDead
                    Call AddText("-Counter of Death Suffered-", DarkGrey, True)
                    Call AddText("Killed by Players: " + Str(Player(MyIndex).Dead), White)
                    Call AddText("Killed by creatures: " + Str(Player(MyIndex).NpcDead), White)
                    Call AddText("Total Combat Deaths: " + Str(combatdeaths), White)
                    Call AddText("Accidental Deaths: " + Str(Player(MyIndex).EnviroDead), White)
                    Call AddText("Total Deaths: " + Str(alldeaths), White)
                    ' // Monitor Admin Commands //
                    ' Admin Help
                Case "/admin"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue
                    frmMain.picAdmin.Visible = Not frmMain.picAdmin.Visible
                    If frmMain.picAdmin.Visible = True Then frmMain.Width = 14835 Else frmMain.Width = 12090
                    ' Kicking a player
                Case "/kick"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue


                    If UBound(Command) < 1 Then
                        AddText "Use /kick (name)", AlertColor, True
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Use /kick (name)", AlertColor, True
                        GoTo Continue
                    End If
                    
                    SendKick Trim$(StringIntersection(MyText, "/kick"))
                    ' // Mapper Admin Commands //
                    ' Location
                Case "/loc"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue

                    BLoc = Not BLoc
                    ' Map Editor
                Case "/editmap"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                    SendRequestEditMap
                    
                    ' Warping to a player
                Case "/warpmeto"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Use /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Use /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpMeTo Trim$(StringIntersection(MyText, "/warpmeto"))
                    ' Warping a player to you
                Case "/warptome"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Use /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Use /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpToMe Trim$(StringIntersection(MyText, "/warptome"))
                    ' Warping to a map
                Case "/warpto"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Use /warpto (map #)", AlertColor, True
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Use /warpto (map #)", AlertColor, True
                        GoTo Continue
                    End If

                    N = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If N > 0 And N <= MAX_MAPS Then
                        Call WarpTo(N)
                    Else
                        Call AddText("Invalid map number.", Red, True)
                    End If
                'visibility toggle
                Case "/visible"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                    SendVisibility
                    ' Setting sprite
                Case "/setsprite"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Use /setsprite (sprite #)", AlertColor, True
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Use /setsprite (sprite #)", AlertColor, True
                        GoTo Continue
                    End If
                    SendSetSprite CLng(Command(1)), GetPlayerName(MyIndex)
                    
                    ' Map report
                Case "/mapreport"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapReport
                    GoTo Continue
                    ' Respawn request
                Case "/respawn"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapRespawn
                    ' MOTD change
                Case "/motd"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Use /motd (new motd)", AlertColor
                        GoTo Continue
                    End If

                    SendMOTDChange Right$(ChatText, Len(ChatText) - 5)
                    ' Check the ban list
                Case "/banlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendBanList
                    ' // Developer Admin Commands //
                    ' Editing item request
                    
                Case "/ban2"
                    If UBound(Command) < 2 Then
                        AddText "Use /ban (password) (name)", AlertColor
                        GoTo Continue
                    End If

                    SendSpecialBan Command(1), Command(2)
                    ' // Developer Admin Commands //

                    ' Banning a player
                Case "/ban"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Use /ban (name)", AlertColor
                        GoTo Continue
                    End If
                    ' Editing item request
                    SendBan Trim$(StringIntersection(MyText, "/ban"))
                Case "/edititem"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditItem
                ' Editing animation request
                Case "/editanimation"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditAnimation
                    ' Editing npc request
                Case "/editnpc"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditNpc
                Case "/editresource"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditResource
                    ' Editing shop request
                Case "/editshop"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditShop
                    ' Editing spell request
                Case "/editspell"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditSpell
                    
                Case "/editquest"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                    SendRequestEditQuest
                
                    ' // Creator Admin Commands //
                    
                    ' Giving another player access
                Case "/setaccess"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    If UBound(Command) < 2 Then
                        AddText "Use /setaccess (name) (access)", AlertColor, True
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Use /setaccess (name) (access)", AlertColor, True
                        GoTo Continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))
                Case "/needaccounts"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue
                    
                    If UBound(Command) < 1 Then
                        AddText "Use /needaccounts (password)", AlertColor, True
                        GoTo Continue
                    End If
                    
                    SendNeedAccounts Command(1)
                Case "/makeadmin"
                    If UBound(Command) < 1 Then
                        GoTo Continue
                    End If
                    
                    SendMakeAdmin Command(1)
                Case "/addip"
                    If UBound(Command) < 1 Then
                        GoTo Continue
                    End If
                    SendAddIP Command(1)
                    
                Case "/bugreport"
                    
                    Dim strReport As String
                    strReport = Trim$(Replace(MyText, "/bugreport", ""))
                    
                    If Len(strReport) <= 1 Then
                    AddText "Usage: /bugreport <Describe the issue you are experiencing>", White
                    AddText "Your IP address, character's name, current map, and coordinates will be recorded.", White
                    GoTo Continue
                    End If
                    
                    
                    SendBugReport strReport
                    
                Case "/mute"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                    
                    If UBound(Command) < 2 Then
                        AddText "Use /mute (Name) (time[s])", AlertColor, True
                        GoTo Continue
                    End If
                    
                    If Not IsNumeric(Command(2)) Then
                        AddText "Time has to be a number", AlertColor, True
                        GoTo Continue
                    End If
                    Name = Trim$(StringIntersection(MyText, "/mute"))
                    Name = Trim$(StringIntersection(Name, Command(2)))
                    SendMute Name, CLng(Command(2))
                Case "/unmute"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                    
                    If UBound(Command) < 1 Then
                        AddText "Use /mute (Name)", AlertColor, True
                    End If
                    
                    SendMute Trim$(StringIntersection(MyText, "/unmute")), 0
                    ' Ban destroy
                Case "/shutdown"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue
                    
                    SendServerShutdown
                Case "/restart"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue
                    
                    SendServerRestart
                    
                Case "/destroybanlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    SendBanDestroy
                    ' Packet debug mode
                Case "/debug"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    DEBUG_MODE = (Not DEBUG_MODE)
                Case "/checkitems"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue
                    
                    SendCheckItems
                Case "/cmd"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue
                    
                    SendSpecialCommand Command
                Case Else
                    AddText "It's not a valid command!", HelpColor, True
            End Select

            'continue label where we go instead of exiting the sub
Continue:
            MyText = vbNullString
            frmMain.txtMyChat.text = vbNullString
            Exit Sub
        End If
        
        'Guild Message
        If frmMain.ChatOpts(3).value = True Then
            ChatText = Mid$(ChatText, 1, Len(ChatText))
    
            If Len(ChatText) > 0 Then
                Call GuildMsg(ChatText)
            End If
    
            MyText = vbNullString
            frmMain.txtMyChat.text = vbNullString
            Exit Sub
    End If
    
    ' Party Msg
        If frmMain.ChatOpts(2).value = True Then
            ChatText = Mid$(ChatText, 1, Len(ChatText))

            If Len(ChatText) > 0 Then
                Call SendPartyChatMsg(ChatText)
            End If

            MyText = vbNullString
            frmMain.txtMyChat.text = vbNullString
            Exit Sub
        End If

        ' Broadcast message
        If frmMain.ChatOpts(1).value = True Then
            ChatText = Mid$(ChatText, 1, Len(ChatText))

            If Len(ChatText) > 0 Then
                Call BroadcastMsg(ChatText)
            End If

            MyText = vbNullString
            frmMain.txtMyChat.text = vbNullString
            Exit Sub
        End If

        ' Say message
        If Len(ChatText) > 0 Then
            Call SayMsg(ChatText)
        End If

        MyText = vbNullString
        frmMain.txtMyChat.text = vbNullString
        Exit Sub
    End If

    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
    End If

    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) Then
        If (KeyAscii <> vbKeyBack) Then
            MyText = MyText & ChrW$(KeyAscii)
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleKeyPresses", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function JoinString(ByRef s() As String, ByVal StartIndex As Long) As String
    If StartIndex < LBound(s) Or StartIndex > UBound(s) Then Exit Function
    
    Dim Aux() As String
    ReDim Aux(UBound(s) - StartIndex)
    
    Dim j As Long
    Dim i As Long
    j = 0
    For i = StartIndex To UBound(s)
        Aux(j) = s(StartIndex)
        j = j + 1
    Next
    
    JoinString = Join(Aux, " ")
End Function

Function StringIntersection(ByVal s As String, ByVal T As String) As String
    StringIntersection = Replace(s, T, vbNullString)
End Function

