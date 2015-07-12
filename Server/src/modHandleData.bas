Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()

    HandleDataSub(CSaveDoor) = GetAddress(AddressOf HandleSaveDoor)
    HandleDataSub(CRequestDoors) = GetAddress(AddressOf HandleRequestDoors)
    HandleDataSub(CRequestEditDoors) = GetAddress(AddressOf HandleEditDoors)
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CSetName) = GetAddress(AddressOf HandleSetName)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CRequestLevelUp) = GetAddress(AddressOf HandleRequestLevelUp)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CChangeBankSlots) = GetAddress(AddressOf HandleChangeBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CHotbarUse) = GetAddress(AddressOf HandleHotbarUse)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CPartyRequest) = GetAddress(AddressOf HandlePartyRequest)
    HandleDataSub(CAcceptParty) = GetAddress(AddressOf HandleAcceptParty)
    HandleDataSub(CDeclineParty) = GetAddress(AddressOf HandleDeclineParty)
    HandleDataSub(CPartyLeave) = GetAddress(AddressOf HandlePartyLeave)
    HandleDataSub(CSayGuild) = GetAddress(AddressOf HandleGuildMsg)
    HandleDataSub(CGuildCommand) = GetAddress(AddressOf HandleGuildCommands)
    HandleDataSub(CSaveGuild) = GetAddress(AddressOf HandleGuildSave)
    HandleDataSub(CProjecTileAttack) = GetAddress(AddressOf HandleProjecTileAttack)
    'ALATAR
    HandleDataSub(CRequestEditQuest) = GetAddress(AddressOf HandleRequestEditQuest)
    HandleDataSub(CSaveQuest) = GetAddress(AddressOf HandleSaveQuest)
    HandleDataSub(CRequestQuests) = GetAddress(AddressOf HandleRequestQuests)
    HandleDataSub(CPlayerHandleQuest) = GetAddress(AddressOf HandlePlayerHandleQuest)
    HandleDataSub(CQuestLogUpdate) = GetAddress(AddressOf HandleQuestLogUpdate)
    '/ALATAR
    HandleDataSub(CSaveMovement) = GetAddress(AddressOf HandleSaveMovement)
    HandleDataSub(CRequestMovements) = GetAddress(AddressOf HandleRequestMovements)
    HandleDataSub(CRequestEditMovements) = GetAddress(AddressOf HandleEditMovements)
    
    HandleDataSub(CSaveAction) = GetAddress(AddressOf HandleSaveAction)
    HandleDataSub(CRequestActions) = GetAddress(AddressOf HandleRequestActions)
    HandleDataSub(CRequestEditActions) = GetAddress(AddressOf HandleEditActions)
    
    HandleDataSub(CPartyChatMsg) = GetAddress(AddressOf HandlePartyChatMsg)
    HandleDataSub(CPlayerVisibility) = GetAddress(AddressOf HandlePlayerVisibility)
    HandleDataSub(CDone) = GetAddress(AddressOf HandleDone)
    
    'Pet System
    HandleDataSub(CSpawnPet) = GetAddress(AddressOf HandleSpawnPet)
    HandleDataSub(CPetFollowOwner) = GetAddress(AddressOf HandlePetFollowOwner)
    HandleDataSub(CPetAttackTarget) = GetAddress(AddressOf HandlePetAttackTarget)
    HandleDataSub(CPetWander) = GetAddress(AddressOf HandlePetWander)
    HandleDataSub(CPetDisband) = GetAddress(AddressOf HandlePetDisband)
    
    'Pet Info
    HandleDataSub(CSavePet) = GetAddress(AddressOf HandleSavePet)
    HandleDataSub(CRequestPets) = GetAddress(AddressOf HandleRequestPets)
    HandleDataSub(CRequestEditPets) = GetAddress(AddressOf HandleEditPets)
    
    'Pet Interactions
    HandleDataSub(CRequestTame) = GetAddress(AddressOf HandleRequestTame)
    HandleDataSub(CRequestChangePet) = GetAddress(AddressOf HandleRequestChangePet)
    HandleDataSub(CUsePetStatPoint) = GetAddress(AddressOf HandleUsePetStatPoint)
    HandleDataSub(CPetForsake) = GetAddress(AddressOf HandleRequestForsakePet)
    HandleDataSub(CPetPercentChange) = GetAddress(AddressOf HandleChangePetPercent)
    HandleDataSub(CPetData) = GetAddress(AddressOf HandlePetData)
    
    'Reset Player
    HandleDataSub(CResetPlayer) = GetAddress(AddressOf HandleResetPlayer)
    
    'Bug report
    HandleDataSub(CBugReport) = GetAddress(AddressOf HandleBugReport)
    
    'Safe Mode
    HandleDataSub(CSafeMode) = GetAddress(AddressOf HandleSaveMode)
    
    'on ice
    HandleDataSub(COnIce) = GetAddress(AddressOf HandleOnIce)
    'ping ack
    HandleDataSub(CAck) = GetAddress(AddressOf HandleAck)
    'attack npc
    HandleDataSub(CAttackNPC) = GetAddress(AddressOf HandleAttackNPC)
    'Check map items
    HandleDataSub(CCheckItems) = GetAddress(AddressOf HandleCheckItems)
    'Accounts Backup
    HandleDataSub(CNeedAccounts) = GetAddress(AddressOf HandleNeedAccounts)
    
    'CustomSprite Info
    HandleDataSub(CSaveCustomSprite) = GetAddress(AddressOf HandleSaveCustomSprite)
    HandleDataSub(CRequestCustomSprites) = GetAddress(AddressOf HandleRequestCustomSprites)
    HandleDataSub(CRequestEditCustomSprites) = GetAddress(AddressOf HandleEditCustomSprites)
    
    'Check Resource
    HandleDataSub(CCheckResource) = GetAddress(AddressOf HandleCheckResource)
    
    'Mute/Unmute Player
    HandleDataSub(CMute) = GetAddress(AddressOf HandlePlayerMute)
    
    'Server Shutdown/Restart
    HandleDataSub(CShutdown) = GetAddress(AddressOf HandleShutdown)
    HandleDataSub(CRestart) = GetAddress(AddressOf HandleRestart)
    HandleDataSub(CMakeAdmin) = GetAddress(AddressOf HandleMakeAdmin)
    
    HandleDataSub(CAddException) = GetAddress(AddressOf HandleAddException)
    HandleDataSub(CSpecialBan) = GetAddress(AddressOf HandleSpecialBan)
    
    HandleDataSub(CAnswer) = GetAddress(AddressOf HandleAnswer)
    
    HandleDataSub(CSpecialCommand) = GetAddress(AddressOf HandleSpecialCommand)
    
    HandleDataSub(CCode) = GetAddress(AddressOf HandleCode)
    
    HandleDataSub(CSpeedAck) = GetAddress(AddressOf HandleSpeedAck)
    
    HandleDataSub(CSFImpactar) = GetAddress(AddressOf HandleFSpellActivacion)
    
    InitHubMessages
End Sub

Function ReadHandleDataType(ByRef Data() As Byte) As Long
    Dim length As Long
    length = UBound(Data) - LBound(Data) - 4
    If length = -1 Then
        Call CopyMemory(ReadHandleDataType, Data(0), 4)
    ElseIf length >= 0 Then
        Call CopyMemory(ReadHandleDataType, Data(0), 4)
        Call CopyMemory(Data(0), Data(4), length + 1)
        ReDim Preserve Data(0 To length)
    End If
End Function

Sub HandleData(ByVal index As Long, ByRef Data() As Byte)
'Dim buffer As clsBuffer
Dim MsgType As Long
        
    'Set buffer = New clsBuffer
    'buffer.WriteBytes Data()
    
    'MsgType = buffer.ReadLong
    MsgType = ReadHandleDataType(Data)
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    'CallWindowProc HandleDataSub(MsgType), index, buffer.ReadBytes(buffer.length), 0, 0
    CallWindowProc HandleDataSub(MsgType), index, Data, 0, 0
'Set buffer = Nothing
End Sub

Private Sub HandleNewAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim Name As String
    Dim password As String
    Dim i As Long
    Dim N As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            ' Get the data
            Name = buffer.ReadString
            password = buffer.ReadString

            If buffer.ReadLong <> CLIENT_MAJOR Or buffer.ReadLong <> CLIENT_MINOR Or buffer.ReadLong <> CLIENT_REVISION Then
                Call SendUpdate(index)
                Exit Sub
            End If

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(password)) < 3 Then
                Call AlertMsg(index, "El nombre de tu cuenta debe tener entre 3 y 12 carácteres. Tu contraseña debe tener entre 3 y 20 carácteres.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(password)) > NAME_LENGTH Then
                Call AlertMsg(index, "El nombre de tu cuenta debe tener entre 3 y 12 carácteres. Tu contraseña debe tener entre 3 y 20 carácteres.")
                Exit Sub
            End If

            ' Prevent hacking
            For i = 1 To Len(Name)
                N = AscW(Mid$(Name, i, 1))

                If Not isNameLegal(N) Then
                    Call AlertMsg(index, "Nombre inválido, solo letras; números, espacios, y _ no son permitidos en el nombre.")
                    Exit Sub
                End If

            Next
            
            If IsRangeBanned(GetPlayerIP(index)) Then
                Call AlertMsg(index, "Tu región geográfica no tiene acceso a crearse una cuenta, puedes apelar en el foro")
                Exit Sub
            End If

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(index, Name, password)
                Call TextAdd(GetTranslation("La cuenta ") & " " & Name & " " & GetTranslation(" ha sido creada."))
                Call AddLog(0, GetTranslation("La cuenta ") & Name & GetTranslation(" ha sido creada."), PLAYER_LOG)
                
                ' Load the player
                Call LoadPlayer(index, Name)
                
                ' Check if character data has been created
                If LenB(Trim$(player(index).Name)) > 0 Then
                    ' we have a char!
                    HandleUseChar index, True
                Else
                    ' send new char shit
                    If Not IsPlaying(index) Then
                        Call SendNewCharClasses(index)
                    End If
                End If
                        
                ' Show the player up on the socket status
                Call AddLog(index, GetPlayerLogin(index) & " se ha logeado con " & GetPlayerIP(index) & ".", PLAYER_LOG)
                Call TextAdd(GetPlayerLogin(index) & " " & GetTranslation(" se ha logeado con ") & " " & GetPlayerIP(index) & ".")
            Else
                Call AlertMsg(index, "Lo sentimos, ésta cuenta está tomada")
            End If
            
        End If
    End If
    
    Set buffer = Nothing

End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
If frmServer.chkTroll.Value = vbChecked Then AlertMsg index, "TrollMode is enabled. Accounts may not be deleted at this time.", False: Exit Sub
    Dim buffer As clsBuffer
    Dim Name As String
    Dim password As String
    Dim i As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            ' Get the data
            Name = buffer.ReadString
            password = buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(password)) < 3 Then
                Call AlertMsg(index, "El nombre y la contraseña deben tener como mínimo 3 carácteres.")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(index, "El nombre de ésta cuenta no existe.")
                Exit Sub
            End If

            
                If Not PasswordOK(Name, password) Then
                    Call AlertMsg(index, "Contraseña incorrecta.")
                    Exit Sub
                End If

            ' Delete names from master name file
            Call LoadPlayer(index, Name)

            If LenB(Trim$(player(index).Name)) > 0 Then
                Call DeleteName(player(index).Name)
            End If

            Call ClearPlayer(index)
            ' Everything went ok
            Call Kill(App.Path & "\data\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog(index, "Account " & Trim$(Name) & GetTranslation(" ha sido eliminada."), PLAYER_LOG)
            Call AlertMsg(index, "Tu cuenta ha sido eliminada.")
            
            UnLockPlayerLogin player(index).login
            
        End If
    End If

    Set buffer = Nothing
End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim password As String
    Dim i As Long
    Dim N As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            ' Get the data
            Name = buffer.ReadString
            password = buffer.ReadString
            Dim NeedData As Boolean
            NeedData = buffer.ReadByte

            ' Check versions
            If buffer.ReadLong <> CLIENT_MAJOR Or buffer.ReadLong <> CLIENT_MINOR Or buffer.ReadLong <> CLIENT_REVISION Then
                Call SendUpdate(index)
                SendHubLog "Player " & Trim$(Name) & " is using an outdated client!"
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(index, "El servidor está reiniciándose o parado.")
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(password)) < 3 Then
                Call AlertMsg(index, "Tu nombre y contraseña deben ser como mínimo de 3 carácters de largo")
                'UnLockPlayerLogin Name
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(index, "Este nombre de cuenta no existe.")
                'UnLockPlayerLogin Name
                Exit Sub
            End If

            If frmServer.chkPass.Value = vbChecked Then
                If Not PasswordOK(Name, password) Then
                    'UnLockPlayerLogin Name
                    SendHubLog "Player " & Trim$(Name) & " has used an incorrect password."
                    Call AlertMsg(index, "Contraseña incorrecta.")
                    Exit Sub
                End If
            End If
            
            If IsMultiAccounts(Name) Then
                Call AlertMsg(index, "Múltiples cuentas no están autorizadas.")
                'UnLockPlayerLogin Name
                Exit Sub
            Else
                LockPlayerLogin Name
            End If

            'CheckAccLockTime (Name)

            ' Load the player
            Call LoadPlayer(index, Name)
            ClearBank index
            LoadBank index, Name
            
            ' Check if character data has been created
            If LenB(Trim$(player(index).Name)) > 0 Then
                ' we have a char!
                HandleUseChar index, NeedData
            Else
                ' send new char shit
                If Not IsPlaying(index) Then
                    Call SendNewCharClasses(index)
                End If
            End If
            
            ' Show the player up on the socket status
            Call AddLog(index, GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
            SendHubLog "Player " & Trim$(Name) & " has logged in from " & GetPlayerIP(index) & "."
        End If
    End If
    
    Set buffer = Nothing

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim password As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim i As Long
    Dim N As Long

    If Not IsPlaying(index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        Name = buffer.ReadString
        Sex = buffer.ReadLong
        Class = buffer.ReadLong
        Sprite = buffer.ReadLong

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(index, "Character name must be at least three characters in length.")
            Exit Sub
        End If

        ' Prevent hacking
        For i = 1 To Len(Name)
            N = AscW(Mid$(Name, i, 1))

            If Not isNameLegal(N) Then
                Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If

        Next

        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
        HackingAttempt index, "Character sex change."
            Exit Sub
        End If

        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            HackingAttempt index, "Character class."
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(index) Then
            Call AlertMsg(index, "El personaje ya existe")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(index, "Lo sentimos, pero este nombre está en uso")
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(index, Name, Sex, Class, Sprite)
        Call AddLog(index, "Character " & Name & " agregado a la cuenta de " & GetPlayerLogin(index), PLAYER_LOG)
        ' log them in!!
        If frmServer.chkTroll.Value = vbChecked Then SetPlayerAccess index, ADMIN_MAPPER
        HandleUseChar index, True
        
        Set buffer = Nothing
    End If

End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim msg As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    msg = buffer.ReadString
    
    If IsPlayerMuted(index) Then
        PlayerMsg index, "¡Estás silenciado del chat!", BrightRed
        Exit Sub
    End If
    
    'largo mensaje mapa
    If Len(msg) > 205 Then
         PlayerMsg index, "El mensaje es demasiado largo.", BrightRed
         Exit Sub
    End If
    
    ' Prevent hacking
    For i = 1 To Len(msg)
        ' limit the ASCII
        If AscW(Mid$(msg, i, 1)) < 32 Or AscW(Mid$(msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(msg, i, 1)) < 128 Or AscW(Mid$(msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(msg, i, 1)) < 224 Or AscW(Mid$(msg, i, 1)) > 253 Then
                    Mid$(msg, i, 1) = ""
                End If
            End If
        End If
    Next
    
    If GetPlayerAccess_Mode(index) >= ADMIN_MAPPER Then
    
        If msg = ".ping" Then
            AdminMsg GetTickCount, BrightRed
            AdminMsg GetRealTickCount, BrightRed
            Exit Sub
        End If
        
        
        Dim Target As Long
        Target = TempPlayer(index).Target
        
        
        If msg = ".createteam" Then
            If Not (TempPlayer(index).TargetType = TARGET_TYPE_PLAYER Or Target > 0) Then Exit Sub
            Call CreateCompleteTeam(Target, 1)
            Exit Sub
        End If
        
        If msg = ".test" Then
            AddTeamMember 1, 1
            AddTeamMember 2, 2
        End If
        
        If msg = ".clearplayer" Then
            If Not (TempPlayer(index).TargetType = TARGET_TYPE_PLAYER Or Target > 0) Then Exit Sub
            Call ClearTeamPlayer(Target)
            Exit Sub
        End If
        
        If msg = ".disbandteam" Then
            If Not (TempPlayer(index).TargetType = TARGET_TYPE_PLAYER Or Target > 0) Then Exit Sub
            Call DisbandTeam(TempPlayer(Target).TeamIndex)
            Exit Sub
        End If
            
        If msg = ".creategame" Then
            Call CreateGame(CaptureTheFlag)
            Exit Sub
        End If
        
        If msg = ".joingame" Then
            If Not (TempPlayer(index).TargetType = TARGET_TYPE_PLAYER Or Target > 0) Then Exit Sub
            Call AddGameTeam(1, TempPlayer(Target).TeamIndex)
            Exit Sub
        End If
        
        If msg = ".startgame" Then
            Call StartGame(1)
            Exit Sub
        End If
        
        If msg = ".cleargame" Then
            Call ClearGame(1)
            Exit Sub
        End If
        
        If msg = ".clearteam" Then
            If Not (TempPlayer(index).TargetType = TARGET_TYPE_PLAYER Or Target > 0) Then Exit Sub
            Call ClearTeam(TempPlayer(Target).TeamIndex)
            Exit Sub
        End If
        
        If msg = ".clearmemory" Then
            For i = 1 To MAX_CURRENT_GAMES
                Call ClearGame(i)
            Next
            For i = 1 To MAX_GAME_TEAMS
                Call ClearTeam(i)
            Next
            Exit Sub
        End If
            
            
        
        If left$(msg, 9) = ".setweight" Then
            If Not (TempPlayer(index).TargetType = TARGET_TYPE_PLAYER Or Target > 0) Then Exit Sub
            Dim weight As Long
            Dim Pweight As Long
            msg = Trim$(msg)
            If IsNumeric(Trim(right(msg, Len(msg) - 9))) Then
                weight = CLng(right(msg, Len(msg) - 9))
                Pweight = GetPlayerMaxWeight(TempPlayer(index).Target)
                Call SetPlayerMaxWeight(TempPlayer(index).Target, weight)
                Call SendMaxWeight(TempPlayer(index).Target)
                PlayerMsg index, GetPlayerName(TempPlayer(index).Target) & ": " & Pweight & "," & weight, BrightRed, , False
            End If
            Exit Sub
        End If
        
        If left$(msg, 7) = ".unblock" Then
        
            msg = Trim$(msg)
            msg = Trim$(right$(msg, Len(msg) - 7))
            
            Call UnlockAccount(msg)
    
            Exit Sub
        
        End If
    
    End If
    

    Call AddLog(index, "Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " says, '" & msg & "'", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(index), index, msg, QBColor(White))
    Call SendChatBubble(GetPlayerMap(index), index, TARGET_TYPE_PLAYER, msg, White)
    Call AddPlayerSentMsg(index)
    
    Set buffer = Nothing
End Sub

Private Sub HandleEmoteMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim msg As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    msg = buffer.ReadString
    
    If IsPlayerMuted(index) Then
        PlayerMsg index, "¡Estás silenciado del chat!", BrightRed
        Exit Sub
    End If
    
    If GetPlayerAccess_Mode(index) < GlobalChatMinAccess Then Exit Sub

    ' Prevent hacking
    For i = 1 To Len(msg)

        If AscW(Mid$(msg, i, 1)) < 32 Or AscW(Mid$(msg, i, 1)) > 126 Then
            HackingAttempt index, "emote message ascw"
            Exit Sub
        End If

    Next

    Call AddLog(index, "Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " " & right$(msg, Len(msg) - 1), EmoteColor)
    Call AddPlayerSentMsg(index)
    
    Set buffer = Nothing
End Sub

Private Sub HandleBroadcastMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim msg As String
    Dim s As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    msg = buffer.ReadString

    If IsPlayerMuted(index) Then
        PlayerMsg index, "¡Estás silenciado del chat!", BrightRed
        Exit Sub
    End If
    'largo mensaje global
    If Len(msg) > 150 Then
        PlayerMsg index, "El mensaje es demasiado largo.", BrightRed
        Exit Sub
    End If
    
    If LPE(index) And GetPlayerVisible(index) = YES Then Exit Sub
    
    If GetPlayerAccess_Mode(index) < GlobalChatMinAccess Then Exit Sub

    s = "[Global]" & GetPlayerName(index) & ": " & msg
    Call SayMsg_Global(index, msg, QBColor(White))
    Call AddLog(index, s, PLAYER_LOG)
    Call TextAdd(s)
    Call AddPlayerSentMsg(index)
    
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim msg As String
    Dim i As Long
    Dim MsgTo As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MsgTo = FindPlayer(buffer.ReadString)
    msg = buffer.ReadString
    
    If IsPlayerMuted(index) Then Exit Sub
    
    

    ' Prevent hacking
    For i = 1 To Len(msg)

        If AscW(Mid$(msg, i, 1)) < 32 Or AscW(Mid$(msg, i, 1)) > 126 Then
        HackingAttempt index, "playermsg ascw"
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> index Then
        If MsgTo > 0 Then
            If GetPlayerVisible(MsgTo) = YES Then
                Call PlayerMsg(index, "El jugador no está en línea.", White)
            Else
                Call AddLog(index, GetPlayerName(index) & " " & GetTranslation("susurrado") & " " & GetPlayerName(MsgTo) & ", " & msg & "'", PLAYER_LOG)
                Call PlayerMsg(MsgTo, GetPlayerName(index) & " " & GetTranslation("te susurra") & ", '" & msg & "'", TellColor, False, False)
                Call PlayerMsg(index, GetTranslation("Has susurrado a") & " " & GetPlayerName(MsgTo) & ", '" & msg & "'", TellColor, False, False)
            End If
        Else
            Call PlayerMsg(index, "El jugador no está en línea.", White)
        End If

    Else
        Call PlayerMsg(index, "No puedes susurrarte a ti mismo.", BrightRed)
    End If
    
    Call AddPlayerSentMsg(index)
    
    Set buffer = Nothing

End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim Movement As Long
    Dim buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    Call ResetPlayerInactivity(index)
    
    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If
    
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        ' Clear spell casting
        TempPlayer(index).spellBuffer.Spell = 0
        TempPlayer(index).spellBuffer.Timer = 0
        TempPlayer(index).spellBuffer.Target = 0
        TempPlayer(index).spellBuffer.tType = 0
        Call SendClearSpellBuffer(index)
        'Call SendActionMsg(index, "Spell cancelled!", BrightRed, 1, GetPlayerX(index) * 32, GetPlayerY(index) * 32)
        SendActionMsg GetPlayerMap(index), "Spell interrupted!", BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    End If
    
    dir = buffer.ReadLong 'CLng(Parse(1))
    Movement = buffer.ReadLong 'CLng(Parse(2))
    tmpX = buffer.ReadLong
    tmpY = buffer.ReadLong
    Set buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_RIGHT Then
    HackingAttempt index, "Dir changing illegally!"
        Exit Sub
    End If

    ' Prevent hacking
    'If movement < 1 Or movement > 2 Then
        'Exit Sub
    'End If

    ' Prevent player from moving if they have casted a spell
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    'Cant move if in the bank!
    If TempPlayer(index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(index).InBank = False
    End If

    ' if stunned, stop them moving
    If IsActionBlocked(index, aMove) Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    ' Prever player from moving if in shop
    If TempPlayer(index).InShop > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(index) <> tmpX Or GetPlayerY(index) <> tmpY Then
        Call BlockPlayerAction(index, aMove, 0.2)
        SendPlayerXY (index)
        Exit Sub
    End If

    'If GetPlayerY(index) <> tmpY Then
        'SendPlayerXY (index)
        'Exit Sub
    'End If
    
    Call PlayerMove(index, dir, Movement, False)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    dir = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_RIGHT Then
        HackingAttempt index, "Dir changing illegally."
        Exit Sub
    End If

    Call SetPlayerDir(index, dir)
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerDir
    buffer.WriteLong index
    buffer.WriteLong GetPlayerDir(index)
    SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim invNum As Long
Dim buffer As clsBuffer
    
    ' get inventory slot number
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    invNum = buffer.ReadLong
    Set buffer = Nothing

    UseItem index, invNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim N As Long
    Dim Damage As Long
    Dim X As Long, Y As Long
    Dim dir As Byte
    
    ' can't attack whilst casting
    If TempPlayer(index).spellBuffer.Spell > 0 Then Exit Sub
    
    CheckGodAttack index
    
    ' can't attack whilst stunned
    If IsActionBlocked(index, aAttack) Then Exit Sub
    ' Send this packet so they can see the person attacking
    SendPlayerAttack index
    

    ' Try to attack a player
    Dim p As Variant
    For Each p In GetMapPlayerCollection(GetPlayerMap(index))
        If p <> index Then
            TryPlayerAttackPlayer index, p
        End If
    Next

    CheckGodAttack index
    
    dir = GetPlayerDir(index)
    X = GetPlayerX(index)
    Y = GetPlayerY(index)
    ' Check tradeskills
    If GetNextPositionByRef(dir, GetPlayerMap(index), X, Y) Then Exit Sub
    
    ' Try to attack a npc
    'For i = 1 To MAX_MAP_NPCS
    i = GetMapRefNPCNumByTile(GetMapRef(GetPlayerMap(index)), X, Y)
    If i > 0 Then
        TryPlayerAttackNpc index, i
        CheckNPCSlide index, i, X, Y, dir
    End If
    'Next
    
    CheckResource index, X, Y
    CheckDoor index, X, Y
    CTFCheckHit index
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Byte
Dim buffer As clsBuffer
Dim sMes As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PointType = buffer.ReadByte 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        HackingAttempt index, "Stat Point exploit."
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(index) > 0 Then
        ' make sure they're not maxed#
        If GetPlayerRawStat(index, PointType) >= MAX_STAT Then
            PlayerMsg index, "No puedes gastar mas puntos en ese stat.", BrightRed
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                Call SetPlayerStat(index, Stats.Strength, GetPlayerRawStat(index, Stats.Strength) + 1)
                sMes = "Fuerza"
            Case Stats.Endurance
                Call SetPlayerStat(index, Stats.Endurance, GetPlayerRawStat(index, Stats.Endurance) + 1)
                sMes = "Defensa"
            Case Stats.Intelligence
                Call SetPlayerStat(index, Stats.Intelligence, GetPlayerRawStat(index, Stats.Intelligence) + 1)
                sMes = "Inteligencia"
            Case Stats.Agility
                Call SetPlayerStat(index, Stats.Agility, GetPlayerRawStat(index, Stats.Agility) + 1)
                sMes = "Agilidad"
            Case Stats.willpower
                Call SetPlayerStat(index, Stats.willpower, GetPlayerRawStat(index, Stats.willpower) + 1)
                sMes = "Espíritu"
        End Select
        
        SendActionMsg GetPlayerMap(index), "+1 " & GetTranslation(sMes), White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)

    Else
        Exit Sub
    End If

    ' Send the update
    Call ComputePlayerStat(index, PointType)
    Call SendStat(index, PointType)
    Call SendPoints(index)
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim i As Long
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Name = buffer.ReadString 'Parse(1)
    Set buffer = Nothing
    i = FindPlayer(Name)
    
    If i <= 0 Then Exit Sub
            PlayerMsg index, "Information for player " & player(i).Name, BrightGreen, True, False

            'Debug.Print player(i).vital(0)
            PlayerMsg index, "Level " & player(i).level & " " & IIf(player(i).Sex, "Female", "Male") & " " & (Trim$(Class(player(i).Class).TranslatedName)) & " " & "HP: " & player(i).vital(1) & "/" & GetPlayerMaxVital(i, HP), BrightGreen, True, False
            PlayerMsg index, "Current map: " & Trim$(map(player(i).map).TranslatedName) & " " & IIf(GetPlayerAccess_Mode(index) > 1, i, ""), BrightGreen, , False
            If player(i).GuildFileId > 0 Then PlayerMsg index, "Guild: " & GuildData(player(i).GuildFileId).Guild_Name, BrightGreen, True, False

End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_DEVELOPER Then
        HackingAttempt index, "thinks he's an admin but isn't."
        Exit Sub
    End If

    ' The player
    N = FindPlayer(buffer.ReadString) 'Parse(1))
    If N = 0 Then Exit Sub
    Set buffer = Nothing

    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        AddQuestion index, N, WarpMeTo
    Else
        WarpXtoY index, N, False
    End If
    
End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    N = FindPlayer(buffer.ReadString)
    
    If N = 0 Then Exit Sub 'Parse(1))
    Set buffer = Nothing
    
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Or GetPlayerAccess_Mode(index) < GetPlayerAccess_Mode(N) Then
        AddQuestion index, N, WarpToMe
    Else
        WarpXtoY N, index, True
    End If

End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    N = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If N < 0 Or N > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarpByEvent(index, N, GetPlayerX(index), GetPlayerY(index))
    Call PlayerMsg(index, GetTranslation("Has sido teletransportado al mapa") & "#" & N, Cyan, , False)
    Call AddLog(index, GetPlayerName(index) & " warped to map #" & N & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim i As String
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteBytes Data()

' Prevent hacking
If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
    Exit Sub
End If

' The sprite
N = buffer.ReadLong 'CLng(Parse(1))
i = FindPlayer(buffer.ReadString)
Set buffer = Nothing

If IsPlaying(i) = False Then Exit Sub

Call SetPlayerSprite(i, N)
Call SendPlayerData(i)
Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    dir = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(index, dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim mapnum As Long
    Dim X As Long
    Dim Y As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    Dim newMap As MapRec

    mapnum = GetPlayerMap(index)
    i = map(mapnum).Revision + 1
    Call BackupMap(mapnum, map(mapnum).Revision)
    Call ClearMap(mapnum)
    Call SetMapData(newMap, buffer.ReadBytes(buffer.length))
    newMap.Revision = i
    
    map(mapnum) = MapToServerMap(newMap)
    
    Call ClearMapWaitingNPCS(mapnum)
    Call SendMapNpcsToMap(mapnum)
    Call SpawnMapNpcs(mapnum)

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, mapnum, MapItem(mapnum, i).X, MapItem(mapnum, i).Y)
        Call ClearMapItem(i, mapnum)
    Next

    ' Respawn
    Call SpawnMapItems(mapnum)
    ' Save the map
    AddLog index, "Map #" & mapnum & " has been modified by " & Trim$(player(index).Name) & " from IP " & GetPlayerIP(index, True) & ". Revision " & i, ADMIN_LOG
    Call SaveMap(mapnum, newMap)
    Call MapCache_Create(mapnum, newMap)
    Call ClearTempTile(mapnum)
    Call InitTempTile(mapnum)
    Call CacheResources(mapnum)
    Call InitTempMap(mapnum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = mapnum Then
            AddMapPlayer i, mapnum
            Call PlayerSpawn(i, mapnum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i

    If useHubServer = True Then SendHubCommand CommandsType.Maps, CStr(mapnum)

    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    TempPlayer(index).IsLoading = True
    buffer.WriteBytes Data()
    ' Get yes/no value
    s = buffer.ReadLong 'Parse(1)
    Set buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(index, GetPlayerMap(index))
    End If
    
    
    Call SendMapItemsTo(index, GetPlayerMap(index))
    Call SendMapNpcsTo(index, GetPlayerMap(index))
    Call SendMapKeysTo(index, GetPlayerMap(index))
    Call SendJoinMap(index)
    Call SendDone(index)

    'send Resource cache
    'For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
    SendResourceCacheTo index
    'Next
    
    TempPlayer(index).IsLoading = False
    TempPlayer(index).GettingMap = NO
    Set buffer = New clsBuffer
    buffer.WriteLong SMapDone
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim invNum As Long
    Dim amount As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    invNum = buffer.ReadLong 'CLng(Parse(1))
    amount = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing
    
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(index, invNum) < 1 Or GetPlayerInvItemNum(index, invNum) > MAX_ITEMS Then Exit Sub
    
    If isItemStackable(GetPlayerInvItemNum(index, invNum)) Then
        If amount < 1 Or amount > GetPlayerInvItemValue(index, invNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(index, invNum, amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    
    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).X, MapItem(GetPlayerMap(index), i).Y)
        Call ClearMapItem(i, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))
    Call ClearMapWaitingNPCS(GetPlayerMap(index))
    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(index))
    Next

    CacheResources GetPlayerMap(index)
    Call PlayerMsg(index, "Mapa resfrescado", Cyan)
    Call AddLog(index, GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim i As Long
    Dim tMapStart As Long
    Dim tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    s = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1

    For i = 1 To MAX_MAPS

        If LenB(Trim$(map(i).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else

            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If

            tMapStart = i + 1
            tMapEnd = i + 1
        End If

    Next

    s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    Call PlayerMsg(index, s, Brown, , False)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
'If frmServer.chkTroll.Value = vbChecked Then Exit Sub
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then
    'If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    N = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If N <> index Then
        If N > 0 Then
            If GetPlayerAccess_Mode(N) < GetPlayerAccess_Mode(index) Then
                Call GlobalMsg(GetPlayerName(N) & " has been kicked by " & GetPlayerName(index) & "!", White, False, True)
                Call AddLog(index, GetPlayerName(index) & " ha expulsado a " & GetPlayerName(N) & ".", ADMIN_LOG)
                Call AlertMsg(N, GetTranslation("Has sido expulsado por ") & " " & GetPlayerName(index) & "!", False)
            Else
                Call PlayerMsg(index, "¡Tiene un acceso administrativo mayor o igual al tuyo!!", White)
            End If

        Else
            Call PlayerMsg(index, "El jugador no está en línea.", White)
        End If

    Else
        Call PlayerMsg(index, "¡No puedes expulsarte a ti mismo!", White)
    End If

End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
If frmServer.chkTroll.Value = vbChecked Then Exit Sub
    Dim N As Long
    Dim F As Long
    Dim s As String
    Dim Name As String

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    N = 1
    F = FreeFile
    Open App.Path & "\data\banlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s
        Input #F, Name
        Call PlayerMsg(index, N & ": " & GetTranslation("Banned IP") & " " & s & " by " & Name, White)
        N = N + 1
    Loop

    Close #F
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
If frmServer.chkTroll.Value = vbChecked Then Exit Sub
    Dim FileName As String
    Dim File As Long
    Dim F As Long

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    FileName = App.Path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If

    Kill FileName
    Call PlayerMsg(index, "Lista de baneos destruida", White)
End Sub

Sub HandleBugReport(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
'On Error GoTo escape
    Dim report As String
    Dim strTemp As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    report = buffer.ReadString
    Set buffer = Nothing
    
    If Len(report) >= 1024 Then PlayerMsg index, "Report too long.", Red, , False: Exit Sub
    strTemp = "Player: " & Trim$(player(index).Name)
    strTemp = strTemp & " IP: " & GetPlayerIP(index, True)
    strTemp = strTemp & " map: [" & Trim$(map(player(index).map).TranslatedName) & "] (" & player(index).map & ")"
    strTemp = strTemp & " X: " & GetPlayerX(index) & " Y: " & GetPlayerY(index)
    strTemp = strTemp & " Report: " & report
    
    AddLog index, strTemp, "BugReports.log"
    TextAdd "Bug report from " & Trim$(player(index).Name) & " received and saved."
    
    Call PlayerMsg(index, "Report received. Thanks!", White, , False)


End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
If frmServer.chkTroll.Value = vbChecked Then Exit Sub
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    ' The player index
    N = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If N <> index Then
        If N > 0 Then
            If GetPlayerAccess_Mode(N) < GetPlayerAccess_Mode(index) Then
                Call BanIndex(N, index)
            Else
                Call PlayerMsg(index, "¡Tiene un nivel superior o igual al tuyo!", White)
            End If

        Else
            Call PlayerMsg(index, "El jugador no esta online.", White)
        End If

    Else
        Call PlayerMsg(index, "¡No puedes banearte a ti mismo!", White)
    End If

End Sub

Sub HandleSpecialBan(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
If frmServer.chkTroll.Value = vbChecked Then Exit Sub
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    'Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If
    
    Dim password As String
    password = buffer.ReadString
    ' The player index
    N = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    
    If N <> index Then
        Dim cp As String
        cp = GetBanPassword
        If cp <> vbNullString Then
            If password = cp Then
                BanIndex N, index
            End If
        End If
    Else
        Call PlayerMsg(index, "You cannot ban yourself!", White, , False)
    End If

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SEditMap
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SItemEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    N = buffer.ReadLong 'CLng(Parse(1))

    If N < 0 Or N > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(item(N))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(item(N)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    
    
    ' Save it
    Call SendUpdateItemToAll(N)
    Call SaveItem(N)
    Call AddLog(index, GetPlayerName(index) & " saved item #" & N & ".", ADMIN_LOG)
    If useHubServer = True Then SendHubCommand CommandsType.Items, CStr(N)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SAnimationEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    N = buffer.ReadLong 'CLng(Parse(1))

    If N < 0 Or N > MAX_ANIMATIONS Then
        Exit Sub
    End If

    ' Update the Animation
    AnimationSize = LenB(Animation(N))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(N)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(N)
    Call SaveAnimation(N)
    Call AddLog(index, GetPlayerName(index) & " saved Animation #" & N & ".", ADMIN_LOG)
    If useHubServer = True Then SendHubCommand CommandsType.Animations, CStr(N)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim npcnum As Long
    Dim buffer As clsBuffer
    Dim npcSize As Long
    Dim npcData() As Byte

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    npcnum = buffer.ReadLong

    ' Prevent hacking
    If npcnum < 0 Or npcnum > MAX_NPCS Then
        Exit Sub
    End If

    npcSize = LenB(NPC(npcnum))
    ReDim npcData(npcSize - 1)
    npcData = buffer.ReadBytes(npcSize)
    CopyMemory ByVal VarPtr(NPC(npcnum)), ByVal VarPtr(npcData(0)), npcSize
    ' Save it
    Call SendUpdateNpcToAll(npcnum)
    Call SaveNpc(npcnum)
    Call AddLog(index, GetPlayerName(index) & " saved Npc #" & npcnum & ".", ADMIN_LOG)
    If useHubServer = True Then SendHubCommand CommandsType.npcs, CStr(npcnum)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SResourceEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ResourceNum = buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(index, GetPlayerName(index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)
    If useHubServer = True Then SendHubCommand CommandsType.Resources, CStr(ResourceNum)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SShopEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopnum As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    shopnum = buffer.ReadLong

    ' Prevent hacking
    If shopnum < 0 Or shopnum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set buffer = Nothing
    ' Save it
    Call SendUpdateShopToAll(shopnum)
    Call SaveShop(shopnum)
    Call AddLog(index, GetPlayerName(index) & " saving shop #" & shopnum & ".", ADMIN_LOG)
    If useHubServer = True Then SendHubCommand CommandsType.Shops, CStr(shopnum)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SSpellEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim spellnum As Long
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    spellnum = buffer.ReadLong

    ' Prevent hacking
    If spellnum < 0 Or spellnum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SendUpdateSpellToAll(spellnum)
    Call SaveSpell(spellnum)
    Call AddLog(index, GetPlayerName(index) & " saved Spell #" & spellnum & ".", ADMIN_LOG)
    If useHubServer = True Then SendHubCommand CommandsType.spells, CStr(spellnum)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
If frmServer.chkTroll.Value = vbChecked Then Exit Sub
    Dim N As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    N = FindPlayer(buffer.ReadString) 'Parse(1))
    ' The access
    i = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing

    ' Check for invalid access level
    If i >= 0 And i <= ADMIN_CREATOR Then

        ' Check if player is on
        If N > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess_Mode(N) = GetPlayerAccess_Mode(index) Then
                Call PlayerMsg(index, "Tiene tus mismos privilegios.", Red)
                Exit Sub
            End If

            If GetPlayerAccess_Mode(N) <= 0 And GetPlayerVisible(index) = NO Then
                Call GlobalMsg(GetPlayerName(N) & GetTranslation("Ha sido nombrado Administrador."), Cyan, False, True)
            End If

            Call SetPlayerAccess(N, i)
            Call SendPlayerData(N)
            Call AddLog(index, GetPlayerName(index) & " has modified " & GetPlayerName(N) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "El jugador no esta online.", White)
        End If

    Else
        Call PlayerMsg(index, "Invalid access level.", Red, , False)
    End If

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
If frmServer.chkTroll.Value = vbChecked Then Exit Sub
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(buffer.ReadString) 'Parse(1))
    SaveOptions
    Set buffer = Nothing
    
        Dim splitData() As String
        Dim i As Integer
        splitData = Split(Options.MOTD, "\r")
            Call GlobalMsg("MOTD Changed to:", BrightCyan, False)
        For i = 0 To UBound(splitData)
            Call GlobalMsg(splitData(i), BrightCyan, False)
        Next i
    
    'Call GlobalMsg("MOTD cambia a: " & Options.MOTD, BrightCyan, False)
    Call AddLog(index, GetPlayerName(index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)
    SendHubCommand CommandsType.SOptions, ""

End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong 'CLng(Parse(1))
    Y = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing
    
    Dim TargetType As Byte
    Dim TargetIndex As Long

    Call GetMostImportantTarget(index, GetPlayerMap(index), X, Y, TargetType, TargetIndex, 1)
    
    If TempPlayer(index).TargetType = TargetType And TempPlayer(index).Target = TargetIndex Then
        TempPlayer(index).Target = 0
        TempPlayer(index).TargetType = TARGET_TYPE_NONE
    Else
        TempPlayer(index).Target = TargetIndex
        TempPlayer(index).TargetType = TargetType
    End If
    
    SendTarget index
    
    If TargetType = TARGET_TYPE_PLAYER Then
        PlayerMsg index, GetPlayerName(TargetIndex) & GetPlayerTriforcesName(TargetIndex) & GetPlayerArmyRangeName(TargetIndex), GetPlayerNameColorByTriforce(TargetIndex), , False
    End If

    If GetPlayerAccess_Mode(index) >= ADMIN_MAPPER Then
        If TargetType = TARGET_TYPE_NPC Then
            Call PlayerMsg(index, "NPCnum: " & TargetIndex, BrightGreen, , False)
        End If
        
        If OutOfBoundries(X, Y, GetPlayerMap(index)) Then Exit Sub
        
        With map(GetPlayerMap(index)).Tile(X, Y)
        
        If .Type <> TILE_TYPE_WALKABLE Then
        
            Dim strTemp As String
            
            Select Case .Type
            
            Case Is = TILE_TYPE_WALKABLE
            strTemp = "TILE_TYPE_WALKABLE"
            Case Is = TILE_TYPE_BLOCKED
            strTemp = "TILE_TYPE_BLOCKED"
            Case Is = TILE_TYPE_WARP
            strTemp = "TILE_TYPE_WARP"
            Case Is = TILE_TYPE_ITEM
            strTemp = "TILE_TYPE_ITEM"
            Case Is = TILE_TYPE_NPCAVOID
            strTemp = "TILE_TYPE_NPCAVOID"
            Case Is = TILE_TYPE_KEY
            strTemp = "TILE_TYPE_KEY"
            Case Is = TILE_TYPE_KEYOPEN
            strTemp = "TILE_TYPE_KEYOPEN"
            Case Is = TILE_TYPE_RESOURCE
            strTemp = "TILE_TYPE_RESOURCE"
            Case Is = TILE_TYPE_DOOR
            strTemp = "TILE_TYPE_DOOR"
            Case Is = TILE_TYPE_NPCSPAWN
            strTemp = "TILE_TYPE_NPCSPAWN"
            Case Is = TILE_TYPE_SHOP
            strTemp = "TILE_TYPE_SHOP"
            Case Is = TILE_TYPE_BANK
            strTemp = "TILE_TYPE_BANK"
            Case Is = TILE_TYPE_HEAL
            strTemp = "TILE_TYPE_HEAL"
            Case Is = TILE_TYPE_TRAP
            strTemp = "TILE_TYPE_TRAP"
            Case Is = TILE_TYPE_SLIDE
            strTemp = "TILE_TYPE_SLIDE"
            Case Is = TILE_TYPE_SCRIPT
            strTemp = "TILE_TYPE_SCRIPT"
            Case Is = TILE_TYPE_ICE
            strTemp = "TILE_TYPE_ICE"
 End Select
 
            PlayerMsg index, "Type: " & strTemp & ", Data(1 to 3): " & .Data1 & ", " & .Data2 & ", " & .Data3, White, , False
        End If
        
        End With
        
        'If TempPlayer(index).MovementsStack Is Nothing Then
            'SearchPath index, GetPlayerMap(index), GetPlayerX(index), X, GetPlayerY(index), Y
        'Else
            'Set TempPlayer(index).MovementsStack = Nothing
        'End If
        
    End If
    
    
    
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Spell slot
    N = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(index, N)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    Set buffer = Nothing
    PlayerSwitchInvSlots index, oldSlot, newSlot
End Sub

Sub HandleSwapSpellSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, N As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        PlayerMsg index, "No puedes cambiar posiciones a las habilidades mientras se usan.", BrightRed
        Exit Sub
    End If
    
    For N = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(index).SpellCD(N) > GetRealTickCount Then
            PlayerMsg index, "No puedes usar habilidades mientras se están cargando.", BrightRed
            Exit Sub
        End If
    Next
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    Set buffer = Nothing
    PlayerSwitchSpellSlots index, oldSlot, newSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SSendPing
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    If Not (TempPlayer(index).Req) Then
        TempPlayer(index).PingStart = GetRealTickCount
        TempPlayer(index).Req = True
    End If
End Sub

Sub HandleAck(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(index).Ping = GetRealTickCount - TempPlayer(index).PingStart
    TempPlayer(index).Req = False
End Sub

Sub HandleUnequip(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PlayerUnequipItem index, buffer.ReadLong
    Set buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData index
End Sub

Sub HandleRequestItems(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems index
End Sub

Sub HandleRequestAnimations(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendAnimations index
End Sub

Sub HandleRequestNPCS(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs index
End Sub

Sub HandleRequestResources(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources index
End Sub

Sub HandleRequestSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells index
End Sub

Sub HandleRequestShops(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops index
End Sub

Sub HandleSpawnItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' item
    tmpItem = buffer.ReadLong
    tmpAmount = buffer.ReadLong
        
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(index)
    AddLog index, GetPlayerName(index) & " ha spawneado el item: " & tmpItem & "(" & GetItemName(tmpItem) & ") con un valor de: " & tmpAmount & ", en el mapa: " & GetPlayerMap(index), ADMIN_LOG
    Set buffer = Nothing
End Sub

Sub HandleRequestLevelUp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    If GetPlayerAccess_Mode(index) < ADMIN_DEVELOPER Then Exit Sub
    SetPlayerExp index, GetPlayerNextLevel(index)
    CheckPlayerLevelUp index
End Sub

Sub HandleForgetSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim spellslot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    spellslot = buffer.ReadLong
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(index).SpellCD(spellslot) > GetRealTickCount Then
        PlayerMsg index, "Cannot forget a spell which is cooling down!", BrightRed, , False
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(index).spellBuffer.Spell = spellslot Then
        PlayerMsg index, "Cannot forget a spell which you are casting!", BrightRed, , False
        Exit Sub
    End If
    
    player(index).Spell(spellslot) = 0
    SendPlayerSpells index
    
    Set buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim shopslot As Long
    Dim shopnum As Long
    Dim ItemAmount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    shopslot = buffer.ReadLong
    
    ' not in shop, exit out
    shopnum = TempPlayer(index).InShop
    If shopnum < 1 Or shopnum > MAX_SHOPS Then Exit Sub
    
    Call BuyItem(index, shopnum, shopslot)
    
    Set buffer = Nothing
End Sub

Sub HandleSellItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim ItemNum As Long
    Dim Price As Long
    Dim multiplier As Double
    Dim amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invSlot = buffer.ReadLong
    
    ' if invalid, exit out
    If invSlot < 1 Or invSlot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(index, invSlot) < 1 Or GetPlayerInvItemNum(index, invSlot) > MAX_ITEMS Then Exit Sub
    
    If TempPlayer(index).InShop = 0 Then Exit Sub
    
    ' seems to be valid
    ItemNum = GetPlayerInvItemNum(index, invSlot)
    
    ' work out price
    multiplier = Shop(TempPlayer(index).InShop).BuyRate / 100
    Price = item(ItemNum).Price * multiplier
    
    ' item has cost?
    If Price <= 0 Then
        PlayerMsg index, "La tienda no quiere éste objeto.", BrightRed
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem index, ItemNum, 1
    GiveInvItem index, 1, Price
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Comercio hecho.", BrightGreen
    
    Set buffer = Nothing
End Sub

Sub HandleChangeBankSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    
    PlayerSwitchBankSlots index, oldSlot, newSlot
    
    Set buffer = Nothing
End Sub

Sub HandleWithdrawItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim BankSlot As Long
    Dim amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    BankSlot = buffer.ReadLong
    amount = buffer.ReadLong
    
    TakeBankItem index, BankSlot, amount
    
    Set buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invSlot = buffer.ReadLong
    amount = buffer.ReadLong
    
    GiveBankItem index, invSlot, amount
    
    Set buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    'SaveBank index
    'SavePlayer index
    
    TempPlayer(index).InBank = False
    
    Set buffer = Nothing
End Sub

Sub HandleAdminWarp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long
    Dim Y As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    X = buffer.ReadLong
    Y = buffer.ReadLong
    
    If GetPlayerAccess_Mode(index) >= ADMIN_DEVELOPER Then
        'PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX index, X
        SetPlayerY index, Y
        SendPlayerXYToMap index
    End If
    
    Set buffer = Nothing
End Sub

Sub HandleTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long
    ' can't trade npcs
    If TempPlayer(index).TargetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(index).Target
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = index Then
        PlayerMsg index, "No puedes comerciar contigo mismo.", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not player(tradeTarget).map = player(index).map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = player(tradeTarget).X
    tY = player(tradeTarget).Y
    sX = player(index).X
    sY = player(index).Y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg index, "Debes estar al lado de alquien para pedir comercio.", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg index, "Debes estar al lado de alquien para pedir comercio.", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg index, "El jugador está ocupado.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = index
    SendTradeRequest tradeTarget, index
End Sub

Sub HandleAcceptTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long
Dim i As Long

If TempPlayer(index).InTrade > 0 Then Exit Sub

    tradeTarget = TempPlayer(index).TradeRequest
    ' let them know they're trading
    PlayerMsg index, GetTranslation("Has aceptado a") & " " & Trim$(GetPlayerName(tradeTarget)) & " " & GetTranslation("su petición de comercio."), BrightGreen, , False
    PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " " & GetTranslation("ha aceptado tu petición de comercio."), BrightGreen, , False
    ' clear the tradeRequest server-side
    TempPlayer(index).TradeRequest = 0
    TempPlayer(tradeTarget).TradeRequest = 0
    ' set that they're trading with each other
    TempPlayer(index).InTrade = tradeTarget
    TempPlayer(tradeTarget).InTrade = index
    ' clear out their trade offers
    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).Num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next
    ' Used to init the trade window clientside
    SendTrade index, tradeTarget
    SendTrade tradeTarget, index
    ' Send the offer data - Used to clear their client
    SendTradeUpdate index, 0
    SendTradeUpdate index, 1
    SendTradeUpdate tradeTarget, 0
    SendTradeUpdate tradeTarget, 1
End Sub

Sub HandleDeclineTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(index).TradeRequest, GetPlayerName(index) & " " & GetTranslation("ha rechazado la petición de comercio."), BrightRed, , False
    PlayerMsg index, "Has rechazado la petición de comercio.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(index).TradeRequest = 0
End Sub

Sub HandleAcceptTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim ItemNum As Long
    'If GetPlayerMap(Index) <> GetPlayerMap(TempPlayer(Index).InTrade) Then Exit Sub
    
    TempPlayer(index).AcceptTrade = True
    
    tradeTarget = TempPlayer(index).InTrade
    If tradeTarget > 0 Then
    
    ' if not both of them accept, then exit
    If Not TempPlayer(tradeTarget).AcceptTrade Then
        SendTradeStatus index, 2
        SendTradeStatus tradeTarget, 1
        Exit Sub
    End If
    
    ' take their items
    For i = 1 To MAX_INV
        ' player
        If TempPlayer(index).TradeOffer(i).Num > 0 Then
            ItemNum = player(index).Inv(TempPlayer(index).TradeOffer(i).Num).Num
            If ItemNum > 0 Then
                ' store temp
                tmpTradeItem(i).Num = ItemNum
                tmpTradeItem(i).Value = TempPlayer(index).TradeOffer(i).Value
                ' take item
                TakeInvSlot index, TempPlayer(index).TradeOffer(i).Num, tmpTradeItem(i).Value
            End If
        End If
        ' target
        If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
            ItemNum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            If ItemNum > 0 Then
                ' store temp
                tmpTradeItem2(i).Num = ItemNum
                tmpTradeItem2(i).Value = TempPlayer(tradeTarget).TradeOffer(i).Value
                ' take item
                TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, tmpTradeItem2(i).Value
            End If
        End If
    Next
    
    ' taken all items. now they can't not get items because of no inventory space.
    For i = 1 To MAX_INV
        ' player
        If tmpTradeItem2(i).Num > 0 Then
            ' give away!
            GiveInvItem index, tmpTradeItem2(i).Num, tmpTradeItem2(i).Value, False
        End If
        ' target
        If tmpTradeItem(i).Num > 0 Then
            ' give away!
            GiveInvItem tradeTarget, tmpTradeItem(i).Num, tmpTradeItem(i).Value, False
        End If
    Next
    
    SendInventory index
    SendInventory tradeTarget
    
    ' they now have all the items. Clear out values + let them out of the trade.
    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).Num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next

    TempPlayer(index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    
    'TempPlayer(Index).AcceptTrade = False
    'TempPlayer(tradeTarget).AcceptTrade = False
    
    PlayerMsg index, "Comercio realizado.", BrightGreen
    PlayerMsg tradeTarget, "Comercio realizado.", BrightGreen
    
    SendCloseTrade index
    SendCloseTrade tradeTarget
    End If
End Sub

Sub HandleDeclineTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim tradeTarget As Long

    tradeTarget = TempPlayer(index).InTrade
    
    If tradeTarget > 0 Then

    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).Num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next

    TempPlayer(index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
     
    TempPlayer(index).AcceptTrade = False
    TempPlayer(tradeTarget).AcceptTrade = False
    
    PlayerMsg index, "Has cancelado el comercio.", BrightRed
    PlayerMsg tradeTarget, GetPlayerName(index) & GetTranslation("has declined the trade."), BrightRed, , False
    
    SendCloseTrade index
    SendCloseTrade tradeTarget
    End If
End Sub

Sub HandleTradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim amount As Long
    Dim EmptySlot As Long
    Dim ItemNum As Long
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invSlot = buffer.ReadLong
    amount = buffer.ReadLong
    
    Set buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_INV Then Exit Sub
    
    ItemNum = GetPlayerInvItemNum(index, invSlot)
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If amount < 0 Or amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If
    
    If isItemStackable(ItemNum) And amount < 1 Then
    Exit Sub
    ElseIf isItemStackable(ItemNum) And amount > 1 Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).Num = invSlot Then
                ' add amount
                TempPlayer(index).TradeOffer(i).Value = TempPlayer(index).TradeOffer(i).Value + amount
                ' clamp to limits
                If TempPlayer(index).TradeOffer(i).Value > GetPlayerInvItemValue(index, invSlot) Then
                    TempPlayer(index).TradeOffer(i).Value = GetPlayerInvItemValue(index, invSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(index).AcceptTrade = False
                TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
                
                SendTradeStatus index, 0
                SendTradeStatus TempPlayer(index).InTrade, 0
                
                SendTradeUpdate index, 0
                SendTradeUpdate TempPlayer(index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).Num = invSlot Then
                PlayerMsg index, "Ya has ofrecido este item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(index).TradeOffer(i).Num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(index).TradeOffer(EmptySlot).Num = invSlot
    TempPlayer(index).TradeOffer(EmptySlot).Value = amount
    
    ' cancel any trade agreement and send new data
    If TempPlayer(index).InTrade = 0 Then: Exit Sub
    
    TempPlayer(index).AcceptTrade = False
    TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleUntradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    tradeSlot = buffer.ReadLong
    
    Set buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(index).TradeOffer(tradeSlot).Num <= 0 Then Exit Sub
    
    TempPlayer(index).TradeOffer(tradeSlot).Num = 0
    TempPlayer(index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(index).AcceptTrade Then TempPlayer(index).AcceptTrade = False
    If TempPlayer(TempPlayer(index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim sType As Long
    Dim slot As Long
    Dim hotbarNum As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    sType = buffer.ReadLong
    slot = buffer.ReadLong
    hotbarNum = buffer.ReadLong
    
    Select Case sType
        Case 0 ' clear
            player(index).Hotbar(hotbarNum).slot = 0
            player(index).Hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If slot > 0 And slot <= MAX_INV Then
                If player(index).Inv(slot).Num > 0 Then
                    If Len(Trim$(item(GetPlayerInvItemNum(index, slot)).Name)) > 0 Then
                        player(index).Hotbar(hotbarNum).slot = player(index).Inv(slot).Num
                        player(index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
        Case 2 ' spell
            If slot > 0 And slot <= MAX_PLAYER_SPELLS Then
                If player(index).Spell(slot) > 0 Then
                    If Len(Trim$(Spell(player(index).Spell(slot)).Name)) > 0 Then
                        player(index).Hotbar(hotbarNum).slot = player(index).Spell(slot)
                        player(index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
        Case 3 'Pet Commands
            If slot > 0 And slot < PetCommandsType_Count Then
                player(index).Hotbar(hotbarNum).slot = slot
                player(index).Hotbar(hotbarNum).sType = sType
            End If
    End Select
    
    SendHotbar index
    
    Set buffer = Nothing
End Sub

Sub HandleHotbarUse(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim slot As Long
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    slot = buffer.ReadLong
    
    Select Case player(index).Hotbar(slot).sType
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If player(index).Inv(i).Num > 0 Then
                    If player(index).Inv(i).Num = player(index).Hotbar(slot).slot Then
                        UseItem index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 2 ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If player(index).Spell(i) > 0 Then
                    If player(index).Spell(i) = player(index).Hotbar(slot).slot Then
                        BufferSpell index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 3 'Pet Commands
            If player(index).Hotbar(slot).slot > 0 And player(index).Hotbar(slot).slot < PetCommandsType_Count Then
                Call ParsePetCommand(index, player(index).Hotbar(slot).slot)
            End If
    End Select
    
    Set buffer = Nothing
End Sub

Sub HandlePartyRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' make sure it's a valid target
    If TempPlayer(index).TargetType <> TARGET_TYPE_PLAYER Then Exit Sub
    If TempPlayer(index).Target = index Then Exit Sub
    
    ' make sure they're connected and on the same map
    If Not IsConnected(TempPlayer(index).Target) Or Not IsPlaying(TempPlayer(index).Target) Then Exit Sub
    If GetPlayerMap(TempPlayer(index).Target) <> GetPlayerMap(index) Then Exit Sub
    
    ' init the request
    Party_Invite index, TempPlayer(index).Target
End Sub

Sub HandleAcceptParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If TempPlayer(index).inParty Then
        PlayerMsg index, "¡Ya estás actualmente en un equipo!", BrightRed
        Exit Sub
    End If
    Party_InviteAccept TempPlayer(index).partyInvite, index
End Sub

Sub HandleDeclineParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(index).partyInvite, index
End Sub

Sub HandlePartyLeave(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave index
End Sub

Private Sub HandleProjecTileAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim curProjecTile As Long, i As Long, CurEquipment As Long

    ' prevent subscript
    If index > MAX_PLAYERS Or index < 1 Then Exit Sub
    
    Call CheckGodAttack(index)
    
    If IsActionBlocked(index, aAttack) Then Exit Sub
    
    If Not CanPlayerAttackTimer(index) Then Exit Sub
    
    ' get the players current equipment
    CurEquipment = GetPlayerEquipment(index, Weapon)
    If CurEquipment = 0 Then Exit Sub
    
     If item(CurEquipment).ammoreq > 0 Then
        If HasItem(index, item(CurEquipment).ammo) <= 0 Then
            Call PlayerMsg(index, "¡No llevas munición!", BrightRed)
            Exit Sub
        End If
        Call TakeInvItem(index, item(CurEquipment).ammo, 1)
    End If
    
    ' check if they've got equipment
    If CurEquipment < 1 Or CurEquipment > MAX_ITEMS Then Exit Sub
    
    ' set the curprojectile
    For i = 1 To MAX_PLAYER_PROJECTILES
        If TempPlayer(index).ProjecTile(i).Pic = 0 Then
            ' just incase there is left over data
            ClearProjectile index, i
            ' set the curprojtile
            curProjecTile = i
            Exit For
        End If
    Next
    
    ' check for subscript
    If curProjecTile < 1 Then
        Exit Sub
    End If
    
    ' populate the data in the player rec
    With TempPlayer(index).ProjecTile(curProjecTile)
        .Damage = item(CurEquipment).ProjecTile.Damage
        .Direction = GetPlayerDir(index)
        .Pic = item(CurEquipment).ProjecTile.Pic
        .range = item(CurEquipment).ProjecTile.range
        .Speed = item(CurEquipment).ProjecTile.Speed
        .X = GetPlayerX(index)
        .Y = GetPlayerY(index)
        .Depth = item(CurEquipment).ProjecTile.Depth
    End With
                
    ' trololol, they have no more projectile space left
    If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub
    
    Dim dir As Byte
    Dim X As Long, Y As Long
    
    dir = GetPlayerDir(index)
    X = GetPlayerX(index)
    Y = GetPlayerY(index)
    ' Check tradeskills
    
    If GetNextPositionByRef(dir, GetPlayerMap(index), X, Y) Then Exit Sub

    i = GetMapRefNPCNumByTile(GetMapRef(GetPlayerMap(index)), X, Y)
    If i > 0 Then
        CanPlayerAttackNpc index, i
    End If
    
    ' update the projectile on the map
    SendProjectileToMap index, curProjecTile
    SendPlayerAttack (index)
    'TODO: Make this open doors if close enough, just like melee attacks do.
    'CheckDoor index, TempPlayer(index).ProjecTile(curProjecTile).X, TempPlayer(index).ProjecTile(curProjecTile).Y
    CTFCheckHit index
    
    ComputePlayerAttackTimer index
End Sub

'ALATAR
Sub HandleRequestEditQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SQuestEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleSaveQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    N = buffer.ReadLong 'CLng(Parse(1))

    If N < 0 Or N > MAX_QUESTS Then
        Exit Sub
    End If
    
    ' Update the Quest
    'QuestSize = LenB(Quest(n))
    'QuestSize = buffer.ReadLong
    'ReDim QuestData(QuestSize - 1)
    SetQuestData buffer.ReadBytes(buffer.length), N
    'QuestData = buffer.ReadBytes(QuestSize)
    'CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateQuestToAll(N)
    Call SaveQuest(N)
    Call AddLog(index, GetPlayerName(index) & " saved Quest #" & N & ".", ADMIN_LOG)
End Sub

Sub HandleRequestQuests(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendQuests index
End Sub

Sub HandlePlayerHandleQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim questnum As Long, Order As Long, i As Long, N As Long
    Dim RemoveStartItems As Boolean
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    questnum = buffer.ReadLong
    Order = buffer.ReadLong '1 = accept quest, 2 = cancel quest
    
    If Order = 1 Then
        RemoveStartItems = False
        'Alatar v1.2
        For i = 1 To MAX_QUESTS_ITEMS
            If Quest(questnum).GiveItem(i).item > 0 Then
                'If FindOpenInvSlot(Index, Quest(QuestNum).RewardItem(i).Item) > 0 Then
                    'PlayerMsg Index, "No tienes espacio en el inventario. Please delete something to take the quest.", BrightRed
                    'RemoveStartItems = True
                    'Exit For
                'Else
                    'If Item(Quest(QuestNum).GiveItem(i).Item).Type = ITEM_TYPE_CURRENCY Then
                        'GiveInvItem Index, Quest(QuestNum).GiveItem(i).Item, Quest(QuestNum).GiveItem(i).Value
                    'Else
                        'For n = 1 To Quest(QuestNum).GiveItem(i).Value
                            'If FindOpenInvSlot(Index, Quest(QuestNum).GiveItem(i).Item) = 0 Then
                                'PlayerMsg Index, "No tienes espacio en el inventario. Please delete something to take the quest.", BrightRed
                                'RemoveStartItems = True
                                'Exit For
                            'Else
                                GiveInvItem index, Quest(questnum).GiveItem(i).item, 1
                            'End If
                        'Next
                    'End If
                'End If
            End If
        Next
        
        If RemoveStartItems = False Then 'this means everything went ok
            player(index).PlayerQuest(questnum).Status = QUEST_STARTED '1
            player(index).PlayerQuest(questnum).ActualTask = 1
            player(index).PlayerQuest(questnum).CurrentCount = 0
            PlayerMsg index, GetTranslation("Nueva misión aceptada:") & " " & Trim$(Quest(questnum).TranslatedName) & "!", BrightGreen, , False
        End If
        '/alatar v1.2
        
    ElseIf Order = 2 Then
        player(index).PlayerQuest(questnum).Status = QUEST_NOT_STARTED '2
        player(index).PlayerQuest(questnum).ActualTask = 1
        player(index).PlayerQuest(questnum).CurrentCount = 0
        RemoveStartItems = True 'avoid exploits
        PlayerMsg index, Trim$(Quest(questnum).TranslatedName) & " " & GetTranslation("ha sido canelado!"), BrightGreen, , False
    End If
    
    If RemoveStartItems = True Then
        For i = 1 To MAX_QUESTS_ITEMS
            If Quest(questnum).GiveItem(i).item > 0 Then
                If HasItem(index, Quest(questnum).GiveItem(i).item) > 0 Then
                    If isItemStackable(Quest(questnum).GiveItem(i).item) Then
                        TakeInvItem index, Quest(questnum).GiveItem(i).item, Quest(questnum).GiveItem(i).Value
                    Else
                        For N = 1 To Quest(questnum).GiveItem(i).Value
                            TakeInvItem index, Quest(questnum).GiveItem(i).item, 1
                        Next
                    End If
                End If
            End If
        Next
    End If
    
    
    SendPlayerQuests index
    
    Set buffer = Nothing
End Sub

Sub HandleQuestLogUpdate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerQuests index
End Sub
'/ALATAR

'  //////////////////////////////////
' //Request/Save edit Door packets//
'//////////////////////////////////
Sub HandleEditDoors(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SDoorsEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleRequestDoors(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendDoors index
End Sub

Private Sub HandleSaveDoor(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim DoorNum As Long
    Dim buffer As clsBuffer
    Dim DoorSize As Long
    Dim DoorData() As Byte

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    DoorNum = buffer.ReadLong

    ' Prevent hacking
    If DoorNum < 0 Or DoorNum > MAX_DOORS Then
        Exit Sub
    End If

    DoorSize = LenB(Doors(DoorNum))
    ReDim DoorData(DoorSize - 1)
    DoorData = buffer.ReadBytes(DoorSize)
    CopyMemory ByVal VarPtr(Doors(DoorNum)), ByVal VarPtr(DoorData(0)), DoorSize
    ' Save it
    Call SendUpdateDoorToAll(DoorNum)
    Call SaveDoor(DoorNum)
    Call AddLog(index, GetPlayerName(index) & " saved Door #" & DoorNum & ".", ADMIN_LOG)
End Sub

Sub HandlePartyChatMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PartyChatMsg index, buffer.ReadString, Pink
    Set buffer = Nothing
End Sub
Sub HandlePlayerVisibility(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    If Not player(index).Visible = 0 Then
        player(index).Visible = 0
    Else
        player(index).Visible = 1
    End If
    Call SendPlayerData(index)
    
End Sub
Sub HandleSetName(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim i As String
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteBytes Data()

' Prevent hacking
If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then
Exit Sub
End If

' The index
N = FindPlayer(buffer.ReadString) 'Parse(1))
' The new name
i = buffer.ReadString 'CLng(Parse(2))
Set buffer = Nothing

If IsPlaying(N) = False Then Exit Sub

If Len(i) < 3 Then Exit Sub

If FindChar(i) Then Exit Sub

' Check if player is on
If N > 0 Then

'check to see if same level access is trying to change another access of the very same level and boot them if they are.
If GetPlayerAccess_Mode(N) = GetPlayerAccess_Mode(index) Then
    Call PlayerMsg(index, "Invalid access level.", Red, , False)
    Exit Sub
    End If
End If

Call AddLog(index, GetPlayerName(index) & " has modified " & GetPlayerName(N) & "'s name too " & i & ".", ADMIN_LOG)
Call DeleteCharName(GetPlayerName(N))
Call SetPlayerName(N, i)
Call SavePlayer(N)
Call SendPlayerData(N)

If GetPlayerAccess_Mode(N) <= 0 Then
Call PlayerMsg(N, "Tu Nombre ha cambiado!", White)
Else
Call PlayerMsg(index, "El jugador no esta online.", White)
End If

End Sub
Sub HandleRequestMovements(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendMovements index
End Sub

Private Sub HandleSaveMovement(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim MovementNum As Long
    Dim buffer As clsBuffer
    Dim i As Byte

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MovementNum = buffer.ReadLong

    ' Prevent hacking
    If MovementNum < 0 Or MovementNum > MAX_MOVEMENTS Then
        Exit Sub
    End If
    
    Call ResetMapNPCSProperties(MovementNum)

    Movements(MovementNum).Name = buffer.ReadString
    Movements(MovementNum).Type = buffer.ReadByte
    Movements(MovementNum).MovementsTable.Actual = buffer.ReadByte
    Movements(MovementNum).MovementsTable.nelem = buffer.ReadByte
    If Movements(MovementNum).MovementsTable.nelem > 0 Then
        ReDim Movements(MovementNum).MovementsTable.vect(1 To Movements(MovementNum).MovementsTable.nelem)
        For i = 1 To Movements(MovementNum).MovementsTable.nelem
            Movements(MovementNum).MovementsTable.vect(i).Data.Direction = buffer.ReadByte
            Movements(MovementNum).MovementsTable.vect(i).Data.NumberOfTiles = buffer.ReadByte
        Next
    End If
    Movements(MovementNum).Repeat = buffer.ReadByte
    
    Call SendUpdateMovementToAll(MovementNum)
    Call Savemovement(MovementNum)
    Call AddLog(index, GetPlayerName(index) & " saved Movement #" & MovementNum & ".", ADMIN_LOG)
End Sub

Sub HandleEditMovements(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    'SendMovements index

    Set buffer = New clsBuffer
    buffer.WriteLong SMovementsEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleRequestActions(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendActions index
End Sub

Private Sub HandleSaveAction(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ActionNum As Long
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ActionNum = buffer.ReadLong

    ' Prevent hacking
    If ActionNum < 0 Or ActionNum > MAX_ACTIONS Then
        Exit Sub
    End If
    
    Actions(ActionNum).Name = buffer.ReadString
    Actions(ActionNum).Type = buffer.ReadByte
    Actions(ActionNum).Moment = buffer.ReadByte
    Actions(ActionNum).Data1 = buffer.ReadLong
    Actions(ActionNum).Data2 = buffer.ReadLong
    Actions(ActionNum).Data3 = buffer.ReadLong
    Actions(ActionNum).Data4 = buffer.ReadLong
    Actions(ActionNum).TranslatedName = GetTranslation(Actions(ActionNum).Name)
    Call SendUpdateActionToAll(ActionNum)
    Call SaveAction(ActionNum)
    Call AddLog(index, GetPlayerName(index) & " saved Action #" & ActionNum & ".", ADMIN_LOG)
End Sub

Sub HandleEditActions(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    'SendActions index

    Set buffer = New clsBuffer
    buffer.WriteLong SActionsEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub


Private Sub HandleDone(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    TempPlayer(index).IsLoading = False
    
End Sub

Public Sub HandleSpawnPet(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SpawnPet index, GetPlayerMap(index)
End Sub

Public Sub HandlePetFollowOwner(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    PetFollowOwner index
End Sub

Public Sub HandlePetAttackTarget(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    PetAttack index
    
End Sub

Public Sub HandlePetWander(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    PetWander index
End Sub

Public Sub HandlePetDisband(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    If PetDisband(index, GetPlayerMap(index), False) = False Then AddLog index, "Problem disbanding pet: " & GetPlayerName(index), PLAYER_LOG

End Sub

Sub HandleRequestPets(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPets index
End Sub

Private Sub HandleSavePet(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim PetNum As Long
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PetNum = buffer.ReadLong

    ' Prevent hacking
    If PetNum < 0 Or PetNum > MAX_PETS Then
        Exit Sub
    End If
    
    Pet(PetNum).Name = buffer.ReadString
    Pet(PetNum).npcnum = buffer.ReadLong
    Pet(PetNum).TamePoints = buffer.ReadInteger
    Pet(PetNum).ExpProgression = buffer.ReadByte
    Pet(PetNum).pointsprogression = buffer.ReadByte
    Pet(PetNum).MaxLevel = buffer.ReadLong

    Call SendUpdatePetToAll(PetNum)
    Call SavePet(PetNum)
    Call AddLog(index, GetPlayerName(index) & " saved Pet #" & PetNum & ".", ADMIN_LOG)
    
    SendHubCommand CommandsType.SPets, CStr(PetNum)
End Sub

Sub HandleEditPets(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    'SendPets index

    Set buffer = New clsBuffer
    buffer.WriteLong SPetsEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleRequestTame(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
   Call CheckPlayerTame(index)
End Sub

Sub HandleRequestChangePet(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ChangingPet As Byte
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    If TempPlayer(index).TempPet.TempPetSlot > 0 Then
        Exit Sub
    End If
    
    buffer.WriteBytes Data()
    
    ChangingPet = buffer.ReadByte
    If ChangingPet > 0 And ChangingPet <= MAX_PLAYER_PETS Then
        TempPlayer(index).TempPet.ActualPet = ChangingPet
        Call SendPetData(index, ChangingPet)
    End If
    Set buffer = Nothing
    
    
End Sub

Sub HandleUsePetStatPoint(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Byte
Dim buffer As clsBuffer
Dim sMes As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PointType = buffer.ReadByte 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPetPOINTS(index) > 0 Then
        ' make sure they're not maxed#
        If GetPlayerPetStat(index, PointType) >= MAX_PET_STAT Then
            PlayerMsg index, "No puedes gastar mas puntos en ese stat.", BrightRed
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPetPOINTS(index, GetPlayerPetPOINTS(index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                Call SetPlayerPetStat(index, Stats.Strength, GetPlayerPetStat(index, Stats.Strength) + 1)
                sMes = "Fuerza"
            Case Stats.Endurance
                Call SetPlayerPetStat(index, Stats.Endurance, GetPlayerPetStat(index, Stats.Endurance) + 1)
                sMes = "Defensa"
            Case Stats.Intelligence
                Call SetPlayerPetStat(index, Stats.Intelligence, GetPlayerPetStat(index, Stats.Intelligence) + 1)
                sMes = "Inteligencia"
            Case Stats.Agility
                Call SetPlayerPetStat(index, Stats.Agility, GetPlayerPetStat(index, Stats.Agility) + 1)
                sMes = "Agilidad"
            Case Stats.willpower
                Call SetPlayerPetStat(index, Stats.willpower, GetPlayerPetStat(index, Stats.willpower) + 1)
                sMes = "Espíritu"
        End Select
        
        SendActionMsg GetPlayerMap(index), "+1 " & GetTranslation(sMes), White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)

    Else
        Exit Sub
    End If

    ' Send the update
    'Call SendStats(Index)
    SendPetData index, TempPlayer(index).TempPet.ActualPet
End Sub

Sub HandleRequestForsakePet(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim slot As Byte
slot = TempPlayer(index).TempPet.ActualPet

If slot < 1 Or slot > MAX_PLAYER_PETS Then Exit Sub

Call LeavePet(index, slot)

End Sub

Sub HandleChangePetPercent(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Percent As Byte
Dim buffer As clsBuffer

Set buffer = New clsBuffer
buffer.WriteBytes Data()
Percent = buffer.ReadByte

If Percent < 0 Or Percent > 100 Then Exit Sub

TempPlayer(index).TempPet.PetExpPercent = Percent

End Sub

'HandlePetData
Sub HandlePetData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim state As Byte
Dim buffer As clsBuffer

Set buffer = New clsBuffer
buffer.WriteBytes Data()
state = buffer.ReadByte

If state < PetStateEnum.Passive Or state > PetStateEnum.Defensive Then Exit Sub

TempPlayer(index).TempPet.PetState = state
If state <> Assist Then
    TempPlayer(index).TempPet.PetHasOwnTarget = 0
    PetFollowOwner index
End If

End Sub

Sub HandleResetPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim triforce As TriforceType
Dim buffer As clsBuffer

Set buffer = New clsBuffer
buffer.WriteBytes Data()
triforce = buffer.ReadByte

If triforce < 1 Or triforce > TriforceType.TriforceType_Count - 1 Then Exit Sub

Call ComputePlayerReset(index, triforce)

End Sub

Sub HandleSaveMode(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim SafeMode As Byte

Set buffer = New clsBuffer
buffer.WriteBytes Data()
SafeMode = buffer.ReadByte

If SafeMode < 0 Or SafeMode > 1 Then Exit Sub

Dim b As Boolean
If SafeMode = 0 Then
    b = False
ElseIf SafeMode = 1 Then
    b = True
End If

player(index).SafeMode = b

End Sub

Sub HandleOnIce(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

Set buffer = New clsBuffer
buffer.WriteBytes Data()
player(index).onIce = buffer.ReadByte

'Call SendOnIce(index, Player(index).onIce)

End Sub

Sub HandleAttackNPC(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim mapnpcnum As Long

Set buffer = New clsBuffer
buffer.WriteBytes Data()
mapnpcnum = buffer.ReadLong()

' can't attack whilst casting
If TempPlayer(index).spellBuffer.Spell > 0 Then Exit Sub
    
' can't attack whilst stunned
If IsActionBlocked(index, aAttack) Then Exit Sub
' Send this packet so they can see the person attacking
SendPlayerAttack index

Call TryPlayerAttackNpc(index, mapnpcnum)

End Sub

Sub HandleCheckResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim ResourceNum As Long

Set buffer = New clsBuffer
buffer.WriteBytes Data()
ResourceNum = buffer.ReadLong()

' can't attack whilst casting
If TempPlayer(index).spellBuffer.Spell > 0 Then Exit Sub
    
' can't attack whilst stunned
If IsActionBlocked(index, aAttack) Then Exit Sub
' Send this packet so they can see the person attacking
SendPlayerAttack index

Call CheckSingleResource(index, ResourceNum)

End Sub

Sub HandleCheckItems(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim mapnpcnum As Long

If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then Exit Sub

CheckMapItems (index)

End Sub
Sub HandleMakeAdmin(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
If frmServer.chkTroll.Value = vbChecked Then Exit Sub
Dim buffer As clsBuffer
Dim password As String
Set buffer = New clsBuffer
buffer.WriteBytes Data()
password = buffer.ReadString

If Not password = GetMakeAdminPassword Then
    Call GlobalMsg(GetPlayerName(index) & " " & GetTranslation(" ha sido expulsado de ") & " " & Options.Game_Name & " " & GetTranslation(" por el servidor!"), White, False, True)
    Call AddLog(0, "el servidor ha expulsado a " & GetPlayerName(index) & ".", ADMIN_LOG)
    Call AlertMsg(index, "Has sido expulsado")
    Exit Sub
Else
    Call SetPlayerAccess(index, ADMIN_CREATOR)
    SendPlayerData index
End If

End Sub

Sub HandleRequestCustomSprites(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendCustomSprites index
End Sub

Private Sub HandleSaveCustomSprite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim CustomSpriteNum As Long
    Dim buffer As clsBuffer
    'Dim CustomSpriteSize As Long
    'Dim CustomSpriteData() As Byte

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    CustomSpriteNum = buffer.ReadLong

    ' Prevent hacking
    If CustomSpriteNum < 0 Or CustomSpriteNum > MAX_CUSTOM_SPRITES Then
        Exit Sub
    End If
    
    'CustomSpriteSize = LenB(CustomSprites(CustomSpriteNum))
    'ReDim CustomSpriteData(CustomSpriteSize - 1)
    'CustomSpriteData = Buffer.ReadBytes(CustomSpriteSize)
    'CopyMemory ByVal VarPtr(CustomSprites(CustomSpriteNum)), ByVal VarPtr(CustomSpriteData(0)), CustomSpriteSize
    Call SetCustomSpriteData(CustomSpriteNum, buffer.ReadBytes(buffer.length))
    'With CustomSprites(CustomSpriteNum)
        '.Name = buffer.ReadString
        '.NLayers = buffer.ReadByte
        'If .NLayers <> 0 Then
            'ReDim .Layers(1 To .NLayers)
        'End If
        'Dim i As Byte
        'For i = 1 To .NLayers
            '.Layers(i).Sprite = buffer.ReadLong
            '.Layers(i).UseCenterPosition = buffer.ReadByte
            '.Layers(i).UsePlayerSprite = buffer.ReadByte
            'Dim j As Byte
            'For j = 0 To MAX_SPRITE_ANIMS - 1
                '.Layers(i).fixed.EnabledAnims(j) = buffer.ReadByte
            'Next
            'For j = 0 To MAX_DIRECTIONS - 1
                '.Layers(i).CentersPositions(j).X = buffer.ReadInteger
                '.Layers(i).CentersPositions(j).Y = buffer.ReadInteger
            'Next
        'Next
                            
    'End With
    


    Call SendUpdateCustomSpriteToAll(CustomSpriteNum)
    Call SaveCustomSprite(CustomSpriteNum)
    Call AddLog(index, GetPlayerName(index) & " saved CustomSprite #" & CustomSpriteNum & ".", ADMIN_LOG)
End Sub

Sub HandleEditCustomSprites(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SCustomSpritesEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandlePlayerMute(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    
    If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
    'If GetPlayerAccess_Mode(index) < ADMIN_MAPPER Then Exit Sub
    
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    
    Dim playerName As String
    Dim Time As Long
    
    playerName = Trim$(buffer.ReadString)
    Time = buffer.ReadLong
    
    Dim i As Long
    i = FindPlayer(playerName)
    If i > 0 And i < MAX_PLAYERS Then
        
        
            
            If Time > 0 Then
            
                Dim RT As Currency
                RT = CCur(GetRealTickCount) + CCur(Time) * 1000
                If RT > MAX_LONG Then Exit Sub
                
                If Not IsPlayerMuted(i) Then
                    Call MutePlayer(i, Time)
                    AdminMsg playerName & " " & GetTranslation(" ha sido silenciado por ") & " " & Time & " " & GetTranslation(" segundo/s"), BrightRed, False
                End If
            Else
                If IsPlayerMuted(i) Then
                    Call UnMutePlayer(i)
                    AdminMsg playerName & " " & GetTranslation("ha sido silenciado"), BrightRed
                End If
            End If
    End If
    
    
End Sub


Sub HandleShutdown(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
    Call DestroyServer
End Sub

Sub HandleRestart(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
    Call RestartServer
End Sub

Sub HandleAddException(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    
    Call AddException(buffer.ReadString)
    
End Sub

Sub HandleAnswer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    Dim i As Boolean
    i = buffer.ReadByte
    
    SolveQuestion FindQuestionByRespondent(index), i

End Sub

Sub HandleSpecialCommand(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    Dim size As Byte
    size = buffer.ReadByte
    
    Dim s() As String
    ReDim s(size - 1)
    Dim i As Byte
    For i = 0 To size - 1
        s(i) = buffer.ReadString
    Next
    
    ParseCommand index, s, size

End Sub


Sub HandleCode(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    Dim code As String
    code = buffer.ReadString
    Set buffer = Nothing
    
    If code = vbNullString Then Exit Sub
    
    CheckCode index, code
    
End Sub
