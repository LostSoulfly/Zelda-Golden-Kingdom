Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GetAddress = FunAddr
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetAddress", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub InitMessages()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    HandleDataSub(SSpeechWindow) = GetAddress(AddressOf HandleSpeechWindow)
    HandleDataSub(SDoorsEditor) = GetAddress(AddressOf HandleDoorsEditor)
    HandleDataSub(SUpdateDoors) = GetAddress(AddressOf HandleUpdateDoors)
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SFullMsg) = GetAddress(AddressOf HandleFullMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerStat) = GetAddress(AddressOf HandlePlayerStat)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SUpdateItems) = GetAddress(AddressOf HandleUpdateItems)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SDoorAnimation) = GetAddress(AddressOf HandleDoorAnimation)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    HandleDataSub(SSendGuild) = GetAddress(AddressOf HandleSendGuild)
    HandleDataSub(SAdminGuild) = GetAddress(AddressOf HandleAdminGuild)
    HandleDataSub(SQuestEditor) = GetAddress(AddressOf HandleQuestEditor)
    HandleDataSub(SUpdateQuest) = GetAddress(AddressOf HandleUpdateQuest)
    HandleDataSub(SPlayerQuest) = GetAddress(AddressOf HandlePlayerQuest)
    HandleDataSub(SQuestMessage) = GetAddress(AddressOf HandleQuestMessage)
    HandleDataSub(SHandleProjectile) = GetAddress(AddressOf HandleProjectile)
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    HandleDataSub(SLoad) = GetAddress(AddressOf HandleLoading)
    HandleDataSub(SDone) = GetAddress(AddressOf HandleLoading)
    HandleDataSub(SSendWeather) = GetAddress(AddressOf HandleSendWeather)
    HandleDataSub(SMovementsEditor) = GetAddress(AddressOf HandleMovementsEditor)
    HandleDataSub(SUpdateMovements) = GetAddress(AddressOf HandleUpdatemovements)
    HandleDataSub(SActionsEditor) = GetAddress(AddressOf HandleActionsEditor)
    HandleDataSub(SUpdateActions) = GetAddress(AddressOf HandleUpdateActions)
    HandleDataSub(SNPCCache) = GetAddress(AddressOf HandleNPCCache)
    HandleDataSub(SPetsEditor) = GetAddress(AddressOf HandlePetsEditor)
    HandleDataSub(SUpdatePets) = GetAddress(AddressOf HandleUpdatePets)
    HandleDataSub(SPetData) = GetAddress(AddressOf HandlePetData)
    HandleDataSub(SOpenTriforce) = GetAddress(AddressOf HandleOpenTriforce)
    HandleDataSub(SOnIce) = GetAddress(AddressOf HandleOnIce)
    HandleDataSub(SIceDir) = GetAddress(AddressOf HandleIceDir)
    HandleDataSub(SBags) = GetAddress(AddressOf HandleBags)
    HandleDataSub(SPoints) = GetAddress(AddressOf HandlePoints)
    HandleDataSub(SLevel) = GetAddress(AddressOf HandleLevel)
    HandleDataSub(SJustice) = GetAddress(AddressOf HandleJustice)
    HandleDataSub(SPlayerAttack) = GetAddress(AddressOf HandlePlayerAttack)
    HandleDataSub(SMapSingularNpcData) = GetAddress(AddressOf HandleMapSingularNpcData)
    HandleDataSub(SAccounts) = GetAddress(AddressOf HandleAccounts)
    HandleDataSub(SCustomSpritesEditor) = GetAddress(AddressOf HandleCustomSpritesEditor)
    HandleDataSub(SUpdateCustomSprites) = GetAddress(AddressOf HandleUpdateCustomSprites)
    HandleDataSub(SPlayerSprite) = GetAddress(AddressOf HandlePlayerSprite)
    HandleDataSub(SSingleResourceCache) = GetAddress(AddressOf HandleSingleResourceCache)
    HandleDataSub(SGuildData) = GetAddress(AddressOf HandleGuildData)
    HandleDataSub(SMaxWeight) = GetAddress(AddressOf HandleMaxWeight)
    HandleDataSub(SMapSingularItemData) = GetAddress(AddressOf HandleMapSingularItemData)
    HandleDataSub(SBanks) = GetAddress(AddressOf HandleBanks)
    HandleDataSub(SGuilds) = GetAddress(AddressOf HandleGuilds)
    HandleDataSub(SQuestion) = GetAddress(AddressOf HandleQuestion)
    HandleDataSub(SKillPoints) = GetAddress(AddressOf HandleKillPoints)
    HandleDataSub(SBonusPoints) = GetAddress(AddressOf HandleBonusPoints)
    HandleDataSub(SUpdateNPCS) = GetAddress(AddressOf HandleUpdateNPCS)
    HandleDataSub(SSpeedReq) = GetAddress(AddressOf HandleSpeedReq)
    HandleDataSub(SPlayerSpeed) = GetAddress(AddressOf HandlePlayerSpeed)
    HandleDataSub(SRunningSprites) = GetAddress(AddressOf HandleRunningSprites)
    HandleDataSub(SPlayerState) = GetAddress(AddressOf HandlePlayerState)
    HandleDataSub(SUpdate) = GetAddress(AddressOf HandleUpdate)
    HandleDataSub(SStaminaInfo) = GetAddress(AddressOf HandleStaminaInfo)
    HandleDataSub(SCharList) = GetAddress(AddressOf HandleCharList)
    HandleDataSub(SSaveFiles) = GetAddress(AddressOf HandleSaveFiles)
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.length), 0, 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAlertMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim msg As String
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    frmLoad.Visible = False
    frmMenu.Visible = True
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    frmMenu.picMain.Visible = True
    
    msg = Buffer.ReadString 'Parse(1)
    
    Set Buffer = Nothing
    Call MsgBox(msg, vbOKOnly, Options.Game_Name)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleFullMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim msg As String
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    frmLoad.Visible = False
    frmMenu.Visible = True
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    frmMenu.picMain.Visible = True
    
    msg = Buffer.ReadString 'Parse(1)
    
    Set Buffer = Nothing
    Call MsgBox(msg, vbOKOnly, Options.Game_Name)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleLoginOk(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' save options
    Options.SavePass = frmMenu.chkPass.value
    Options.Username = Trim$(frmMenu.txtLUser.text)

    If frmMenu.chkPass.value = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = Trim$(frmMenu.txtLPass.text)
    End If
    
    SaveOptions
    
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong
    
    ' player high index
    Player_HighIndex = Buffer.ReadLong
    
    Set Buffer = Nothing
    frmLoad.Visible = True
    Call SetStatus("Login OK, receiving data...")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLoginOk", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleNewCharClasses(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim i As Long
Dim Z As Long, X As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    N = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    N = N + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString
            .TranslatedName = GetTranslation(Class(i).Name)
            .vital(Vitals.HP) = Buffer.ReadLong
            .vital(Vitals.mp) = Buffer.ReadLong
            
            ' get array size
            Z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .MaleSprite(X) = Buffer.ReadLong
            Next
            
            ' get array size
            Z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .FemaleSprite(X) = Buffer.ReadLong
            Next
            
            For X = 1 To Stats.Stat_Count - 1
                .stat(X) = Buffer.ReadLong
            Next
        End With

        N = N + 10
    Next

    Set Buffer = Nothing
    
    ' Used for if the player is creating a new character
    frmMenu.Visible = True
    frmMenu.picCharacter.Visible = True
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picRegister.Visible = False
    frmLoad.Visible = False
    frmMenu.cmbClass.Clear
    For i = 1 To Max_Classes
        frmMenu.cmbClass.AddItem Trim$(Class(i).TranslatedName)
    Next

    frmMenu.cmbClass.ListIndex = 0
    N = frmMenu.cmbClass.ListIndex + 1
    
    newCharSprite = 0
    NewCharacterBltSprite
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNewCharClasses", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleClassesData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim i As Long
Dim Z As Long, X As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    N = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    N = N + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString 'Trim$(Parse(n))
            .TranslatedName = GetTranslation(Class(i).Name)
            .vital(Vitals.HP) = Buffer.ReadLong 'CLng(Parse(n + 1))
            .vital(Vitals.mp) = Buffer.ReadLong 'CLng(Parse(n + 2))
            
            ' get array size
            Z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .MaleSprite(X) = Buffer.ReadLong
            Next
            
            ' get array size
            Z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .FemaleSprite(X) = Buffer.ReadLong
            Next
                            
            .Face = Buffer.ReadLong
            
            For X = 1 To Stats.Stat_Count - 1
                .stat(X) = Buffer.ReadLong
            Next
        End With

        N = N + 10
    Next

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClassesData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleInGame(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InGame = True
    IHaveData = True
    
    Call GameInit
    Call GameLoop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInv(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    N = 1

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, Buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, Buffer.ReadLong)
        N = N + 2
    Next
    
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    DisplayWeightPercent

    Set Buffer = Nothing
    BltInventory
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInv", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInvUpdate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    N = Buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, N, Buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, N, Buffer.ReadLong) 'CLng(Parse(3)))
    ' changes, clear drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    DisplayWeightPercent
    
    BltInventory
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInvUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerWornEq(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Shield)
    
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    DisplayWeightPercent
    
    BltInventory
    BltEquipment
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapWornEq(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim playerNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    playerNum = Buffer.ReadLong
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Shield)
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerHp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.HP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, Buffer.ReadLong)

    If GetPlayerMaxVital(MyIndex, Vitals.HP) > 0 Then
        'frmMain.lblHP.Caption = Int(GetPlayerVital(MyIndex, Vitals.HP) / GetPlayerMaxVital(MyIndex, Vitals.HP) * 100) & "%"
        frmMain.lblHP(1).Caption = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
        'hp bar
        frmMain.imgHPBar.Width = ((GetPlayerVital(MyIndex, Vitals.HP) / HPBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / HPBar_Width)) * HPBar_Width
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerHP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.mp) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.mp, Buffer.ReadLong)

    If GetPlayerMaxVital(MyIndex, Vitals.mp) > 0 Then
        'frmMain.lblMP.Caption = Int(GetPlayerVital(MyIndex, Vitals.MP) / GetPlayerMaxVital(MyIndex, Vitals.MP) * 100) & "%"
        frmMain.lblMP.Caption = GetPlayerVital(MyIndex, Vitals.mp) & "/" & GetPlayerMaxVital(MyIndex, Vitals.mp)
        'mp bar
        frmMain.imgMPBar.Width = ((GetPlayerVital(MyIndex, Vitals.mp) / SPRBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.mp) / SPRBar_Width)) * SPRBar_Width
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerStats(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat MyIndex, i, Buffer.ReadLong
        frmMain.lblCharStat(i).Caption = GetPlayerStat(MyIndex, i)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerStat(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Dim stat As Stats
    stat = Buffer.ReadByte
    SetPlayerStat MyIndex, stat, Buffer.ReadInteger
    frmMain.lblCharStat(stat).Caption = GetPlayerStat(MyIndex, stat)

    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerExp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim TNL As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SetPlayerExp MyIndex, Buffer.ReadLong
    TNL = Buffer.ReadLong
    frmMain.lblEXP.Caption = GetTranslation(("Experiencia:")) & GetPlayerExp(MyIndex) & "/" & TNL
    ' mp bar
    frmMain.imgEXPBar.Width = ((GetPlayerExp(MyIndex) / EXPBar_Width) / (TNL / EXPBar_Width)) * EXPBar_Width
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerExp", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long, X As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Call SetPlayerName(i, Buffer.ReadString)
    Call SetPlayerLevel(i, Buffer.ReadLong)
    Call SetPlayerPOINTS(i, Buffer.ReadLong)
    Call SetPlayerSprite(i, Buffer.ReadLong)
    Call SetPlayerMap(i, Buffer.ReadLong)
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Call SetPlayerDir(i, Buffer.ReadLong)
    Call SetPlayerAccess(i, Buffer.ReadLong)
    Call SetPlayerPK(i, Buffer.ReadLong)
    Call SetPlayerClass(i, Buffer.ReadLong)
    Call SetPlayerVisible(i, Buffer.ReadLong)
    'Kill Counter
    Player(i).Kill = Buffer.ReadLong
    Player(i).Dead = Buffer.ReadLong
    Player(i).NpcKill = Buffer.ReadLong
    Player(i).NpcDead = Buffer.ReadLong
    Player(i).EnviroDead = Buffer.ReadLong
    
    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, X, Buffer.ReadLong
    Next
    
    If Buffer.ReadByte = 1 Then
        Player(i).GuildName = Buffer.ReadString
        Player(i).GuildMemberId = Buffer.ReadLong
    Else
        Player(i).GuildName = vbNullString
        Player(i).GuildMemberId = 0
    End If
    
    'Triforces
    For X = 1 To TriforceType.TriforceType_Count - 1
        Player(i).triforce(X) = Buffer.ReadByte
    Next
    
    'Ice System
    Player(i).onIce = Buffer.ReadByte
    Player(i).IceDir = Buffer.ReadByte
    
    'Rupee system
    Player(i).RupeeBags = Buffer.ReadByte
    
    'Custom Sprite
    Player(i).CustomSprite = Buffer.ReadByte
    
    Call SetPlayerSpeed(i, MOVING_WALKING, Buffer.ReadLong)
    Call SetPlayerSpeed(i, MOVING_RUNNING, Buffer.ReadLong)
    
    Player(i).State = Buffer.ReadByte
    

    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
        
        ' Set the character windows
        frmMain.lblCharName = GetPlayerName(MyIndex) & " - Level " & GetPlayerLevel(MyIndex)
        
        For X = 1 To Stats.Stat_Count - 1
            frmMain.lblCharStat(X).Caption = GetPlayerStat(MyIndex, X)
        Next
        
        ' Set training label visiblity depending on points
        frmMain.lblPoints.Caption = GetPlayerPOINTS(MyIndex)
        If GetPlayerPOINTS(MyIndex) > 0 Then
            For X = 1 To Stats.Stat_Count - 1
                If GetPlayerStat(index, X) < 255 Then
                    frmMain.lblTrainStat(X).Visible = True
                Else
                    frmMain.lblTrainStat(X).Visible = False
                End If
            Next
            
        Else
            For X = 1 To Stats.Stat_Count - 1
                frmMain.lblTrainStat(X).Visible = False
            Next
        End If
        
        'BltFace
    End If

    ' Make sure they aren't walking
    Player(i).Moving = 0
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    
    Player(i).MovementSprite = False
    Set Player(i).LagMovements = New clsQueue
    Set Player(i).LagDirections = New clsQueue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim X As Long
Dim Y As Long
Dim dir As Long
Dim N As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    'X = Buffer.ReadLong
    'Y = Buffer.ReadLong
    dir = Buffer.ReadByte
    N = Buffer.ReadLong
    
    If Not Player(i).LagDirections Is Nothing Then
        Player(i).LagDirections.Push dir
        Player(i).LagMovements.Push N
    End If
    
    If i = MyIndex Then
        Player(i).Started = True
    End If
    
    'PlayerMove i, Player(i).LagDirections.Front, Player(i).LagMovements.Front
    'Call SetPlayerX(i, X)
    'Call SetPlayerY(i, Y)
    'X = Player(i).X
    'Y = Player(i).Y
    'If GetNextPositionByRef(dir, X, Y) Then Exit Sub
    'SetPlayerX i, X
    'SetPlayerY i, Y
        
    'Call SetPlayerDir(i, dir)
    'Player(i).XOffset = 0
    'Player(i).YOffset = 0
    'Player(i).Moving = n

    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim mapnpcnum As Long
Dim X As Long
Dim Y As Long
Dim dir As Long
Dim movement As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    mapnpcnum = Buffer.ReadLong
    'X = Buffer.ReadLong
    'Y = Buffer.ReadLong
    'Dir = Buffer.ReadLong
    X = MapNpc(mapnpcnum).X
    Y = MapNpc(mapnpcnum).Y
    dir = Buffer.ReadByte
    If GetNextPositionByRef(dir, X, Y) Then Exit Sub

    movement = Buffer.ReadLong

    With MapNpc(mapnpcnum)
        .X = X
        .Y = Y
        .dir = dir
        '.XOffset = 0
        '.YOffset = 0
        .Moving = movement

        Select Case .dir
            Case DIR_UP
                .YOffset = PIC_Y
            Case DIR_DOWN
                .YOffset = PIC_Y * -1
            Case DIR_LEFT
                .XOffset = PIC_X
            Case DIR_RIGHT
                .XOffset = PIC_X * -1
        End Select

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim dir As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    dir = Buffer.ReadLong
    
    If index = 0 Then Exit Sub
    
    'If Player(i).Moving > 0 Then AddText "Moving but rec'd playerDir for " & i, White
    'If Not (Player(index).LagDirections Is Nothing) Then AddText "hasLag but rec'd playerDir for " & index, White
    'If Not (Player(index).LagDirections Is Nothing) Then Call SetPlayerDir(i, dir)
    
    If Not (Player(i).LagDirections Is Nothing) Then
        If Player(i).Moving = 0 Then
            Call SetPlayerDir(i, dir)
        'Else
            'AddText "hasLag but rec'd playerDir for " & i, White
            'Player(i).LagDirections.Push dir
            'Player(i).LagMovements.Push 0
        End If
    End If

    'With Player(i)
        '.XOffset = 0
        '.YOffset = 0
        '.Moving = 0
    'End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim dir As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    dir = Buffer.ReadLong

    With MapNpc(i)
        .dir = dir
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXY(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim dir As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadByte
    Y = Buffer.ReadByte
    dir = Buffer.ReadByte
    Call SetPlayerX(MyIndex, X)
    Call SetPlayerY(MyIndex, Y)
    Call SetPlayerDir(MyIndex, dir)
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).XOffset = 0
    Player(MyIndex).YOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXY", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXYMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim dir As Long
Dim Buffer As clsBuffer
Dim thePlayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    thePlayer = Buffer.ReadLong
    X = Buffer.ReadByte
    Y = Buffer.ReadByte
    dir = Buffer.ReadByte
    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
    Call SetPlayerDir(thePlayer, dir)
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).XOffset = 0
    Player(thePlayer).YOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXYMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = GetTickCount
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = GetTickCount
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCheckForMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim i As Long
Dim NeedMap As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearTempTile
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap
    
    ' clear the blood
    For i = 1 To MAX_BYTE
        Blood(i).X = 0
        Blood(i).Y = 0
        Blood(i).sprite = 0
        Blood(i).Timer = 0
    Next
    
    ' Get map num
    X = Buffer.ReadLong
    ' Get revision
    Y = Buffer.ReadLong

    If FileExist(MAP_PATH & "map" & X & MAP_EXT, False) Then
        Call LoadMap(X)
        ' Check to see if the revisions match
        NeedMap = 1

        If map.Revision = Y Then
            ' We do so we dont need the map
            'Call SendData(CNeedMap & SEP_CHAR & "n" & END_CHAR)
            NeedMap = 0
        End If

    Else
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong NeedMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCheckForMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim X As Long
Dim Y As Long
Dim i As Long
Dim Buffer As clsBuffer
Dim mapnum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()

    mapnum = Buffer.ReadLong
    SetMapData map, Decompress(Buffer.ReadBytes(Buffer.length))

    ClearTempTile
    
    Set Buffer = Nothing
    
    ' Save the map
    Call SaveMap(mapnum)

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapItemData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Dim ItemHighIndex As Long
    ItemHighIndex = Buffer.ReadLong

    For i = 1 To ItemHighIndex
        With MapItem(i)
            .PlayerName = Buffer.ReadString
            .num = Buffer.ReadLong
            .value = Buffer.ReadLong
            .X = Buffer.ReadByte
            .Y = Buffer.ReadByte
        End With
    Next
    
    For i = ItemHighIndex + 1 To MAX_MAP_ITEMS
        With MapItem(i)
            .PlayerName = vbNullString
            .num = 0
            .value = 0
            .X = 0
            .Y = 0
        End With
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapItemData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapSingularItemData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Dim ItemIndex As Long
    ItemIndex = Buffer.ReadLong

    With MapItem(ItemIndex)
        .PlayerName = Buffer.ReadString
        .num = Buffer.ReadLong
        .value = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
    End With
    

End Sub



Private Sub HandleMapNpcData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Npc_HighIndex = Buffer.ReadLong

    For i = 1 To Npc_HighIndex
        With MapNpc(i)
            .num = Buffer.ReadLong
            map.NPC(i) = .num
            .X = Buffer.ReadByte
            .Y = Buffer.ReadByte
            .dir = Buffer.ReadByte
            .vital(HP) = Buffer.ReadLong
            .petData.Owner = Buffer.ReadLong
        End With
    Next
    
    For i = Npc_HighIndex + 1 To MAX_MAP_NPCS
        With MapNpc(i)
            .num = 0
            map.NPC(i) = 0
            .X = 0
            .Y = 0
            .dir = 0
            .vital(HP) = 0
            .petData.Owner = 0
        End With
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapSingularNpcData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    Dim i As Long
    i = Buffer.ReadLong 'Get the mapnpcnum
    
    With MapNpc(i)
        .num = Buffer.ReadLong
        map.NPC(i) = .num
        .X = Buffer.ReadByte
        .Y = Buffer.ReadByte
        .dir = Buffer.ReadByte
        .vital(HP) = Buffer.ReadLong
        .petData.Owner = Buffer.ReadLong
    End With
    
    If i > Npc_HighIndex Then
        Npc_HighIndex = i
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapDone()
Dim i As Long
Dim MusicFile As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' clear the action msgs
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i
    Action_HighIndex = 1
    
    ' load tilesets we need
    LoadTilesets
            
    MusicFile = Trim$(map.Music)
    If Not MusicFile = "None." Then
        PlayMidi MusicFile
    Else
        StopMidi
    End If
    
    ' re-position the map name
    Call UpdateDrawMapName
    
    ' get the npc high index
    For i = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(i).num > 0 Then
            Npc_HighIndex = i
            Exit For
        End If
    Next
    ' make sure we're not overflowing
    If Npc_HighIndex > MAX_MAP_NPCS Then Npc_HighIndex = MAX_MAP_NPCS

    GettingMap = False
    CanMoveNow = True
    
    TempPlayer(MyIndex).IsLoading = False
    Call SendDone

    frmMain.picLoad.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapDone", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBroadcastMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(msg, Color)
    

    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBroadcastMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleGlobalMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    Color = Buffer.ReadLong
    'Call AddText(msg, Color)
    
    Dim Chatmsg  As ChatMsgRec
    
    Chatmsg.colour = QBColor(Color)
    Chatmsg.header = msg
    Chatmsg.saycolour = QBColor(Color)
    Chatmsg.text = vbNullString
    Chatmsg.ArrivedAt = GetTickCount
    
    Dim chatroom As Byte
    chatroom = SystemChat
    
    Call ListPush(ChatRooms(chatroom), Chatmsg)
    
    If CanMsgBeDisplayed(chatroom) Then
        Call AddText(msg, Color)
    End If
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleGlobalMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim Color As Byte
Dim IsSystem As Boolean

    ' If debug mode, handle error then exit out

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    Color = Buffer.ReadLong
    IsSystem = Buffer.ReadByte
    
    Dim Chatmsg  As ChatMsgRec
    
    Chatmsg.colour = QBColor(Color)
    Chatmsg.header = msg
    Chatmsg.saycolour = QBColor(Color)
    Chatmsg.text = vbNullString
    Chatmsg.ArrivedAt = GetTickCount
    
    Dim chatroom As Byte
    If IsSystem Then
        chatroom = SystemChat
    Else
        chatroom = WhisperChat
    End If
    
    Call ListPush(ChatRooms(chatroom), Chatmsg)
    
    If CanMsgBeDisplayed(chatroom) Then
        Call AddText(msg, Color)
    End If

End Sub

Private Sub HandleMapMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(msg, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAdminMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(msg, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAdminMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    N = Buffer.ReadLong

    With MapItem(N)
        .PlayerName = Buffer.ReadString
        .num = Buffer.ReadLong
        .value = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleItemEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).TranslatedName)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleItemEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimationEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).TranslatedName)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimationEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    N = Buffer.ReadLong
    Dim UseFulData As Boolean
    UseFulData = Buffer.ReadByte
    If UseFulData Then
        SetItemUsefulData N, UnCompressData(Buffer.ReadBytes(Buffer.length), 2)
    Else
        ItemSize = LenB(Item(N))
        ReDim ItemData(ItemSize - 1)
        ItemData = Buffer.ReadBytes(ItemSize)
        CopyMemory ByVal VarPtr(Item(N)), ByVal VarPtr(ItemData(0)), ItemSize
    End If
    Set Buffer = Nothing
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    BltInventory
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub HandleUpdateItems(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    'n = buffer.ReadLong
    Dim UseFulData As Boolean
    UseFulData = Buffer.ReadByte
    
    Dim UnCompressedBuffer As clsBuffer
    Set UnCompressedBuffer = New clsBuffer
    
    
    UnCompressedBuffer.WriteBytes UnCompressData(Buffer.ReadBytes(Buffer.length), 3)
    
    Dim i As Long
    If UseFulData Then
        
        For i = 1 To MAX_ITEMS
            If UnCompressedBuffer.ReadByte Then
                Dim UnCompressedData() As Byte
                SetItemUsefulData i, UnCompressedBuffer.ReadBytes(UnCompressedBuffer.ReadLong)
            End If
        Next
    Else
        For i = 1 To MAX_ITEMS
            ' Update the item
            ItemSize = LenB(Item(i))
            ReDim ItemData(ItemSize - 1)
            ItemData = UnCompressedBuffer.ReadBytes(ItemSize)
            CopyMemory ByVal VarPtr(Item(i)), ByVal VarPtr(ItemData(0)), ItemSize
        Next
    End If
    Set Buffer = Nothing
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    BltInventory
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateNPCS(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim Buffer As clsBuffer
Dim npcSize As Long
Dim npcData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    'n = buffer.ReadLong
    Dim UseFulData As Boolean
    UseFulData = Buffer.ReadByte
    
    Dim UnCompressedBuffer As clsBuffer
    Set UnCompressedBuffer = New clsBuffer
    
    
    UnCompressedBuffer.WriteBytes UnCompressData(Buffer.ReadBytes(Buffer.length), 3)
    
    Dim i As Long
    If UseFulData Then
        
        For i = 1 To MAX_NPCS
            If UnCompressedBuffer.ReadByte Then
                Dim UnCompressedData() As Byte
                SetNPCUsefulData i, UnCompressedBuffer.ReadBytes(UnCompressedBuffer.ReadLong)
            End If
        Next
    Else
        For i = 1 To MAX_NPCS
            ' Update the npc
            npcSize = LenB(NPC(i))
            ReDim npcData(npcSize - 1)
            npcData = UnCompressedBuffer.ReadBytes(npcSize)
            CopyMemory ByVal VarPtr(NPC(i)), ByVal VarPtr(npcData(0)), npcSize
        Next
    End If
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdatenpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub HandleUpdateAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim Buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    N = Buffer.ReadLong
    ' Update the Animation
    Dim UseFulData As Boolean
    UseFulData = Buffer.ReadByte
    If UseFulData Then
        SetAnimationUseFulData N, Buffer.ReadBytes(Buffer.length)
    Else
        AnimationSize = LenB(Animation(N))
        ReDim AnimationData(AnimationSize - 1)
        AnimationData = Buffer.ReadBytes(AnimationSize)
        CopyMemory ByVal VarPtr(Animation(N)), ByVal VarPtr(AnimationData(0)), AnimationSize
    End If
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    N = Buffer.ReadLong

    With MapNpc(N)
        .num = Buffer.ReadLong
        map.NPC(N) = .num
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .dir = Buffer.ReadLong
        .petData.Owner = Buffer.ReadLong
        ' Client use only
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
        
    End With
    
    If N > Npc_HighIndex Then
        Npc_HighIndex = N
    End If
    

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDead(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    N = Buffer.ReadLong
    Call ClearMapNpc(N)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDead", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(NPC(i).TranslatedName)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim Buffer As clsBuffer
Dim npcSize As Long
Dim npcData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    N = Buffer.ReadLong
    Dim UseFulData As Boolean
    UseFulData = Buffer.ReadByte
    If UseFulData Then
        SetNPCUsefulData N, Buffer.ReadBytes(Buffer.length)
    Else
        npcSize = LenB(NPC(N))
        ReDim npcData(npcSize - 1)
        npcData = Buffer.ReadBytes(npcSize)
        CopyMemory ByVal VarPtr(NPC(N)), ByVal VarPtr(npcData(0)), npcSize
    End If
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem i & ": " & Trim$(Resource(i).TranslatedName)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim ResourceNum As Long
Dim Buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ResourceNum = Buffer.ReadLong
    
    Dim UseFulData As Boolean
    UseFulData = Buffer.ReadByte
    
    ClearResource ResourceNum
    If UseFulData Then
        SetResourceUsefulData ResourceNum, Buffer.ReadBytes(Buffer.length)
    Else
        ResourceSize = LenB(Resource(ResourceNum))
        ReDim ResourceData(ResourceSize - 1)
        ResourceData = Buffer.ReadBytes(ResourceSize)
        CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    End If
    
    
    
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapKey(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Boolean
Dim X As Long
Dim Y As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    N = Buffer.ReadByte
    Dim f As Byte
    f = BTI(N)
    TempTile(X, Y).DoorOpen = f
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapKey", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorInit
    
    frmMain.PicBars(1).Visible = False
    frmMain.PicBars(2).Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEditMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleShopEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).TranslatedName)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleShopEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Shopnum As Long
Dim Buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Shopnum = Buffer.ReadLong
    
    'ShopSize = LenB(Shop(shopnum))
    'ReDim ShopData(ShopSize - 1)
    'ShopData = buffer.ReadBytes(ShopSize)
    'CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    ShopData = UnCompressData(Buffer.ReadBytes(Buffer.length), 2)
    CopyMemory ByVal VarPtr(Shop(Shopnum)), ByVal VarPtr(ShopData(0)), UBound(ShopData()) - LBound(ShopData())
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpellEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(Spell(i).TranslatedName)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpellEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim spellnum As Long
Dim Buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    spellnum = Buffer.ReadLong
    
    Dim UseFulData As Boolean
    UseFulData = Buffer.ReadByte
    If UseFulData Then
        SetSpellUsefulData spellnum, Buffer.ReadBytes(Buffer.length)
    Else
        SpellSize = LenB(Spell(spellnum))
        ReDim SpellData(SpellSize - 1)
        SpellData = Buffer.ReadBytes(SpellSize)
        CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    End If
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i) = Buffer.ReadLong
    Next
    
    BltPlayerSpells
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLeft(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call ClearPlayer(Buffer.ReadLong)
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLeft", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceCache(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Resource_Index = Buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim MapResource(0 To Resource_Index)

        For i = 0 To Resource_Index
            MapResource(i).ResourceState = Buffer.ReadByte
            MapResource(i).X = Buffer.ReadByte
            MapResource(i).Y = Buffer.ReadByte
            If MapResource(i).X > map.MaxX Then
                Debug.Print
            End If
            If MapResource(i).Y > map.MaxY Then
                Debug.Print
            End If
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If
    
    InitializeMapResources

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub HandleSingleResourceCache(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    If Not Resources_Init Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Dim Resource_Num As Long
    Resource_Num = Buffer.ReadLong
    
    MapResource(Resource_Num).ResourceState = Buffer.ReadByte
    MapResource(Resource_Num).X = Buffer.ReadByte
    MapResource(Resource_Num).Y = Buffer.ReadByte

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSendPing(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PingEnd = GetTickCount
    Ping = PingEnd - PingStart
    Call DrawPing
    Call SendAck
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSendPing", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub HandleDoorAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    With TempTile(X, Y)
        .DoorFrame = 1
        .DoorAnimate = 1 ' 0 = nothing| 1 = opening | 2 = closing
        .DoorTimer = GetTickCount
    End With
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDoorAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleActionMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long, message As String, Color As Long, tmpType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    message = Buffer.ReadString
    Color = Buffer.ReadLong
    tmpType = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong

    Set Buffer = Nothing
    
    CreateActionMsg message, Color, tmpType, X, Y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleActionMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBlood(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long, sprite As Long, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    Y = Buffer.ReadLong

    Set Buffer = Nothing
    
    ' randomise sprite
    sprite = Rand(1, BloodCount)
    
    ' make sure tile doesn't already have blood
    For i = 1 To MAX_BYTE
        If Blood(i).X = X And Blood(i).Y = Y Then
            ' already have blood :(
            Exit Sub
        End If
    Next
    
    ' carry on with the set
    BloodIndex = BloodIndex + 1
    If BloodIndex >= MAX_BYTE Then BloodIndex = 1
    
    With Blood(BloodIndex)
        .X = X
        .Y = Y
        .sprite = sprite
        .Timer = GetTickCount
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBlood", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    
    With AnimInstance(AnimationIndex)
        .Animation = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .LockType = Buffer.ReadByte
        .lockindex = Buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
    End With
    
    ' play the sound if we've got one
    PlayMapSound AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcVitals(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim mapnpcnum As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    mapnpcnum = Buffer.ReadLong
    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(mapnpcnum).vital(i) = Buffer.ReadLong
    Next
    
    Set Buffer = Nothing
    
    If GetPlayerPetMapNPCNum(MyIndex) = mapnpcnum Then
        RefreshPetVitals (MyIndex)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCooldown(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Slot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Slot = Buffer.ReadLong
    SpellCD(Slot) = GetTickCount
    
    BltPlayerSpells
    BltHotbar
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCooldown", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleClearSpellBuffer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SpellBuffer = 0
    SpellBufferTimer = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClearSpellBuffer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSayMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Access As Long
Dim Name As String
Dim message As String
Dim colour As Long
Dim header As String
Dim PK As Long
Dim saycolour As Long
Dim Chat As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    Access = Buffer.ReadLong
    PK = Buffer.ReadLong
    message = Buffer.ReadString
    header = Buffer.ReadString
    saycolour = Buffer.ReadLong
    Chat = Buffer.ReadLong
    
    colour = GetNameColorByJustice(PK)
    
    
    ' Check access level
    'If PK = NO Then
    'acces colors
        Select Case Access
            'Case 0
                'colour = RGB(255, 170, 70)
            Case 1
                colour = RGB(25, 200, 180)
            Case 2
                colour = RGB(100, 255, 0)
            Case 3
                colour = RGB(0, 155, 255)
            Case 4
                colour = RGB(100, 50, 255)
        End Select
    'Else
        'colour = RGB(220, 20, 20)
    'End If
    
    'frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    'frmMain.txtChat.SelColor = colour
    'frmMain.txtChat.SelText = vbNewLine & Header & Name & ": "
    'frmMain.txtChat.SelColor = saycolour
    'frmMain.txtChat.SelText = message
    'frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    Dim msg As ChatMsgRec
    
    msg.colour = colour
    msg.header = header & Name & ": "
    msg.saycolour = saycolour
    msg.text = message
    msg.ArrivedAt = GetTickCount
    
    Dim chatroom As Byte
    chatroom = Chat
    
    If chatroom > 0 And chatroom < ChatType_Count Then
        Call ListPush(ChatRooms(chatroom), msg)
    End If
    
    If Options.ChatToScreen = 2 Then
        ReOrderChat header & Name & ": " & message, colour
    End If
        
    If CanMsgBeDisplayed(chatroom) Then
    'If Options.Chat + 1 = chatroom Or chatroom > 4 Then
        frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
        frmMain.txtChat.SelColor = colour
        frmMain.txtChat.SelText = vbNewLine & header & Name & ": "
        frmMain.txtChat.SelColor = saycolour
        frmMain.txtChat.SelText = message
        frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    End If
        
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleOpenShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Shopnum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Shopnum = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    OpenShop Shopnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleOpenShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResetShopAction(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ShopAction = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResetShopAction", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStunned(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Dim BlockedAction As Byte
    BlockedAction = Buffer.ReadByte
    If BlockedAction > 0 And BlockedAction < PlayerActionsType.PlayerActions_Count Then
        BlockedActions(BlockedAction) = Buffer.ReadByte
    End If
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleStunned", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBank(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_BANK
        Bank.Item(i).num = Buffer.ReadLong
        Bank.Item(i).value = Buffer.ReadLong
    Next
    
    InBank = True
    frmMain.picBank.Visible = True
    BltBank
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBank", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    InTrade = Buffer.ReadLong
    frmMain.picTrade.Visible = True
    BltTrade
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCloseTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InTrade = 0
    frmMain.picTrade.Visible = False
    frmMain.lblTradeStatus.Caption = vbNullString
    ' re-blt any items we were offering
    BltInventory
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCloseTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeUpdate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim dataType As Byte
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    dataType = Buffer.ReadByte
    
    If dataType = 0 Then ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).num = Buffer.ReadLong
            TradeYourOffer(i).value = Buffer.ReadLong
        Next
        frmMain.lblYourWorth.Caption = Buffer.ReadLong & GetTranslation("Rupias")
        ' remove any items we're offering
        BltInventory
    ElseIf dataType = 1 Then 'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).num = Buffer.ReadLong
            TradeTheirOffer(i).value = Buffer.ReadLong
        Next
        frmMain.lblTheirWorth.Caption = Buffer.ReadLong & GetTranslation("Rupias")
    End If
    
    BltTrade
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeStatus(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim tradeStatus As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    tradeStatus = Buffer.ReadByte
    
    Set Buffer = Nothing
    
    Select Case tradeStatus
        Case 0 ' clear
            frmMain.lblTradeStatus.Caption = vbNullString
        Case 1 ' they've accepted
            frmMain.lblTradeStatus.Caption = "El otro jugador ha aceptado."
        Case 2 ' you've accepted
            frmMain.lblTradeStatus.Caption = "Esperando al otro jugador a que acepte."
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeStatus", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTarget(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    myTarget = Buffer.ReadLong
    myTargetType = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTarget", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHotbar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = Buffer.ReadLong
        Hotbar(i).sType = Buffer.ReadByte
    Next
    BltHotbar
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHotbar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHighIndex(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Player_HighIndex = Buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHighIndex", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSound(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long, entityType As Long, entityNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    entityType = Buffer.ReadLong
    entityNum = Buffer.ReadLong
    
    PlayMapSound X, Y, entityType, entityNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    theName = Buffer.ReadString
    
    Dialogue "Petición de comercio", theName & " te pide comerciar. ¿Quieres aceptar?", DIALOGUE_TYPE_TRADE, True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyInvite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    theName = Buffer.ReadString
    
    Dialogue "Invitación al equipo", theName & " te invita a un equipo. ¿Deseas unirte?", DIALOGUE_TYPE_PARTY, True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyInvite", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyUpdate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, i As Long, inParty As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    inParty = Buffer.ReadByte
    
    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        ' reset the labels
        For i = 1 To MAX_PARTY_MEMBERS
            frmMain.lblPartyMember(i).Caption = vbNullString
            frmMain.imgPartyHealth(i).Visible = False
            frmMain.imgPartySpirit(i).Visible = False
        Next
        ' exit out early
        Exit Sub
    End If
    
    ' carry on otherwise
    Party.Leader = Buffer.ReadLong
    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = Buffer.ReadLong
        If Party.Member(i) > 0 Then
            frmMain.lblPartyMember(i).Caption = Trim$(GetPlayerName(Party.Member(i)))
            frmMain.imgPartyHealth(i).Visible = True
            frmMain.imgPartySpirit(i).Visible = True
        Else
            frmMain.lblPartyMember(i).Caption = vbNullString
            frmMain.imgPartyHealth(i).Visible = False
            frmMain.imgPartySpirit(i).Visible = False
        End If
    Next
    Party.MemberCount = Buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyVitals(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim playerNum As Long, partyIndex As Long
Dim Buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' which player?
    playerNum = Buffer.ReadLong
    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = Buffer.ReadLong
        Player(playerNum).vital(i) = Buffer.ReadLong
    Next
    
    ' find the party number
    For i = 1 To MAX_PARTY_MEMBERS
        If Party.Member(i) = playerNum Then
            partyIndex = i
        End If
    Next
    
    ' exit out if wrong data
    If partyIndex <= 0 Or partyIndex > MAX_PARTY_MEMBERS Then Exit Sub
    
    ' hp bar
    frmMain.imgPartyHealth(partyIndex).Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
    ' spr bar
    frmMain.imgPartySpirit(partyIndex).Width = ((GetPlayerVital(playerNum, Vitals.mp) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.mp) / Party_SPRWidth)) * Party_SPRWidth
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleProjectile(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PlayerProjectile As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' create a new instance of the buffer
    Set Buffer = New clsBuffer
    
    ' read bytes from data()
    Buffer.WriteBytes Data()
    
    ' recieve projectile number
    PlayerProjectile = Buffer.ReadLong
    index = Buffer.ReadLong
    
    ' populate the values
    With Player(index).ProjecTile(PlayerProjectile)
    
        ' set the direction
        .direction = Buffer.ReadLong
        
        ' set the direction to support file format
        Select Case .direction
            Case DIR_DOWN
                .direction = 0
            Case DIR_UP
                .direction = 1
            Case DIR_RIGHT
                .direction = 2
            Case DIR_LEFT
                .direction = 3
        End Select
        
        ' set the pic
        .Pic = Buffer.ReadLong
        ' set the coordinates
        .X = GetPlayerX(index)
        .Y = GetPlayerY(index)
        ' get the range
        .range = Buffer.ReadLong
        ' get the damge
        .Damage = Buffer.ReadLong
        ' get the speed
        .Speed = Buffer.ReadLong
        
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleProjectile", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'ALATAR
Private Sub HandleQuestEditor()
    Dim i As Long
    
    With frmEditor_Quest
        Editor = EDITOR_TASKS
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_QUESTS
            .lstIndex.AddItem i & ": " & Trim$(Quest(i).TranslatedName)
        Next

        .Show
        .lstIndex.ListIndex = 0
        QuestEditorInit
    End With

End Sub

Private Sub HandleUpdateQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    N = Buffer.ReadLong
    ' Update the Quest
    QuestSize = Buffer.length
    ReDim QuestData(QuestSize - 1)
    QuestData = Buffer.ReadBytes(QuestSize)

    SetQuestData UnCompressData(QuestData, 2), N
    'CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    For i = 1 To MAX_QUESTS
        Player(MyIndex).PlayerQuest(i).Status = Buffer.ReadLong
        Player(MyIndex).PlayerQuest(i).ActualTask = Buffer.ReadLong
        Player(MyIndex).PlayerQuest(i).CurrentCount = Buffer.ReadLong
    Next
    
    RefreshQuestLog
    
    Set Buffer = Nothing
End Sub

Private Sub HandleQuestMessage(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long, QuestNum As Long, QuestNumForStart As Long
    Dim message As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    QuestNum = Buffer.ReadLong
    message = (Buffer.ReadString)
    QuestNumForStart = Buffer.ReadLong
    
    'todo?
    frmMain.lblQuestName = (Quest(QuestNum).TranslatedName)
    frmMain.lblQuestSay = message
    frmMain.lblQuestSubtitle = "Info:"
    frmMain.picQuestDialogue.Visible = True
    
    If QuestNumForStart > 0 And QuestNumForStart <= MAX_QUESTS Then
        frmMain.lblQuestAccept.Visible = True
        frmMain.lblQuestAccept.Tag = QuestNumForStart
    End If
        
    Set Buffer = Nothing
End Sub

'/ALATAR

Private Sub HandleDoorsEditor()
        Dim i As Long

        With frmEditor_Doors
                Editor = EDITOR_DOORS
                .lstIndex.Clear

                ' Add the names
                For i = 1 To MAX_DOORS
                        .lstIndex.AddItem i & ": " & Trim$(Doors(i).TranslatedName)
                Next

                .Show
                .lstIndex.ListIndex = 0
                DoorEditorInit
        End With

End Sub

Private Sub HandleUpdateDoors(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim DoorNum As Long
Dim Buffer As clsBuffer
Dim DoorSize As Long
Dim DoorData() As Byte
   
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
   
        DoorNum = Buffer.ReadLong
   
        DoorSize = LenB(Doors(DoorNum))
        ReDim DoorData(DoorSize - 1)
        DoorData = Buffer.ReadBytes(DoorSize)
   
        ClearDoor DoorNum
   
        CopyMemory ByVal VarPtr(Doors(DoorNum)), ByVal VarPtr(DoorData(0)), DoorSize
   
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "HandleUpdateDoors", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Private Sub HandleSpeechWindow(ByVal index As Long, ByRef Data() As Byte, ByVal EditorIndex As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

Dim Buffer As clsBuffer
Dim msg As String
Dim NPCNum As Long
Set Buffer = New clsBuffer
Buffer.WriteBytes Data()

msg = Buffer.ReadString
NPCNum = Buffer.ReadLong
'msg = GetTranslation(msg)
'todo?
    If Not FileExist(App.Path & "\data\graphics\faces\" & (NPC(NPCNum).sprite) & ".bmp", True) Then
        frmMain.picSpeechFace.Picture = LoadPicture(App.Path & "\data\graphics\faces\" & "default.bmp")
    Else
        frmMain.picSpeechFace.Picture = LoadPicture(App.Path & "\data\graphics\faces\" & (NPC(NPCNum).sprite) & ".bmp")
    End If

frmMain.picSpeech.Visible = True
frmMain.lblSpeech.Caption = "" & msg & ""
frmMain.lblSpeech.Visible = True
frmMain.picSpeechClose.Visible = True
frmMain.picSpeechFace.Visible = True

'play sound
PlaySound Sound_ButtonDialogue

End Sub

Private Sub HandleChatBubble(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, targetType As Long, target As Long, message As String, colour As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    target = Buffer.ReadLong
    targetType = Buffer.ReadLong
    message = Buffer.ReadString
    colour = Buffer.ReadLong
    
    AddChatBubble target, targetType, message, colour
    Set Buffer = Nothing
End Sub

Private Sub HandleLoading(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    frmMain.picLoad.Visible = True
    TempPlayer(MyIndex).IsLoading = True
    
    Set Buffer = Nothing
End Sub

Private Sub HandleDoneLoading(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
   
    frmMain.picLoad.Visible = False
    TempPlayer(MyIndex).IsLoading = False
    
    Set Buffer = Nothing
End Sub
Private Sub HandleDone(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    
    TempPlayer(MyIndex).IsLoading = False
    
End Sub

Private Sub HandleSendWeather(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Temp As Byte

Set Buffer = New clsBuffer
Buffer.WriteBytes Data()

Temp = Buffer.ReadLong

If Temp = 1 Then Rainon = True
If Temp = 0 Then Rainon = False

Set Buffer = Nothing
End Sub

Private Sub HandleMovementsEditor()
        Dim i As Long

        With frmEditor_Movements
                Editor = EDITOR_MOVEMENTS
                .lstIndex.Clear

                ' Add the names
                For i = 1 To MAX_MOVEMENTS
                        .lstIndex.AddItem i & ": " & Trim$(Movements(i).Name)
                Next

                .Show
                .lstIndex.ListIndex = 0
                MovementsEditorInit
        End With

End Sub

Private Sub HandleUpdatemovements(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim movementNum As Long
Dim Buffer As clsBuffer
Dim movementSize As Long
Dim movementData() As Byte
Dim i As Byte
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
   
        movementNum = Buffer.ReadLong
        ClearMovement movementNum
   
        'movementSize = LenB(Movements(movementNum))
        'ReDim movementData(movementSize - 1)
        'movementData = Buffer.ReadBytes(movementSize)
        With Movements(movementNum)
        
        .Name = Trim$(Buffer.ReadString)
        .Type = Buffer.ReadByte
        .MovementsTable.Actual = Buffer.ReadByte
        .MovementsTable.nelem = Buffer.ReadByte
        If .MovementsTable.nelem > 0 Then
            ReDim Movements(movementNum).MovementsTable.vect(1 To Movements(movementNum).MovementsTable.nelem)
            For i = 1 To .MovementsTable.nelem
                .MovementsTable.vect(i).Data.direction = Buffer.ReadByte
                .MovementsTable.vect(i).Data.NumberOfTiles = Buffer.ReadByte
            Next
        End If
        
        .Repeat = Buffer.ReadByte
        
        End With
   
        
   
        'CopyMemory ByVal VarPtr(Movements(movementNum)), ByVal VarPtr(movementData(0)), movementSize
   
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "HandleUpdatemovements", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Private Sub HandleActionsEditor()
        Dim i As Long

        With frmEditor_Actions
                Editor = EDITOR_ACTIONS
                .lstIndex.Clear

                ' Add the names
                For i = 1 To MAX_ACTIONS
                        .lstIndex.AddItem i & ": " & Trim$(Actions(i).TranslatedName)
                Next

                .Show
                .lstIndex.ListIndex = 0
                ActionsEditorInit
        End With

End Sub

Private Sub HandleUpdateActions(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim ActionNum As Long
Dim Buffer As clsBuffer
Dim ActionSize As Long
Dim ActionData() As Byte
Dim i As Byte
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
   
        ActionNum = Buffer.ReadLong
        ClearAction ActionNum
   
        'ActionSize = LenB(Actions(ActionNum))
        'ReDim ActionData(ActionSize - 1)
        'ActionData = Buffer.ReadBytes(ActionSize)
        With Actions(ActionNum)
        
        .Name = Trim$(Buffer.ReadString)
        .TranslatedName = Trim$(Buffer.ReadString)
        .Type = Buffer.ReadByte
        .Moment = Buffer.ReadByte
        .Data1 = Buffer.ReadLong
        .Data2 = Buffer.ReadLong
        .Data3 = Buffer.ReadLong
        .Data4 = Buffer.ReadLong
        
   
        'CopyMemory ByVal VarPtr(Actions(ActionNum)), ByVal VarPtr(ActionData(0)), ActionSize
        End With
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "HandleUpdateActions", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Private Sub HandleCustomSpritesEditor()
        Dim i As Long

        With frmEditor_CustomSprites
                Editor = EDITOR_CUSTOMSPRITES
                .lstIndex.Clear

                ' Add the names
                For i = 1 To MAX_CUSTOM_SPRITES
                        .lstIndex.AddItem i & ": " & Trim$(CustomSprites(i).Name)
                Next

                .Show
                .lstIndex.ListIndex = 0
                CustomSpritesEditorInit
        End With

End Sub

Private Sub HandleUpdateCustomSprites(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim CustomSpriteNum As Long
Dim Buffer As clsBuffer
Dim CustomSpriteSize As Long
Dim CustomSpriteData() As Byte

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
   
        CustomSpriteNum = Buffer.ReadLong
        
        ClearCustomSprite CustomSpriteNum
        SetCustomSpriteData CustomSpriteNum, Buffer.ReadBytes(Buffer.length)
       
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "HandleUpdateCustomSprites", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub



Private Sub HandleNPCCache(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim PIndex As Long
    Dim mapnum As Long
    Dim NPCNum As Long
    
    Dim i As Long
    
    Set Buffer = New clsBuffer
   
    Buffer.WriteBytes Data()
    
    mapnum = Buffer.ReadLong
    NPCNum = Buffer.ReadLong

    map.NPC(NPCNum) = Buffer.ReadLong
    MapNpc(NPCNum).num = Buffer.ReadLong

    'SaveMap (MapNum)
    
    Set Buffer = Nothing

End Sub

Private Sub HandlePetsEditor()
        Dim i As Long

        With frmEditor_Pets
                Editor = EDITOR_PETS
                .lstIndex.Clear

                ' Add the names
                For i = 1 To MAX_PETS
                        .lstIndex.AddItem i & ": " & Trim$(Pet(i).Name)
                Next

                .Show
                .lstIndex.ListIndex = 0
                PetsEditorInit
        End With

End Sub

Private Sub HandleUpdatePets(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PetNum As Long
Dim Buffer As clsBuffer
Dim i As Byte
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
   
        PetNum = Buffer.ReadLong
        ClearPet PetNum

        With Pet(PetNum)
        
        Pet(PetNum).Name = Buffer.ReadString
        Pet(PetNum).NPCNum = Buffer.ReadLong
        Pet(PetNum).TamePoints = Buffer.ReadInteger
        Pet(PetNum).ExpProgression = Buffer.ReadByte
        Pet(PetNum).PointsProgression = Buffer.ReadByte
        Pet(PetNum).MaxLevel = Buffer.ReadLong
                
        'CopyMemory ByVal VarPtr(Pets(PetNum)), ByVal VarPtr(PetData(0)), PetSize
        End With
        Set Buffer = Nothing
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "HandleUpdatePets", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Private Sub HandlePetData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PetSlot As Byte
Dim Buffer As clsBuffer
Dim i As Byte
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
   
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
   
        PetSlot = Buffer.ReadByte

        With Player(MyIndex).Pet(PetSlot)
        
        '.Name = Trim$(Buffer.ReadString)
        .points = Buffer.ReadInteger
        .Experience = Buffer.ReadLong
        .Level = Buffer.ReadLong
        .NumPet = Buffer.ReadByte
        
        For i = 1 To Stats.Stat_Count - 1
            .StatsAdd(i) = Buffer.ReadByte
        Next
        
        End With
        
        Player(MyIndex).ActualPet = PetSlot
        
        Set Buffer = Nothing
   
        Call RefreshPetData(MyIndex)
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "HandleUpdatePets", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Private Sub HandlePlayerPetStats(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To Stats.Stat_Count - 1
        SetPlayerPetStat MyIndex, i, Buffer.ReadLong
        frmMain.lblCharStat(i + Stats.Stat_Count).Caption = GetPlayerPetStat(MyIndex, i)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleOpenTriforce(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    frmMain.picTriforce.Visible = True
End Sub


Private Sub HandleOnIce(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(MyIndex).onIce = Buffer.ReadByte
End Sub

Private Sub HandleIceDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(MyIndex).IceDir = Buffer.ReadByte
End Sub

Private Sub HandleBags(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(MyIndex).RupeeBags = Buffer.ReadByte
    BltInventory
End Sub

Private Sub HandlePoints(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call SetPlayerPOINTS(MyIndex, Buffer.ReadLong)
    frmMain.lblPoints.Caption = GetPlayerPOINTS(MyIndex)
    
    Dim X As Long
    
    For X = 1 To Stats.Stat_Count - 1
        frmMain.lblCharStat(X).Caption = GetPlayerStat(MyIndex, X)
    Next
        
    
    If GetPlayerPOINTS(MyIndex) > 0 Then
        For X = 1 To Stats.Stat_Count - 1
            If GetPlayerStat(index, X) < 255 Then
                frmMain.lblTrainStat(X).Visible = True
            Else
                frmMain.lblTrainStat(X).Visible = False
            End If
        Next
    Else
        For X = 1 To Stats.Stat_Count - 1
            frmMain.lblTrainStat(X).Visible = False
        Next
    End If
End Sub

Private Sub HandleLevel(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call SetPlayerLevel(MyIndex, Buffer.ReadLong)
     ' Set the character windows
    frmMain.lblCharName = GetPlayerName(MyIndex) & " - Level " & GetPlayerLevel(MyIndex)
End Sub

Private Sub HandleJustice(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Dim i As Long
    i = Buffer.ReadLong
    Player(i).PK = Buffer.ReadByte
End Sub
    
Private Sub HandlePlayerAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Dim i As Long
    i = Buffer.ReadLong
    Player(i).Attacking = 1
    Player(i).AttackTimer = GetTickCount()
End Sub

Private Sub HandleAccounts(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If Not DirExists(App.Path & "\data\accounts") Then
        Call MkDir(App.Path & "\data\accounts")
    End If

    
    Dim i As Long
    i = Buffer.ReadLong
    
    Do While i > 0
        Dim Login As String, Size As Long
        Login = Buffer.ReadString
        Size = Buffer.ReadLong
        SavePlayer Buffer.ReadBytes(Size), Login
        i = i - 1
    Loop
    
    Set Buffer = Nothing
End Sub

Private Sub HandleSaveFiles(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    
    Dim Compressed As Boolean
    Compressed = Buffer.ReadByte
    If Compressed Then
        Dim CompBuffer As clsBuffer
        Set CompBuffer = New clsBuffer
        CompBuffer.WriteBytes Buffer.ReadBytes(Buffer.length)
        CompBuffer.BufferDeCompress
        Set Buffer = CompBuffer
    End If
    
    
    
    Dim dir As String
    dir = Buffer.ReadString
    
    If Not DirExists(App.Path & dir) Then
        Call MkDir(App.Path & dir)
    End If
    
    Dim N As Long
    N = Buffer.ReadLong
    
    Do While N > 0
        Dim Filename As String, FileSize As Long, NotNull As Boolean
        NotNull = Buffer.ReadByte
        Filename = Buffer.ReadString
        
        If NotNull Then
            FileSize = Buffer.ReadLong
            Call SaveFile(dir, Filename, Buffer.ReadBytes(FileSize))
        Else
            Call SaveEmptyFile(dir, Filename)
        End If
        N = N - 1
    Loop

End Sub

Public Sub SaveFile(ByRef dir As String, ByRef Name As String, ByRef Data() As Byte)
    
    Dim f As Long
    f = FreeFile
    Open App.Path & dir & Name For Binary As #f
    Put f, , Data
    Close #f
End Sub

Public Sub SaveEmptyFile(ByRef dir As String, ByRef Name As String)
    If Not DirExists(App.Path & dir) Then
        Call MkDir(App.Path & dir)
    End If
    
    Dim f As Long
    f = FreeFile
    Open App.Path & dir & Name For Binary As #f
    Close #f
End Sub


Private Sub HandleCharList(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Dim BufferLenght As Long
    
    If Not DirExists(App.Path & "\data\accounts") Then
        Call MkDir(App.Path & "\data\accounts")
    End If
    
    If FileExist("data\accounts\charlist.txt") Then
        Kill App.Path & "\data\accounts\charlist.txt"
    End If
    
    Dim f As Long
    f = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Binary As #f
    Put f, , Buffer.ReadBytes(Buffer.length)
    Close #f
    
    Set Buffer = Nothing

    
End Sub

Private Sub HandleBanks(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Dim BufferLenght As Long
    Dim nLenght As Long
    nLenght = Buffer.ReadLong
    BufferLenght = Buffer.length
    Dim i As Long
    i = BufferLenght \ nLenght
    
    If Not DirExists(App.Path & "\data\banks") Then
        Call MkDir(App.Path & "\data\banks")
    End If
    
    If Not FileExist("data\accounts\charlist.txt") Then
        Dim f As Long
        f = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Output As #f
        Close #f
    End If
    
    Do While i > 0
        Dim Login As String * ACCOUNT_LENGTH
        Dim TempBank As ServerBankRec
        Dim BankData() As Byte
        Login = Buffer.ReadString
        Dim j As Long
        For j = 1 To MAX_BANK
            TempBank.Item(j).num = Buffer.ReadLong
            TempBank.Item(j).value = Buffer.ReadLong
        Next
        SaveBank TempBank, Login
        i = i - 1
    Loop
End Sub

Private Sub HandleGuilds(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Dim i As Long
    Dim b As Long
    
    If Not DirExists(App.Path & "\data\guilds") Then
        Call MkDir(App.Path & "\data\guilds")
    End If
    
    If Not DirExists(App.Path & "\data\guildnames") Then
        Call MkDir(App.Path & "\data\guildnames")
    End If

    Do While Buffer.length <> 0
        Dim TempGuild As ServerGuildRec
        
        TempGuild.In_Use = Buffer.ReadByte
        TempGuild.Guild_Name = Buffer.ReadString
        TempGuild.Guild_Fileid = Buffer.ReadLong
        TempGuild.Guild_Color = Buffer.ReadInteger
        TempGuild.Guild_MOTD = Buffer.ReadString
        TempGuild.Guild_RecruitRank = Buffer.ReadInteger
        
        'Get Members
        For i = 1 To MAX_GUILD_MEMBERS
            TempGuild.Guild_Members(i).Used = Buffer.ReadByte
            TempGuild.Guild_Members(i).User_Login = Buffer.ReadString
            TempGuild.Guild_Members(i).User_Name = Buffer.ReadString
            TempGuild.Guild_Members(i).Founder = Buffer.ReadByte
            TempGuild.Guild_Members(i).Rank = Buffer.ReadInteger
            TempGuild.Guild_Members(i).Comment = Buffer.ReadString
        Next i
        
        'Get Ranks
        For i = 1 To MAX_GUILD_RANKS
            TempGuild.Guild_Ranks(i).Name = Buffer.ReadString
            TempGuild.Guild_Ranks(i).Used = Buffer.ReadByte
            For b = 1 To MAX_GUILD_RANKS_PERMISSION
                TempGuild.Guild_Ranks(i).RankPermission(b) = Buffer.ReadByte
            Next b
        Next i

        SaveGuild TempGuild, TempGuild.Guild_Fileid
        i = i - 1
    Loop
End Sub


Private Sub HandlePlayerSprite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Dim i As Long
    i = Buffer.ReadLong
    Call SetPlayerSprite(i, Buffer.ReadLong)
    Call SetPlayerCustomSprite(i, Buffer.ReadByte)
    
End Sub


Private Sub HandleMaxWeight(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Call SetPlayerMaxWeight(MyIndex, Buffer.ReadLong)
    
End Sub

Private Sub HandleQuestion(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim header As String
Dim question As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    header = Buffer.ReadString
    question = Buffer.ReadString
    
    Dialogue header, question, DIALOGUE_TYPE_QUESTION, True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleQuestion", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub HandleKillPoints(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    SetKillPoints Buffer.ReadByte, Buffer.ReadLong 'status and points
   
End Sub

Private Sub HandleBonusPoints(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    Call SetBonusPoints(MyIndex, Buffer.ReadLong)
   
End Sub

Private Sub HandlePlayerState(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    Call SetPlayerState(Buffer.ReadLong, Buffer.ReadByte)
   
End Sub




    
    
