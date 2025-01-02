Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal index As Long, ByVal NeedData As Boolean)
    If Not IsPlaying(index) Then

        Call JoinGame(index, NeedData)
        Call AddLog(index, GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has logged in.", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has logged in.")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal index As Long, ByVal NeedData As Boolean)
    Dim i As Long, j As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
    frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
    frmServer.lvwInfo.ListItems(index).SubItems(3) = GetPlayerName(index)

    
    ' send the login ok
    SendLoginOk index
    DoEvents
    TotalPlayersOnline = TotalPlayersOnline + 1
    CalculateSleepTime
    ByteCounter = 0
    DoEvents
    CheckPlayerStateAtJoin index
    ' Send some more little goodies, no need to explain these
    
    SendMaxWeight index
    
    Call CheckEquippedItems(index)
    
    If NeedData Then
        Call SendClasses(index)
        
        Call SendItems(index) 'done
        
        Call SendAnimations(index) 'done
        
        Call SendNpcs(index) 'done
        
        Call SendShops(index)
        
        Call SendSpells(index) 'done
        
        Call SendResources(index) 'done
        
        Call SendQuests(index)
        
        Call SendPets(index)
        
        Call SendCustomSprites(index)
        
        If GetPlayerAccess_Mode(index) >= ADMIN_CREATOR Then
            Call SendMovements(index)
            
            Call SendActions(index)
            
            Call SendDoors(index)
            
        End If
    End If
    
    Call SendInventory(index)
    
    Call SendWornEquipment(index)
    
    Call SendMapEquipment(index)
    
    Call SendPlayerSpells(index)
    
    Call SendHotbar(index)
    
    Call SendWeather(index)
    
    Call SendRunningSprites(index)
    
    SendKillPoints index
    SendPlayerBonusPoints index
    ComputePlayerSpeed index
    'CheckSpeedHack index
    CheckToAddMap GetPlayerMap(index)
    SendStaminaInfo index
    AddMapPlayer index, GetPlayerMap(index)
    
    
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    
    CheckPlayerOutOfExp index
    SendEXP index
    
    Call ComputeAllPlayerStats(index)
    Call SendStats(index)
        
    ' Warp the player to his saved location
    Call PlayerSpawn(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess_Mode(index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(index) & " has connected.", BrightGreen, False)
    Else
        Call GlobalMsg(GetPlayerName(index) & " has connected.", BrightGreen, False)
    End If
    
    ForwardGlobalMsg "[Hub - " & SERVER_NAME & "] You sense someone not of this world.. " & GetPlayerName(index) & " has connected."
    
    ' Send welcome messages
    Call SendWelcome(index)
    
    If frmServer.chkTroll.Value = vbChecked Then PlayerMsg index, "You are on a Troll server. Type /admin for admin menu.", BrightRed
    
    'Do all the guild start up checks
    Call GuildLoginCheck(index)
    
    If player(index).points > 0 Then
        PlayerMsg index, "You have " & player(index).points & " unspent stat points!", White
    End If
    
    'miscellanious
    Call InitPlayerPets(index)
    Call SendPetData(index, TempPlayer(index).TempPet.ActualPet)
    Call SetPlayerWeight(index, CalculatePlayerWeight(index))
    
    'ping
    TempPlayer(index).Req = False
    
    If IsPlayerOverWeight(index) Then
        PlayerMsg index, "You're carrying too much weight! You can't move, throw items on the floor to lower your weight.", BrightRed
    End If
    

    
    If ArePlayersOnMap(GetPlayerMap(index)) > 0 Then
        Dim a As Variant
        For Each a In GetMapPlayerCollection(GetPlayerMap(index))
            If a <> index Then
                SendMapEquipmentTo a, index
            End If
        Next
    End If
    
    ' Send the flag so they know they can start doing stuff
    SendInGame index
End Sub

Sub LeftGame(ByVal index As Long)
    Dim N As Long, i As Long
    Dim tradeTarget As Long
    
    If TempPlayer(index).InGame Then
        TempPlayer(index).InGame = False

        ' Check if player was the only player on the map and stop npc processing if so

        DeleteMapPlayer index, GetPlayerMap(index)
        
        
        ' cancel any trade they're in
        If TempPlayer(index).InTrade > 0 Then
            tradeTarget = TempPlayer(index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has declined the trade request.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        If IsInTeam(index) Then
            ClearTeamPlayer index
        End If
        
        ' leave party.
        Party_PlayerLeave index

        If player(index).GuildFileId > 0 Then
            'Set player online flag off
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(player(index).GuildMemberId).Online = False
            Call CheckUnloadGuild(TempPlayer(index).tmpGuildSlot)
        End If
        
        ' save and clear data.
        Call SavePlayer(index)
        Call SaveBank(index)
        Call ClearBank(index)

        ' Send a global message that he/she left
        If GetPlayerAccess_Mode(index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(index) & " has disconnected.", BrightRed, False)
        Else
            Call GlobalMsg(GetPlayerName(index) & " has disconnected.", BrightRed)
        End If

            ForwardGlobalMsg "[Hub - " & SERVER_NAME & "] " & GetPlayerName(index) & " has disconnected."

        Call TextAdd(GetPlayerName(index) & " has disconnected .")
        Call SendLeftGame(index)
        TotalPlayersOnline = TotalPlayersOnline - 1
        CalculateSleepTime
        

        
    End If
    
    UnLockPlayerLogin player(index).login
    Call ClearPlayer(index)
    
End Sub

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(index, Armor)
    Helm = GetPlayerEquipment(index, helmet)
    GetPlayerProtection = (GetPlayerStat(index, Stats.Endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(Helm).Data2
    End If

End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim N As Long

    If GetPlayerEquipment(index, Weapon) > 0 Then
        N = (Rnd) * 1.333

        If N = 1 Then
            i = (GetPlayerStat(index, Stats.Strength) \ 2) + (GetPlayerLevel(index) \ 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim N As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(index, Shield)

    If ShieldSlot > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = (GetPlayerStat(index, Stats.Endurance) \ 2) + (GetPlayerLevel(index) \ 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function







Sub ForcePlayerMove(ByVal index As Long, ByVal Movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If Movement < 1 Or Movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = map(GetPlayerMap(index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = map(GetPlayerMap(index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove index, Direction, Movement, True
End Sub

Sub CheckEquippedItems(ByVal index As Long)
    Dim slot As Long
    Dim ItemNum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(index, i)

        If ItemNum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If item(ItemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment index, 0, i
                Case Equipment.Armor

                    If item(ItemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment index, 0, i
                Case Equipment.helmet

                    If item(ItemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment index, 0, i
                Case Equipment.Shield

                    If item(ItemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment index, 0, i
            End Select

        Else
            SetPlayerEquipment index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    FindOpenInvSlot = 0
    
    If isItemStackable(ItemNum) Then
        Dim Tempitemnum As Long
        Dim FreeSlot As Long
        FreeSlot = 0
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            Tempitemnum = GetPlayerInvItemNum(index, i)
    
            If Tempitemnum = ItemNum Then
                'found the stackable item, out of function
                FindOpenInvSlot = i
                Exit Function
            ElseIf Tempitemnum = 0 And FindOpenInvSlot = 0 Then
                'first free slot will be used in case of the itemnum does not exist
                FindOpenInvSlot = i
            End If

        Next
    Else
        For i = 1 To MAX_INV
            ' Try to find an open free slot
            If GetPlayerInvItemNum(index, i) = 0 Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next
    End If

End Function

Public Function CanGiveItem(ByVal index As Long, ByVal ItemNum As Long, ByVal itemval As Long, ByRef GivenValue As Long) As Long
'checks if player can take an item: if can take -> givenvalue = amount that will be given, CanGiveItem = appropiate slot to be given
If index < 1 Or index > MAX_PLAYERS Or ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function

Dim i As Long
i = FindOpenInvSlot(index, ItemNum)

If i > 0 Then
    If ItemNum = 1 Then
        GivenValue = GetGivenMoney(index, GetPlayerInvItemValue(index, i), itemval)
    Else
        GivenValue = itemval
    End If
    
    Dim val As Long
    
    If isItemStackable(ItemNum) Then
        val = GivenValue
    Else
        val = 1
    End If
    
    If CanPlayerHoldWeight(index, GetItemValWeight(ItemNum, val)) Then
        CanGiveItem = i
    Else
        PlayerMsg index, "You can't take any more weight.", BrightRed
        CanGiveItem = 0
    End If
Else
    PlayerMsg index, "You don't have space in your inventory.", BrightRed
    CanGiveItem = 0
End If


        

End Function

Public Function GetGivenMoney(ByVal index As Long, ByVal initialvalue As Long, ByVal Value As Long) As Long
    If GetPlayerMaxMoney(index) < initialvalue + Value Then
        GetGivenMoney = GetPlayerMaxMoney(index) - initialvalue
    Else
        GetGivenMoney = Value
    End If
End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    If Not IsPlaying(index) Then Exit Function
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function HasItem(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If isItemStackable(ItemNum) Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal index As Long, ByVal ItemNum As Long, ByVal itemval As Long, Optional ByVal UpdateWeight As Boolean = True) As Boolean
    Dim i As Long
    Dim N As Long
    Dim TakenValue As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If isItemStackable(ItemNum) Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If itemval >= GetPlayerInvItemValue(index, i) Then
                    TakenValue = GetPlayerInvItemValue(index, i)
                    Call SetPlayerInvItemNum(index, i, 0)
                    Call SetPlayerInvItemValue(index, i, 0)
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - itemval)
                    TakenValue = itemval
                End If
            Else
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                TakeInvItem = True
                TakenValue = 1
            End If
            
            Call SendInventoryUpdate(index, i)
            If UpdateWeight Then Call SetPlayerWeight(index, GetPlayerWeight(index) - GetItemValWeight(ItemNum, TakenValue))
            Exit For
        End If

    Next
    
End Function

Function TakeInvSlot(ByVal index As Long, ByVal invSlot As Byte, ByRef itemval As Long, Optional ByVal Update As Boolean = False) As Boolean
    'itemval returns the taken value
    Dim ItemNum As Integer
    Dim NewItemVal As Long
    Dim NewItemNum As Long
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then Exit Function
    
    ItemNum = GetPlayerInvItemNum(index, invSlot)

    ' Prevent subscript out of range
    If ItemNum < 1 Then Exit Function
    
    If isItemStackable(ItemNum) Then
        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If itemval >= GetPlayerInvItemValue(index, invSlot) Then
            NewItemVal = 0
            NewItemNum = 0
            itemval = GetPlayerInvItemValue(index, invSlot)
        Else
            NewItemVal = GetPlayerInvItemValue(index, invSlot) - itemval
            NewItemNum = GetPlayerInvItemNum(index, invSlot)
        End If
    Else
        NewItemVal = 0
        NewItemNum = 0
        itemval = 1
    End If
    
    SetPlayerInvItemNum index, invSlot, NewItemNum
    SetPlayerInvItemValue index, invSlot, NewItemVal
    SetPlayerWeight index, GetPlayerWeight(index) - GetItemValWeight(ItemNum, itemval)
        
    ' Send the inventory update
    If Update Then
        Call SendInventoryUpdate(index, invSlot)
    End If
    
End Function
Sub GiveInvSlot(ByVal index As Long, ByVal slot As Long, ByVal ItemNum As Long, ByVal Value As Long, Optional ByVal SendUpdate As Boolean = True)
    If index < 1 Or index > MAX_PLAYERS Or slot < 1 Or slot > MAX_INV Then Exit Sub
    
    Dim SetValue As Long
    If isItemStackable(ItemNum) Then
        SetValue = GetPlayerInvItemValue(index, slot) + Value
    Else
        SetValue = 1
        Value = 1
    End If
    
    Call SetPlayerInvItemNum(index, slot, ItemNum)
    Call SetPlayerInvItemValue(index, slot, SetValue)
    Call SetPlayerWeight(index, GetPlayerWeight(index) + GetItemValWeight(ItemNum, Value))
    
    If SendUpdate Then SendInventoryUpdate index, slot
    
End Sub

Function GiveInvItem(ByVal index As Long, ByVal ItemNum As Long, ByVal itemval As Long, Optional ByVal SendUpdate As Boolean = True, Optional ByVal UpdateWeight As Boolean = True) As Boolean
    Dim i As Long
    Dim Value As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(index, ItemNum)
    

    ' Check to see if inventory is full
    If i <> 0 Then
        If CanPlayerHoldWeight(index, GetItemValWeight(ItemNum, itemval)) Or Not UpdateWeight Then
            Call SetPlayerInvItemNum(index, i, ItemNum)
            
            If ItemNum = 1 Then
                Value = CheckMoneyAdd(index, GetPlayerInvItemValue(index, i), itemval)
                itemval = Value - GetPlayerInvItemValue(index, i)
            Else
                Value = GetPlayerInvItemValue(index, i) + itemval
            End If
                 
            Call SetPlayerInvItemValue(index, i, Value)
            
            If SendUpdate Then Call SendInventoryUpdate(index, i)
            If UpdateWeight Then Call SetPlayerWeight(index, GetPlayerWeight(index) + GetItemValWeight(ItemNum, itemval))
            
            GiveInvItem = True
        Else
            Call PlayerMsg(index, "You can't support the weight of the object.", BrightRed)
            GiveInvItem = False
        End If
    Else
        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

End Function

Function HasSpell(ByVal index As Long, ByVal spellnum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, i) = spellnum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal index As Long)
    Dim i As Long
    Dim N As Long
    Dim mapnum As Long
    Dim msg As String

    If Not IsPlaying(index) Then Exit Sub
    mapnum = GetPlayerMap(index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(mapnum, i).Num > 0) And (MapItem(mapnum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(mapnum, i).X = GetPlayerX(index)) Then
                    If (MapItem(mapnum, i).Y = GetPlayerY(index)) Then
                        ' Find open slot
                        'n = FindOpenInvSlot(index, MapItem(mapnum, i).Num)
                        Dim ItemNum As Long
                        ItemNum = MapItem(mapnum, i).Num
    
                        If GiveInvItem(index, MapItem(mapnum, i).Num, MapItem(mapnum, i).Value, True) Then
                        ' Open slot available?
                        'If n <> 0 Then
                            ' Set item in players inventor
                            'Call SetPlayerInvItemNum(index, n, MapItem(mapnum, i).Num)
                            
                            
                            If isItemStackable(ItemNum) Then
                                'Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(mapnum, i).Value)
                                msg = MapItem(mapnum, i).Value & " " & Trim$(item(ItemNum).Name)
                            Else
                                'Call SetPlayerInvItemValue(index, n, 0)
                                msg = Trim$(item(ItemNum).Name)
                            End If
                            
                            If Not MapItem(mapnum, i).isDrop Then
                                Call AddMapWaitingItem(mapnum, GetPlayerX(index), GetPlayerY(index))
                            End If
                            
                            ClearMapItem i, mapnum
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), 0, 0)
                         
                                                    
                            SendActionMsg GetPlayerMap(index), msg, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                            'ALATAR
                            Call CheckTasks(index, QUEST_TYPE_GOGATHER, GetItemNum(Trim$(item(ItemNum).Name)))
                            '/ALATAR
                            Exit For
                        Else
                            'Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal index As Long, ByVal mapItemNum As Long)
Dim mapnum As Long

    mapnum = GetPlayerMap(index)
    
    ' no lock or locked to player?
    If MapItem(mapnum, mapItemNum).playerName = vbNullString Or MapItem(mapnum, mapItemNum).playerName = Trim$(GetPlayerName(index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal invNum As Long, ByVal amount As Long, Optional ByVal SayMsg As Boolean = True)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Or TempPlayer(index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(index, invNum) > 0) Then
        If (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
            If IsPlayerOverWeight(index) Then
                Call TakeInvSlot(index, invNum, amount, True)
                Exit Sub
            End If
        
            i = FindOpenMapItemSlot(GetPlayerMap(index))

            If i <> 0 Then
                MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, invNum)
                MapItem(GetPlayerMap(index), i).X = GetPlayerX(index)
                MapItem(GetPlayerMap(index), i).Y = GetPlayerY(index)
                MapItem(GetPlayerMap(index), i).playerName = Trim$(GetPlayerName(index))
                MapItem(GetPlayerMap(index), i).playerTimer = GetRealTickCount + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(index), i).isDrop = True
                MapItem(GetPlayerMap(index), i).Timer = GetRealTickCount + ITEM_DESPAWN_TIME

                If isItemStackable(GetPlayerInvItemNum(index, invNum)) Then
                    ' Check if its more then they have and if so drop it all
                    If amount >= GetPlayerInvItemValue(index, invNum) Then
                        MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, invNum)
                        If SayMsg Then Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " throws " & GetPlayerInvItemValue(index, invNum) & " " & Trim$(item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                    Else
                        MapItem(GetPlayerMap(index), i).Value = amount
                        If SayMsg Then Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " throws " & amount & " " & Trim$(item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                    End If
                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(index), i).Value = 0
                    ' send message
                    If SayMsg Then Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " throws " & CheckGrammar((item(GetPlayerInvItemNum(index, invNum)).Name)) & ".", Yellow)
                End If
                
                Call TakeInvSlot(index, invNum, amount, True)

                ' Send inventory update
                'Call SendInventoryUpdate(index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, amount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), Trim$(GetPlayerName(index)), MapItem(GetPlayerMap(index), i).isDrop)
            Else
                If SayMsg Then Call PlayerMsg(index, "Too many items on the floor.", BrightRed)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    Dim points As Byte
    
    level_count = 0
    
    Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
        expRollover = GetPlayerExp(index) - GetPlayerNextLevel(index)
        
        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + 1) Then
            Call SetPlayerExp(index, GetPlayerNextLevel(index))
            Exit Sub
        End If
        
        points = 3
        'Check if triforce
        points = points + GetPlayerTriforcesNum(index)
        
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + points)
        Call SetPlayerExp(index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 And Not LPE(index) Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(index) & " has grown by " & level_count & " level.", Brown, False
        Else
            'plural
            GlobalMsg GetPlayerName(index) & " has grown by " & level_count & " levels.", Brown, False
        End If
        SendEXP index
        SendPoints index
        SendLevel index
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seLevelUp, 1
    End If
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal index As Long) As String
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerLogin = Trim$(player(index).login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal login As String)
    player(index).login = login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerPassword = Trim$(player(index).password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal password As String)
    player(index).password = password
End Sub

Function GetPlayerName(ByVal index As Long) As String
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerName = Trim$(player(index).Name)
End Function

Function GetPlayerNameNS(ByVal index As Long) As String
    GetPlayerNameNS = player(index).Name
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    player(index).Name = Name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerClass = player(index).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    player(index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerSprite = player(index).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    player(index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerLevel = player(index).level
End Function

Function SetPlayerLevel(ByVal index As Long, ByVal level As Long) As Boolean
    SetPlayerLevel = False
    If level > MAX_LEVELS Then Exit Function
    player(index).level = level
    SetPlayerLevel = True
End Function


Function GetPlayerExp(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerExp = player(index).exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal exp As Long)
    player(index).exp = exp
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerAccess = player(index).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    If index <= 0 Then Exit Sub
    player(index).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Byte
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerPK = player(index).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    player(index).PK = PK
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerVital = player(index).vital(vital)
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal vital As Vitals, ByVal Value As Long)
    player(index).vital(vital) = Value
    If GetPlayerVital(index, vital) > GetPlayerMaxVital(index, vital) Then
        player(index).vital(vital) = GetPlayerMaxVital(index, vital)
    End If

    If GetPlayerVital(index, vital) < 0 Then
        player(index).vital(vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal index As Long, ByVal stat As Stats) As Long
    Dim X As Long, i As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    
    GetPlayerStat = TempPlayer(index).Stats(stat)
    Exit Function
    
    X = player(index).stat(stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If player(index).Equipment(i) > 0 Then
            If item(player(index).Equipment(i)).Add_Stat(stat) > 0 Then
                X = X + item(player(index).Equipment(i)).Add_Stat(stat)
            End If
        End If
    Next
    
    X = X + GetPlayerStatBuffer(index, stat)
    
    GetPlayerStat = X
End Function

Public Sub ComputePlayerStat(ByVal index As Long, ByVal stat As Stats)
    Dim X As Long, i As Long
    If index > MAX_PLAYERS Then Exit Sub
    If index <= 0 Then Exit Sub
    
    X = player(index).stat(stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If player(index).Equipment(i) > 0 Then
            If item(player(index).Equipment(i)).Add_Stat(stat) > 0 Then
                X = X + item(player(index).Equipment(i)).Add_Stat(stat)
            End If
        End If
    Next
    
    X = X + GetPlayerStatBuffer(index, stat)
    
    TempPlayer(index).Stats(stat) = X

End Sub

Public Sub ComputeAllPlayerStats(ByVal index As Long)
    Dim i As Byte
    For i = 1 To Stats.Stat_Count - 1
        ComputePlayerStat index, i
    Next
End Sub

Public Function GetPlayerRawStat(ByVal index As Long, ByVal stat As Stats) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerRawStat = player(index).stat(stat)
End Function

Public Sub SetPlayerStat(ByVal index As Long, ByVal stat As Stats, ByVal Value As Long)
    player(index).stat(stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerPOINTS = player(index).points
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal points As Long)
    If points <= 0 Then points = 0
    player(index).points = points
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerMap = player(index).map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal mapnum As Long)

    If mapnum > 0 And mapnum <= MAX_MAPS Then
        player(index).map = mapnum
    End If

End Sub

Function GetPlayerX(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerX = player(index).X
End Function

Sub SetPlayerX(ByVal index As Long, ByVal X As Long)
If X < 0 Then Exit Sub
    player(index).X = X
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerY = player(index).Y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal Y As Long)
If Y < 0 Then Exit Sub
    player(index).Y = Y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerDir = player(index).dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal dir As Long)
    player(index).dir = dir
End Sub

Function GetPlayerIP(ByVal index As Long, Optional ByVal genuine As Boolean = False) As String

    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    If genuine Then
        GetPlayerIP = frmServer.Socket(index).RemoteHostIP
    Else
        If LPE(index) Then
            GetPlayerIP = RandomizeIP
        Else
            GetPlayerIP = frmServer.Socket(index).RemoteHostIP
        End If


    End If
End Function

Function GetPlayerHost(ByVal index As Long)
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerHost = frmServer.Socket(index).RemoteHost
    End If
End Function

Function RandomizeIP() As String
    Dim a As Integer
    Dim i As Byte
    i = RAND(3, 4)
    While i > 0
    
        a = RAND(111, 999)
        RandomizeIP = RandomizeIP + CStr(a)
        If i > 1 Then RandomizeIP = RandomizeIP + "."
            
        i = i - 1
    Wend
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    If invSlot = 0 Then Exit Function
    GetPlayerInvItemNum = player(index).Inv(invSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long, ByVal ItemNum As Long)
    player(index).Inv(invSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerInvItemValue = player(index).Inv(invSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long, ByVal Itemvalue As Long)
    player(index).Inv(invSlot).Value = Itemvalue
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal spellslot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
    GetPlayerSpell = player(index).Spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal spellslot As Long, ByVal spellnum As Long)
    player(index).Spell(spellslot) = spellnum
End Sub

Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot <= 0 Or EquipmentSlot > Equipment_Count - 1 Then Exit Function
    GetPlayerEquipment = player(index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    player(index).Equipment(EquipmentSlot) = invNum
End Sub
Function GetPlayerVisible(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If index <= 0 Then Exit Function
GetPlayerVisible = player(index).Visible
End Function
Sub SetPlayerVisible(ByVal index As Long, ByVal Visible As Long)
player(index).Visible = Visible
End Sub

Sub SwapInvEquipment(ByVal index As Long, ByVal invSlot As Long, ByVal EquipmentSlot As Long)
'Player tries to equip itemnum
If index < 1 Or index > MAX_PLAYERS Or invSlot < 1 Or invSlot > MAX_INV Or EquipmentSlot < 1 Or EquipmentSlot > Equipment.Equipment_Count - 1 Then Exit Sub


Dim TempItem As Long
TempItem = GetPlayerInvItemNum(index, invSlot)

Dim NewValue As Long
NewValue = 0
If GetPlayerEquipment(index, EquipmentSlot) > 0 Then
    NewValue = 1
End If

'Set The inventory
Call SetPlayerInvItemNum(index, invSlot, GetPlayerEquipment(index, EquipmentSlot))
Call SetPlayerInvItemValue(index, invSlot, NewValue)

'And The equipment
Call SetPlayerEquipment(index, TempItem, EquipmentSlot)
Call ComputeAllPlayerStats(index)


End Sub


' ToDo
Sub OnDeath(ByVal index As Long, Optional ByVal RespawnSite As Byte = 0)
    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    Dim i As Long
    
    'Respawn Site = 0 if normal fluctuation (warp if map boot is defined), = 1 if always warp to initial site, = 2: warp to army place
    ' Set HP to nothing
    Call SetPlayerVital(index, Vitals.HP, 0)
    
    PetDisband index, GetPlayerMap(index), True
    
    ' Warp player away
    Call SetPlayerDir(index, DIR_DOWN)
    
    SendSoundToMap GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), SoundEntity.seDie, GetPlayerClass(index)
    
    
    Dim mapnum As Long, X As Long, Y As Long
    GetOnDeathMap index, mapnum, X, Y, RespawnSite
    PlayerWarpByEvent index, mapnum, X, Y
    
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .caster = 0
            .StartTime = 0
        End With
    Next
    
    For i = 1 To PlayerActions_Count - 1
        Call UnblockPlayerAction(index, i)
    Next
    
    ' Clear spell casting
    TempPlayer(index).spellBuffer.Spell = 0
    TempPlayer(index).spellBuffer.Timer = 0
    TempPlayer(index).spellBuffer.Target = 0
    TempPlayer(index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(index)
    
    TempPlayer(index).InBank = False
    TempPlayer(index).InShop = 0
    If TempPlayer(index).InTrade > 0 Then
    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).Num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).Num = 0
        TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).Value = 0
    Next

    
    TempPlayer(TempPlayer(index).InTrade).InTrade = 0
    SendCloseTrade TempPlayer(index).InTrade
    
    'must be below
    TempPlayer(index).InTrade = 0
    SendCloseTrade index


End If
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

   

End Sub
Public Function PosOrdenation(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Integer
    If x1 < x2 Then
        PosOrdenation = 1
    ElseIf x1 > x2 Then
        PosOrdenation = -1
    Else
        If y1 < y2 Then
            PosOrdenation = 1
        ElseIf y1 > y2 Then
            PosOrdenation = -1
        Else
            PosOrdenation = 0
        End If
    End If
End Function
Public Function BinarySearchResource(ByVal mapnum As Long, ByVal left As Long, ByVal right As Long, ByVal X As Long, ByVal Y As Long) As Long
    If right < left Then
        BinarySearchResource = 0
    Else
        Dim meddle As Integer
        meddle = (left + right) \ 2
        
        With ResourceCache(mapnum).ResourceData(meddle)
        
        Dim Ordenation As Integer
        Ordenation = PosOrdenation(X, Y, .X, .Y)
        If Ordenation = 1 Then
            BinarySearchResource = BinarySearchResource(mapnum, left, meddle - 1, X, Y)
        ElseIf Ordenation = -1 Then
            BinarySearchResource = BinarySearchResource(mapnum, meddle + 1, right, X, Y)
        Else
            BinarySearchResource = meddle
        End If
        
        End With
    End If
        
        
End Function
 
Function CheckResource(ByVal index As Long, ByVal X As Long, ByVal Y As Long) As Boolean
        Dim Resource_Num As Long
        Dim Resource_index As Long
        Dim Rx As Long, Ry As Long
        Dim i As Long
        Dim Damage As Long
        Dim Reward_index As Byte
        
        If OutOfBoundries(X, Y, GetPlayerMap(index)) Then Exit Function
   
        If map(GetPlayerMap(index)).Tile(X, Y).Type <> TILE_TYPE_RESOURCE Then Exit Function
   
        Resource_Num = 0
        Resource_index = map(GetPlayerMap(index)).Tile(X, Y).Data1
        ' Get the cache number
        
        'Resource_Num = BinarySearchResource(GetPlayerMap(index), 1, ResourceCache(GetPlayerMap(index)).Resource_Count, X, Y)
        Resource_Num = GetMapRefResourceIndexByTile(GetMapRef(GetPlayerMap(index)), X, Y)
        
        
        If Resource_Num > 0 Then
   
                If Resource(Resource_index).ToolRequired > 0 Then
                        If GetPlayerEquipment(index, Weapon) > 0 Then
                                If item(GetPlayerEquipment(index, Weapon)).Data3 <> Resource(Resource_index).ToolRequired Then
                                        PlayerMsg index, "You don't have the right tool equpped.", BrightRed
                                        Exit Function
                                Else
                                        Damage = RAND(1, item(GetPlayerEquipment(index, Weapon)).Data2)
                                End If
                        Else
                                PlayerMsg index, "You need a tool to do that.", BrightRed
                                ' send the sound
                                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
                        
                                Exit Function
                        End If
                Else
                        If GetPlayerEquipment(index, Weapon) > 0 Then
                                Damage = RAND(1, item(GetPlayerEquipment(index, Weapon)).Data2 + GetPlayerStat(index, Stats.Agility))
                        Else
                                Damage = RAND(1, (GetPlayerStat(index, Stats.Strength) / 5))
                        End If
                End If
                   
                ' check if already cut down
                If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).ResourceState = 0 Then
                                   
                        Rx = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).X
                        Ry = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).Y
                                ' check if damage is more than health
                                If Damage > 0 Then
                                        ' cut it down!
                                        If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).cur_health - Damage <= 0 Then
                                                'SendActionMsg GetPlayerMap(Index), "-" & ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_Num).cur_health, BrightRed, 1, (Rx * 32), (Ry * 32)
                                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).ResourceState = 1 ' Cut
                                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).ResourceTimer = GetRealTickCount
                                                SendSingleResourceCacheToMap GetPlayerMap(index), Resource_Num


                                                Reward_index = CalculateResourceRewardindex(Resource_index)
                                                If Reward_index > 0 Then
                                                    Call CheckResourceReward(index, Rx, Ry, Resource_index, Reward_index)
                                                End If
                                                SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, Rx, Ry

                                        Else
                                                ' just do the damage
                                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).cur_health = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).cur_health - Damage
                                                SendActionMsg GetPlayerMap(index), "-" & Damage, BrightRed, 1, (Rx * 32), (Ry * 32)
                                                SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, Rx, Ry
                                        End If
                                        ' send the sound
                                        SendMapSound index, Rx, Ry, SoundEntity.seResource, Resource_index
                                        'ALATAR
                                        Call CheckTasks(index, QUEST_TYPE_GOTRAIN, Resource_index)
                                        '/ALATAR
                                        CheckResource = True
                                Else
                                        ' too weak
                                        SendActionMsg GetPlayerMap(index), "Failed!", BrightRed, 1, (Rx * 32), (Ry * 32)
                                End If
                        Else
                                ' send message if it exists
                                If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                                        SendActionMsg GetPlayerMap(index), Resource(Resource_index).EmptyMessage, BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                End If
                        End If
                End If
           
End Function

Sub CheckSingleResource(ByVal index As Long, ByVal Resource_Num As Long)
        Dim Resource_index As Long
        Dim mapnum As Long
        Dim Rx As Long, Ry As Long
        Dim Damage As Long
        Dim Reward_index As Byte
   
        mapnum = GetPlayerMap(index)
        Rx = GetPlayerX(index)
        Ry = GetPlayerY(index)
        
        If GetNextPositionByRef(GetPlayerDir(index), mapnum, Rx, Ry) Then Exit Sub
        
        If map(mapnum).Tile(Rx, Ry).Type <> TILE_TYPE_RESOURCE Then Exit Sub
   
        If Not (Resource_Num > 0 And Resource_Num <= ResourceCache(mapnum).Resource_Count) Then Exit Sub
        
        If Rx <> ResourceCache(mapnum).ResourceData(Resource_Num).X Or Ry <> ResourceCache(mapnum).ResourceData(Resource_Num).Y Then Exit Sub
        
        Resource_index = map(mapnum).Tile(Rx, Ry).Data1
        
        If Resource_index < 1 Or Resource_index > MAX_RESOURCES Then Exit Sub
        
        If ResourceCache(mapnum).ResourceData(Resource_Num).ResourceState <> 0 Then
            ' send message if it exists
            If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                    SendActionMsg mapnum, Resource(Resource_index).EmptyMessage, BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
            End If
            Exit Sub
        End If
        
        
   
        If Resource(Resource_index).ToolRequired > 0 Then
            If GetPlayerEquipment(index, Weapon) > 0 Then
                    If item(GetPlayerEquipment(index, Weapon)).Data3 <> Resource(Resource_index).ToolRequired Then
                        PlayerMsg index, "You don't have the right tool equpped.", BrightRed
                        Exit Sub
                    Else
                            Damage = RAND(1, item(GetPlayerEquipment(index, Weapon)).Data2)
                    End If
            Else
                    PlayerMsg index, "You need a tool to do that!", BrightRed
                    ' send the sound
                    SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
                    Exit Sub
            End If
        Else
            If GetPlayerEquipment(index, Weapon) > 0 Then
                Damage = RAND(1, item(GetPlayerEquipment(index, Weapon)).Data2 + GetPlayerStat(index, Stats.Agility))
            Else
                Damage = RAND(1, (GetPlayerStat(index, Stats.Strength) / 5))
            End If
        End If
        

        ' check if damage is more than health
        If Damage > 0 Then
            ' cut it down!
            If ResourceCache(mapnum).ResourceData(Resource_Num).cur_health - Damage <= 0 Then
                'SendActionMsg mapnum, "-" & ResourceCache(mapnum).ResourceData(Resource_Num).cur_health, BrightRed, 1, (Rx * 32), (Ry * 32)
                ResourceCache(mapnum).ResourceData(Resource_Num).ResourceState = 1 ' Cut
                ResourceCache(mapnum).ResourceData(Resource_Num).ResourceTimer = GetRealTickCount
                SendSingleResourceCacheToMap mapnum, Resource_Num
                      
                Reward_index = CalculateResourceRewardindex(Resource_index)
                If Reward_index > 0 Then
                    Call CheckResourceReward(index, Rx, Ry, Resource_index, Reward_index)
                End If
                SendAnimation mapnum, Resource(Resource_index).Animation, Rx, Ry
            Else
                ' just do the damage
                ResourceCache(mapnum).ResourceData(Resource_Num).cur_health = ResourceCache(mapnum).ResourceData(Resource_Num).cur_health - Damage
                SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (Rx * 32), (Ry * 32)
                SendAnimation mapnum, Resource(Resource_index).Animation, Rx, Ry
            End If
            
            ' send the sound
            SendMapSound index, Rx, Ry, SoundEntity.seResource, Resource_index
            'ALATAR
            Call CheckTasks(index, QUEST_TYPE_GOTRAIN, Resource_index)
            '/ALATAR
        Else
            ' too weak
            SendActionMsg mapnum, "Failed!", BrightRed, 1, (Rx * 32), (Ry * 32)
        End If
                          
End Sub

Public Sub CheckResourceReward(ByVal index As Long, ByVal Rx As Long, ByVal Ry As Long, ByRef ResourceNum As Long, ByVal ResourceReward As Byte)

Select Case Resource(ResourceNum).Rewards(ResourceReward).RewardType

Case REWARD_ITEM
    Dim RewardItem As Long
    RewardItem = Resource(ResourceNum).Rewards(ResourceReward).Reward
    If RewardItem < 1 Or RewardItem > MAX_ITEMS Then Exit Sub
    
    Dim i As Long
    Dim GivenValue As Long
    i = CanGiveItem(index, RewardItem, 1, GivenValue)
    If i > 0 Then
        GiveInvSlot index, i, RewardItem, GivenValue
        ' send message if it exists
        'If Resource(ResourceNum).ItemSuccessMessage Then
                'SendActionMsg GetPlayerMap(index), Trim$(Resource(ResourceNum).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        'Else
                SendActionMsg GetPlayerMap(index), Trim$(item(RewardItem).Name) & "!", BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        'End If
    End If
Case REWARD_SPAWN_NPC
    Dim npcnum As Long
    npcnum = Resource(ResourceNum).Rewards(ResourceReward).Reward
    
    If npcnum < 1 Or npcnum > MAX_NPCS Then Exit Sub
    
    'Spawn the npc and set the timer
    Dim j As Integer
    j = SpawnTempNPC(npcnum, GetPlayerMap(index), Rx, Ry)
    If j > 0 Then
        'tempnpc dissapears when killed, so can't respawn, we will use spawnwait at inversal prupose, dispawn the npc at certain time
        MapNpc(GetPlayerMap(index)).NPC(j).SpawnWait = GetRealTickCount + Resource(ResourceNum).RespawnTime * 1000
    End If

End Select
    
    
    
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Bank(index).item(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Bank(index).item(BankSlot).Num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Bank(index).item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal Itemvalue As Long)
    Bank(index).item(BankSlot).Value = Itemvalue
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal invSlot As Long, ByVal amount As Long)

Dim BankSlot
Dim Value

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, invSlot))
        
    If BankSlot > 0 Then
        If isItemStackable(GetPlayerInvItemNum(index, invSlot)) Then
            Value = CheckBankMoneyAdd(GetPlayerBankItemValue(index, BankSlot), amount)
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, Value)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, Value)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            End If
        End If
    End If
    
    'SaveBank index
    'SavePlayer index
    SendBank index

End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Long, ByVal amount As Long)

Dim invSlot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    If amount < 0 Or amount > GetPlayerBankItemValue(index, BankSlot) Then
        Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
    If invSlot > 0 Then
        If isItemStackable(GetPlayerBankItemNum(index, BankSlot)) Then
            'its so much money?
            If GetPlayerMaxMoney(index) < amount Then: PlayerMsg index, "You can't withdraw that amount!", BrightRed: Exit Sub
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), amount)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - amount)
                If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                    Call SetPlayerBankItemNum(index, BankSlot, 0)
                    Call SetPlayerBankItemValue(index, BankSlot, 0)
                End If
            Else
                If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                    Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                    Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
                Else
                    Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                    Call SetPlayerBankItemNum(index, BankSlot, 0)
                    Call SetPlayerBankItemValue(index, BankSlot, 0)
                End If
            End If
        End If
    SendBank index
End Sub
Public Sub KillPlayer(ByVal index As Long, Optional ByVal LoseExp As Byte = 0)
Dim exp As Long
    Call OnDeath(index)
End Sub

Public Sub UseItem(ByVal index As Long, ByVal invNum As Long)
Dim N As Long, i As Long, TempItem As Long, X As Long, Y As Long, ItemNum As Long, b As Long, j As Long
Dim ContainerAmount, amount As Long
    
    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_INV Then
        HackingAttempt index, "UseItem <1 or >Max."
        Exit Sub
    End If

    If IsActionBlocked(index, aUseItem) Then Exit Sub
    
    If (GetPlayerInvItemNum(index, invNum) > 0) And (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
        N = item(GetPlayerInvItemNum(index, invNum)).Data2
        ItemNum = GetPlayerInvItemNum(index, invNum)
        
        ' Find out what kind of item it is
        Select Case item(ItemNum).Type
            Case ITEM_TYPE_ARMOR
            
                If CanPlayerEquipItem(index, ItemNum) = False Then Exit Sub

                SwapInvEquipment index, invNum, Armor
                SendInventoryUpdate index, invNum
                SendEquipmentUpdate index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_WEAPON
            
                If CanPlayerEquipItem(index, ItemNum) = False Then Exit Sub
                
                b = 0

             
                If item(ItemNum).istwohander = True Then
                    If GetPlayerEquipment(index, Shield) > 0 Then
                        b = FindOpenInvSlot(index, GetPlayerEquipment(index, Shield))
                        If b > 0 Then
                            SwapInvEquipment index, b, Shield
                            SendInventoryUpdate index, b
                        Else
                            PlayerMsg index, "You don't have enough inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If
                End If
                
                SwapInvEquipment index, invNum, Weapon
                SendInventoryUpdate index, invNum
                SendEquipmentUpdate index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_HELMET
            
                If CanPlayerEquipItem(index, ItemNum) = False Then Exit Sub

                SwapInvEquipment index, invNum, helmet
                
                SendInventoryUpdate index, invNum
                SendEquipmentUpdate index
                              
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_SHIELD
            
                If CanPlayerEquipItem(index, ItemNum) = False Then Exit Sub

                SwapInvEquipment index, invNum, Shield
                
                If GetPlayerEquipment(index, Weapon) > 0 Then
                    If item(GetPlayerEquipment(index, Weapon)).istwohander = True Then
                        SwapInvEquipment index, invNum, Weapon 'The shield slot had to be empty beforc calling procedeture
                    End If
                End If
                
                SendInventoryUpdate index, invNum
                SendEquipmentUpdate index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            ' consumable
            Case ITEM_TYPE_CONSUME
            
                If TempPlayer(index).FreeAction = False Then Exit Sub
   
                If CanPlayerEquipItem(index, ItemNum) = False Then Exit Sub
                
                ' add hp
                If item(ItemNum).AddHP > 0 Then
                    player(index).vital(Vitals.HP) = player(index).vital(Vitals.HP) + item(ItemNum).AddHP
                    SendActionMsg GetPlayerMap(index), "+" & item(ItemNum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, HP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add mp
                If item(ItemNum).AddMP > 0 Then
                    player(index).vital(Vitals.MP) = player(index).vital(Vitals.MP) + item(ItemNum).AddMP
                    SendActionMsg GetPlayerMap(index), "+" & item(ItemNum).AddMP, Cyan, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, MP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add exp
                If item(ItemNum).AddEXP > 0 Then
                    SetPlayerExp index, GetPlayerExp(index) + item(ItemNum).AddEXP
                    CheckPlayerLevelUp index
                    SendActionMsg GetPlayerMap(index), "+" & item(ItemNum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendEXP index
                End If
                Call SendAnimation(GetPlayerMap(index), item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                Call TakeInvItem(index, player(index).Inv(invNum).Num, 1)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            
                If item(ItemNum).ConsumeItem <> 0 Then
                    GiveInvItem index, item(ItemNum).ConsumeItem, 1
                End If
            
                TempPlayer(index).FreeAction = False
            
            Case ITEM_TYPE_KEY
            
                If CanPlayerEquipItem(index, ItemNum) = False Then Exit Sub
                
                X = GetPlayerX(index)
                Y = GetPlayerY(index)
                If GetNextPositionByRef(GetPlayerDir(index), GetPlayerMap(index), X, Y) Then Exit Sub


                ' Check if a key exists
                If map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_KEY Then
                    Dim KeyToOpen As Long
                    KeyToOpen = GetTempDoorNumberByTile(GetPlayerMap(index), X, Y)
                    If KeyToOpen > 0 Then
                        ' Check if the key they are using matches the map key
                        If ItemNum = map(GetPlayerMap(index)).Tile(X, Y).Data1 Then
                            TempTile(GetPlayerMap(index)).Door(KeyToOpen).state = True
                            TempTile(GetPlayerMap(index)).Door(KeyToOpen).DoorTimer = GetRealTickCount + 60000
                            SendMapKeyToMap GetPlayerMap(index), X, Y, 1
                            Call MapMsg(GetPlayerMap(index), "The door has been opened.", White)
                            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSwitchFloor, 1
                            Call SendAnimation(GetPlayerMap(index), item(ItemNum).Animation, X, Y)
                            ' Check if we are supposed to take away the item
                            If map(GetPlayerMap(index)).Tile(X, Y).Data2 = 1 Then
                                Call TakeInvItem(index, ItemNum, 0)
                                Call PlayerMsg(index, "The key broke!", Yellow)
                            End If
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_SPELL
            
                If CanPlayerEquipItem(index, ItemNum) = False Then Exit Sub
                
                ' Get the spell num
                N = item(ItemNum).Data1

                If N > 0 Then

                    ' Make sure they are the right class
                    If Spell(N).ClassReq = GetPlayerClass(index) Or Spell(N).ClassReq = 0 Then
                        ' Make sure they are the right level
                        i = Spell(N).LevelReq

                        If i <= GetPlayerLevel(index) Then
                            i = FindOpenSpellSlot(index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(index, N) Then
                                    Call SetPlayerSpell(index, i, N)
                                    Call SendAnimation(GetPlayerMap(index), item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                                    Call TakeInvItem(index, ItemNum, 0)
                                    Call PlayerMsg(index, "You have learned a new skill: " & Trim$(Spell(N).Name) & ".", BrightGreen)
                                    Call SendPlayerSpells(index)
                                Else
                                    Call PlayerMsg(index, "You already know this ability.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(index, "You cannot learn any more skills", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(index, "You must be level " & i & " to learn this skill.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(index, "This skill can only be learned by " & CheckGrammar(GetClassName(Spell(N).ClassReq)) & ".", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
                
        Case ITEM_TYPE_RESET_POINTS
            If Not IsPetTargetted(index) Then
                i = ResetPlayerPoints(index)
                Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + i)
                Call SendPoints(index)
                Call ComputeAllPlayerStats(index)
                Call SendStats(index)
                Call PlayerMsg(index, "You have reset your points and now have " & i & " points.", BrightGreen)
            Else
                i = ResetPlayerPetPoints(index, GetPlayerPetSlot(index))
                Call SetPlayerPetPOINTS(index, GetPlayerPetPOINTS(index) + i)
                Call SendPetData(index, TempPlayer(index).TempPet.ActualPet)
                Call PlayerMsg(index, "You have reset your pet's points and now have " & i & " points!", BrightGreen)
            End If
            Call TakeInvItem(index, ItemNum, 0)
            ' send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum

        Case ITEM_TYPE_TRIFORCE
        
            'Triforce Type
            If Not GetPlayerLevel(index) >= MIN_LEVEL_TO_RESET Then
                PlayerMsg index, "You must be level " & MIN_LEVEL_TO_RESET & " to aquire the triforce", BrightRed
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
                Exit Sub
            End If
            
            SendOpenTriforce index
            
            Dim hasNumTri As Integer
            
            For i = 1 To 3
                If player(index).triforce(i) = True Then hasNumTri = hasNumTri + 1
            Next i
            
            If hasNumTri = 0 Then
                PlayerMsg index, "Be warned, " & GetPlayerName(index) & ".. you are not prepared for future..", White
                PlayerMsg index, "For the most powerful swords, much like Heros, must be reforged in fires of battle..", White
                PlayerMsg index, "Know that when you choose the Triforce, you will be reborn again..", White
            Else
                PlayerMsg index, GetPlayerName(index) & " you have chosen this path of your own will.", White
                PlayerMsg index, "Strong as it may seem, do you have the resolve to continue the Hero's path?", White
            End If
            
            ' send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum

        Case ITEM_TYPE_REDEMPTION
            
            If GetPlayerPK(index) = PK_PLAYER Then
                PlayerMsg index, "You've been redeemed!", BrightGreen
                Call PlayerRedemption(index)
                Call TakeInvItem(index, ItemNum, 0)
                Call SendJusticeToMap(index)
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            End If
        
        Case ITEM_TYPE_RESIGN
        
            If GetPlayerPK(index) = HERO_PLAYER Then
                PlayerMsg index, "You've resigned your Hero work.", BrightGreen
                Call PlayerRedemption(index)
                Call TakeInvItem(index, ItemNum, 0)
                Call SendJusticeToMap(index)
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            End If
            
         Case ITEM_TYPE_CONTAINER
         
                If CanPlayerEquipItem(index, ItemNum) = False Then Exit Sub
        
                PlayerMsg index, "You have opened " & Trim$(item(ItemNum).Name), Green
                TakeInvItem index, ItemNum, 0
                For i = 0 To MAX_ITEM_CONTAINERS
                    If item(ItemNum).Container(i).ItemNum > 0 And item(ItemNum).Container(i).ItemNum <= MAX_ITEMS Then
                        'Award item
                        If (Not isItemStackable(item(ItemNum).Container(i).ItemNum)) Then
                            amount = 0
                        Else
                            amount = item(ItemNum).Container(i).Value
                            
                        End If
                        Call GiveInvItem(index, item(ItemNum).Container(i).ItemNum, amount)
                        If (amount > 0) Then
                            PlayerMsg index, "You've discovered " & Trim$(item(item(ItemNum).Container(i).ItemNum).Name) & " (" & amount & ")", Green
                        Else
                            PlayerMsg index, "You've discovered " & Trim$(item(item(ItemNum).Container(i).ItemNum).Name), Green
                        End If
                    End If
                Next i
        Case ITEM_TYPE_BAG
        
            If Not CanPlayerEquipItem(index, ItemNum) Then Exit Sub
            
            If GetPlayerBags(index) = MAX_RUPEE_BAGS Then
                PlayerMsg index, "You've already got the max number of rupees possible!", BrightRed
                Exit Sub
            End If
            
            Dim Bags As Byte
            Bags = GetPlayerBags(index) + item(ItemNum).AddBags
            
            TakeInvItem index, ItemNum, 0
            If Bags >= MAX_RUPEE_BAGS Then
                Bags = MAX_RUPEE_BAGS
                Call SetPlayerBags(index, Bags)
                PlayerMsg index, "Now you are at max capacity of rupees: " & GetPlayerMaxMoney(index) & "!", BrightGreen
            Else
                Call SetPlayerBags(index, Bags)
                PlayerMsg index, "You've increased your maximum rupee capacity to " & GetPlayerMaxMoney(index) & " rupees!", BrightGreen
            End If
            
        Case ITEM_TYPE_ADDWEIGHT
        
            If Not CanPlayerEquipItem(index, ItemNum) Then Exit Sub
                
            If GetPlayerMaxWeight(index) = 200000 Then
                Exit Sub
            End If
            
            Dim AddWeight As Long
            AddWeight = item(ItemNum).Data1
            
            If AddWeight < 0 Or AddWeight > 10000 Then Exit Sub
            
            TakeInvItem index, ItemNum, 0
            Call SetPlayerMaxWeight(index, GetPlayerMaxWeight(index) + AddWeight)
            SendMaxWeight index
            PlayerMsg index, "You've increased your maximum weight to: " & GetPlayerMaxWeight(index), BrightGreen
            
        End Select
                        
                
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
    End If
End Sub








Public Function ResetPlayerPoints(ByVal index As Long) As Long
Dim i As Byte, sum As Long
ResetPlayerPoints = 0
'PlayerUnequip (index)
sum = 0

For i = 1 To Stats.Stat_Count - 1
    Do While player(index).stat(i) > Class(GetPlayerClass(index)).stat(i)
        player(index).stat(i) = player(index).stat(i) - 1
        sum = sum + 1
    Loop
Next

ResetPlayerPoints = sum
    
End Function

Public Sub PlayerPVPDrops(ByVal index As Long)
Dim i As Long
Dim ItemNum As Long
Dim Itemvalue As Long

For i = 1 To MAX_INV
 
If GetPlayerInvItemNum(index, i) > 0 Then
    If IsItemDroppable(GetPlayerInvItemNum(index, i), index) Then 'check if dropable
        If isItemStackable(GetPlayerInvItemNum(index, i)) Then
            Itemvalue = GetPlayerInvItemValue(index, i)
            If Itemvalue > 0 Then
                'Drop 1 at least
                Itemvalue = Itemvalue * (1 / 10)
                If Itemvalue = 0 Then Itemvalue = 1
            
                Call PlayerMapDropItem(index, i, Itemvalue, False)
            End If
        Else
            Call PlayerMapDropItem(index, i, 1, False)
        End If
    End If
End If
Next
    
For i = 1 To Equipment.Equipment_Count - 1
    If GetPlayerEquipment(index, i) > 0 Then
        If IsItemDroppable(GetPlayerEquipment(index, i), index) Then
            Call PlayerUnequipItemAndDrop(index, i)
        End If
    End If
Next

'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & "he dies and his items fall to the ground!", Yellow)

End Sub


Public Sub ResetPlayer(ByVal index As Long)
    Dim i As Long
    
    'pk
    player(index).PK = NO
    'lvl
    player(index).level = 1
    
    'points
    player(index).points = 0
    
    'exp
    player(index).exp = 0
    SendEXP (index)
    
    
    'inventory
    For i = 1 To MAX_INV
        player(index).Inv(i).Num = 0
        player(index).Inv(i).Value = 0
    Next
    Call SendInventory(index)
    
    'Equipment
    For i = 1 To Equipment.Equipment_Count - 1
        player(index).Equipment(i) = 0
    Next
    SendWornEquipment index
    SendMapEquipment index
    
    'Quests
    For i = 1 To MAX_QUESTS
        player(index).PlayerQuest(i).Status = 0
        player(index).PlayerQuest(i).ActualTask = 0
        player(index).PlayerQuest(i).CurrentCount = 0
    Next
    Call SendPlayerQuests(index)
    
    'Spells
    For i = 1 To MAX_PLAYER_SPELLS
        player(index).Spell(i) = 0
    Next
    Call SendPlayerSpells(index)
    

    player(index).NPCKills = 0
    
    'hotbars
    For i = 1 To MAX_HOTBAR
        player(index).Hotbar(i).slot = 0
        player(index).Hotbar(i).sType = 0
    Next
    Call SendHotbar(index)
    
    'stats
    For i = 1 To Stats.Stat_Count - 1
        player(index).stat(i) = Class(GetPlayerClass(index)).stat(i)
    Next
    
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    
    Call ClearBank(index)
    Call SaveBank(index)
    
    Call SetPlayerBags(index, 1)
        
    Call ComputeAllPlayerStats(index)
    Call SendStats(index)
    
    Call SendPlayerData(index)

End Sub

Public Sub ComputePlayerReset(ByVal index As Long, ByVal triforce As TriforceType)
    Dim colour As Byte
    Dim message As String
    Dim i As Byte
    Dim found As Boolean
    
    If Not IsPlaying(index) Then Exit Sub
    
    If Not GetPlayerLevel(index) >= 80 Then
        PlayerMsg index, "You must be lvl 80 like m minimum", BrightRed
        Exit Sub
    End If
    If GetPlayerTriforcesNum(index) > 0 Then
        PlayerMsg index, "Ya has renacido", BrightRed
        Exit Sub
    End If
    If GetPlayerTriforce(index, triforce) = True Then
        PlayerMsg index, "You already have that acquired Triforce", BrightRed
        Exit Sub
    End If
    
    found = False
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) > 0 Then
            If item(player(index).Inv(i).Num).Type = ITEM_TYPE_TRIFORCE Then
                found = True
                player(index).Inv(i).Num = 0
                player(index).Inv(i).Value = 0
                Call SendInventoryUpdate(index, i)
                Exit For
            End If
        End If
    Next
    
    If Not found Then
        PlayerMsg index, "You don't have the triforce!", BrightRed
        Exit Sub
    End If
    
    Call ResetPlayer(index)
    player(index).triforce(triforce) = True
    
    Select Case triforce
    Case TRIFORCE_COURAGE
        message = "of Valor"
        colour = BrightGreen
    Case TRIFORCE_WISDOM
        message = "of Wisdom"
        colour = Cyan
    Case TRIFORCE_POWER
        message = "of Power"
        colour = BrightRed
    End Select
    
    
    For i = 1 To TriforceType.TriforceType_Count - 1
        If GetPlayerTriforce(index, i) = True Then
        Select Case i
            Case TRIFORCE_COURAGE
                SetPlayerStat index, Agility, GetPlayerStat(index, Agility) + 15
                SetPlayerStat index, Endurance, GetPlayerStat(index, Endurance) + 15
            Case TRIFORCE_WISDOM
                SetPlayerStat index, Intelligence, GetPlayerStat(index, Intelligence) + 15
                SetPlayerStat index, willpower, GetPlayerStat(index, willpower) + 15
            Case TRIFORCE_POWER
                SetPlayerStat index, Intelligence, GetPlayerStat(index, Intelligence) + 15
                SetPlayerStat index, Strength, GetPlayerStat(index, Strength) + 15
       End Select
       End If
    Next
    
    Call SendPlayerData(index)
    
    PlayerMsg index, "You feel a strange surge of power coursing through you.", BrightBlue
    
    GlobalMsg GetPlayerName(index) & " has acquired the triforce " & message, colour, False
    
    ForwardGlobalMsg "[Hub - " & SERVER_NAME & "] " & GetPlayerName(index) & " has acquired the triforce " & message
    
End Sub
Public Function GetPlayerTriforcesNum(ByVal index As Long) As Byte
Dim i As Byte
GetPlayerTriforcesNum = 0

For i = 1 To TriforceType.TriforceType_Count - 1
    If GetPlayerTriforce(index, i) = True Then
        GetPlayerTriforcesNum = GetPlayerTriforcesNum + 1
    End If
Next

End Function

Public Function GetPlayerTriforce(ByVal index As Long, ByVal triforce As TriforceType) As Boolean
Dim i As Byte
If Not IsPlaying(index) Then Exit Function

GetPlayerTriforce = False

If triforce > 0 And triforce < TriforceType_Count Then
    GetPlayerTriforce = player(index).triforce(triforce)
End If

End Function

Public Function HasPlayerAnyTriforce(ByVal index As Long) As Boolean
HasPlayerAnyTriforce = False
Dim i As Byte

For i = 1 To TriforceType.TriforceType_Count - 1
    If GetPlayerTriforce(index, i) = True Then
        HasPlayerAnyTriforce = True
        Exit Function
    End If
Next
End Function

Public Function CanPlayerEquipItem(ByVal index As Long, ByVal ItemNum As Long) As Boolean
Dim i As Byte

CanPlayerEquipItem = False

If Not (ItemNum > 0 And ItemNum <= MAX_ITEMS) Then Exit Function
' stat requirements
For i = 1 To Stats.Stat_Count - 1
    If GetPlayerRawStat(index, i) < item(ItemNum).Stat_Req(i) Then
        PlayerMsg index, "Your stats aren't good enough to equip this!", BrightRed
        'playsound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
        Exit Function
    End If
Next
                
' level requirement
If GetPlayerLevel(index) < item(ItemNum).LevelReq Then
    PlayerMsg index, "You aren't high enough level to equip this!", BrightRed
    'playsound
    SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
Exit Function
End If
                
' class requirement
If item(ItemNum).ClassReq > 0 Then
    If Not GetPlayerClass(index) = item(ItemNum).ClassReq Then
        PlayerMsg index, "You're the wrong class for this item!", BrightRed
        'playsound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
        Exit Function
    End If
End If
                
' access requirement
If Not GetPlayerAccess_Mode(index) >= item(ItemNum).AccessReq Then
    PlayerMsg index, "Your admin level isn't high enough for this item.", BrightRed
    'playsound
    SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
    Exit Function
End If

'Triforce Requeriment
If item(ItemNum).BindType > 1 And item(ItemNum).BindType < 5 Then
    If player(index).triforce(item(ItemNum).BindType - 1) = False Then
        PlayerMsg index, "You don't have right Triforce to equip this item.", BrightRed
        'playsound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
        Exit Function
    End If
ElseIf item(ItemNum).BindType = 5 Then
    If HasPlayerAnyTriforce(index) = False Then
        PlayerMsg index, "You must have a Triforce to equip this item.", BrightRed
        'playsound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
        Exit Function
    End If
End If

If item(ItemNum).ArmyType_Req <> NONE_PLAYER Then
    If GetPlayerPK(index) <> item(ItemNum).ArmyType_Req Then
        PlayerMsg index, "You do not belong to the army required to equip this item.", BrightRed
        'playsound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
        Exit Function
    Else
        If item(ItemNum).ArmyRange_Req > 0 Then
            If GetPlayerArmyRange(index) < item(ItemNum).ArmyRange_Req Then
                PlayerMsg index, "Your army rank is not high enouh to equip this.", BrightRed
                'playsound
                 SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
                Exit Function
            End If
        End If
    End If

End If

CanPlayerEquipItem = True


End Function



Public Function CheckSafeMode(ByVal attacker As Long, ByVal victim As Long) As Boolean
    'True: Player can't attack cause his safemode
    'False: Player can attack, if safe mode then victim is PK
    If IsPlayerNeutral(victim) Then
        If GetPlayerSafeMode(attacker) = True Then
            CheckSafeMode = True
        Else
            CheckSafeMode = False
        End If
    Else
        CheckSafeMode = False
    End If

End Function

Public Function GetPlayerSafeMode(ByVal index As Long) As Boolean
    GetPlayerSafeMode = player(index).SafeMode

End Function

Public Function GetPlayerNameColorByTriforce(ByVal index As Long) As Long

Dim color As Byte
Dim i As Byte

i = GetPlayerTriforcesNum(index)

'Normal Color
If i = 0 Then
    GetPlayerNameColorByTriforce = BrightGreen
    Exit Function
Else
    If GetPlayerTriforce(index, TRIFORCE_WISDOM) Then
        color = Cyan
    End If
    If GetPlayerTriforce(index, TRIFORCE_COURAGE) Then
        color = Green
    End If
    If GetPlayerTriforce(index, TRIFORCE_POWER) Then
        color = Red
    End If
End If


GetPlayerNameColorByTriforce = color

End Function

Public Function GetPlayerTriforcesName(ByVal index As Long) As String
Dim Chain As String
Dim i As Byte
Dim j As Byte
i = GetPlayerTriforcesNum(index)
Chain = vbNullString
If i = 0 Then
    Chain = vbNullString
Else
    For j = 1 To TriforceType.TriforceType_Count - 1
        If GetPlayerTriforce(index, j) = True Then
        Select Case j
            Case TriforceType.TRIFORCE_COURAGE
                Chain = Chain & "<Valor>"
            Case TriforceType.TRIFORCE_WISDOM
                Chain = Chain & "<Wisdom>"
            Case TriforceType.TRIFORCE_POWER
                Chain = Chain & "<Power>"
        End Select
        End If
    
    Next
End If

GetPlayerTriforcesName = Chain

End Function

Public Function GetPlayerMaxMoney(ByVal index As Long) As Long
    GetPlayerMaxMoney = GetMaxMoneyByBag(GetPlayerBags(index))
End Function

Public Function GetPlayerBags(ByVal index As Long) As Byte
    GetPlayerBags = player(index).RupeeBags
End Function

Sub SetPlayerBags(ByVal index As Long, ByVal Bags As Byte)
    If (Bags <= MAX_RUPEE_BAGS) Then
        player(index).RupeeBags = Bags
        SendBags index, Bags
    End If
End Sub

Public Function CheckMoneyAdd(ByVal index As Long, ByVal initialvalue As Long, ByVal addvalue As Long) As Long
CheckMoneyAdd = initialvalue + addvalue
Dim MaxMoney As Long
MaxMoney = GetPlayerMaxMoney(index)

If CheckMoneyAdd > MaxMoney Then CheckMoneyAdd = MaxMoney

End Function

Public Function CheckBankMoneyAdd(ByVal initialvalue As Long, ByVal addvalue As Long) As Long
CheckBankMoneyAdd = initialvalue + addvalue
If (CheckBankMoneyAdd > MAX_BANK_RUPEES) Then
    CheckBankMoneyAdd = MAX_BANK_RUPEES
End If
End Function

Public Function GetMaxMoneyByBag(ByVal Bags As Byte) As Long
    If (Bags >= MAX_RUPEE_BAGS) Then
        GetMaxMoneyByBag = Bags * BAG_CAPACITY - 1
    Else
        GetMaxMoneyByBag = Bags * BAG_CAPACITY
    End If
End Function

Public Function SetPlayerCustomSprite(ByVal index As Long, ByVal CustomSprite As Byte)
    If CustomSprite > MAX_CUSTOM_SPRITES Then Exit Function
    player(index).CustomSprite = CustomSprite
End Function

Public Function GetPlayerCustomSprite(ByVal index As Long) As Byte
    If player(index).CustomSprite > MAX_CUSTOM_SPRITES Then Exit Function
    GetPlayerCustomSprite = player(index).CustomSprite
End Function


Public Sub SendEquipmentUpdate(ByVal index As Long)
    Call SendWornEquipment(index)
    Call SendMapEquipment(index)
    Call SendStats(index)
                 
    ' send vitals
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
End Sub




Sub ResetPlayerInactivity(ByVal index As Long)
    TempPlayer(index).InactiveTime = 0
End Sub

Function GetInactiveTime(ByVal index As Long) As Long
    GetInactiveTime = TempPlayer(index).InactiveTime
End Function

Sub WarpXtoY(ByVal X As Long, ByVal Y As Long, ByVal carry As Boolean)
    If X = Y Then Exit Sub
    If Not IsPlaying(X) Or Not IsPlaying(Y) Then Exit Sub
    
    Call PlayerWarpByEvent(X, GetPlayerMap(Y), GetPlayerX(Y), GetPlayerY(Y))
    If carry Then
        Call AddLog(Y, GetPlayerName(Y) & " has warped " & GetPlayerName(X) & " to self, map #" & GetPlayerMap(Y) & ".", ADMIN_LOG)
        If GetPlayerVisible(Y) = 0 Then
            Call PlayerMsg(X, "You have been teleported by " & GetPlayerName(Y) & ".", Cyan)
            Call PlayerMsg(Y, GetPlayerName(X) & " has been teleported by", Cyan)
        End If
    Else
        Call AddLog(X, GetPlayerName(X) & " has warped to " & GetPlayerName(Y) & ", map #" & GetPlayerMap(Y) & ".", ADMIN_LOG)
        If GetPlayerVisible(X) = 0 Then
            Call PlayerMsg(Y, GetPlayerName(X) & " has teleported to you.", Cyan)
            Call PlayerMsg(X, "You have been teleported to " & GetPlayerName(Y) & ".", Cyan)
        End If
    End If
End Sub

Sub BlockPlayerAction(ByVal index As Long, ByVal PlayerAction As PlayerActionsType, ByVal Time As Single)
    If index < 1 Or PlayerAction < 1 Or PlayerAction >= PlayerActions_Count Then Exit Sub
    
    TempPlayer(index).BlockedActions(PlayerAction).Value = True
    TempPlayer(index).BlockedActions(PlayerAction).Timer = GetRealTickCount + Time * 1000
    
    SendBlockedAction index, PlayerAction
End Sub

Function IsActionBlocked(ByVal index As Long, ByVal PlayerAction As PlayerActionsType) As Boolean
    If index < 1 Or PlayerAction < 1 Or PlayerAction >= PlayerActions_Count Then Exit Function
    IsActionBlocked = TempPlayer(index).BlockedActions(PlayerAction).Value
End Function

Sub UnblockPlayerAction(ByVal index As Long, ByVal PlayerAction As PlayerActionsType)
    If index < 1 Or PlayerAction < 1 Or PlayerAction >= PlayerActions_Count Then Exit Sub
    
    TempPlayer(index).BlockedActions(PlayerAction).Value = False
    TempPlayer(index).BlockedActions(PlayerAction).Timer = 0
    
    SendBlockedAction index, PlayerAction
End Sub

Sub UnblockAllPlayerActions(ByVal index As Long)
    If index = 0 Then Exit Sub
    Dim i As Long
    For i = 1 To PlayerActions_Count - 1
        If IsActionBlocked(index, i) Then
            UnblockPlayerAction index, i
        End If
    Next
End Sub

Sub CheckPlayerActions(ByVal index As Long, ByVal Tick As Long)
    Dim i As Byte
    For i = 1 To PlayerActions_Count - 1
        If TempPlayer(index).BlockedActions(i).Value = True Then
            If TempPlayer(index).BlockedActions(i).Timer < Tick Then
                UnblockPlayerAction index, i
            End If
        End If
    Next
End Sub

Sub ProtectPlayerAction(ByVal index As Long, ByVal PlayerAction As PlayerActionsType, ByVal Time As Single)
    If index < 1 Or PlayerAction < 1 Or PlayerAction >= PlayerActions_Count Then Exit Sub
    
    TempPlayer(index).ProtectedActions(PlayerAction).Value = True
    TempPlayer(index).ProtectedActions(PlayerAction).Timer = GetRealTickCount + Time * 1000
    
End Sub

Function IsActionProtected(ByVal index As Long, ByVal PlayerAction As PlayerActionsType) As Boolean
    If index < 1 Or Not (0 < PlayerAction < PlayerActions_Count) Then Exit Function
    
    IsActionProtected = TempPlayer(index).ProtectedActions(PlayerAction).Value
End Function

Sub ResetPlayerProtection(ByVal index As Long, ByVal PlayerAction As PlayerActionsType)
    If index < 1 Or PlayerAction < 1 Or PlayerAction >= PlayerActions_Count Then Exit Sub
    
    TempPlayer(index).ProtectedActions(PlayerAction).Value = False
    TempPlayer(index).ProtectedActions(PlayerAction).Timer = 0
    
End Sub


Sub CheckPlayerProtections(ByVal index As Long, ByVal Tick As Long)
    Dim i As Byte
    For i = 1 To PlayerActions_Count - 1
        If TempPlayer(index).ProtectedActions(i).Value Then
            If TempPlayer(index).ProtectedActions(i).Timer < Tick Then
                ResetPlayerProtection index, i
            End If
        End If
    Next
End Sub

Sub CheckPlayerActionsProtections(ByVal index As Long)
    Dim i As Byte
    For i = 1 To PlayerActions_Count - 1
        If IsActionBlocked(index, i) Then
            If IsActionProtected(index, i) Then
                UnblockPlayerAction index, i
            End If
        End If
    Next
End Sub

Sub KickPlayer(ByVal index As Long, Optional ByRef Reason As String = "")
    If index = 0 Or Not IsPlaying(index) Then Exit Sub
    
    Call GlobalMsg(GetPlayerName(index) & " has been kicked for: " & Reason, White, False)
    ForwardGlobalMsg "[Hub - " & SERVER_NAME & "] " & GetPlayerName(index) & " has been kicked for: " & Reason
    Call AddLog(0, GetPlayerName(index) & " has been kicked for: " & Reason, ADMIN_LOG)
    Call AlertMsg(index, " has been kicked for: " & Reason)
End Sub

Sub ClearPlayerTarget(ByVal index As Long)
    TempPlayer(index).Target = 0
    TempPlayer(index).TargetType = TARGET_TYPE_NONE
    SendTarget index
End Sub

Sub EarthQuake(ByVal index As Long)
    Dim a As Variant
    For Each a In GetMapPlayerCollection(GetPlayerMap(index))
        If a <> index Then
            If IsinRange(4, GetPlayerX(index), GetPlayerY(index), GetPlayerX(a), GetPlayerY(a)) Then
                Call PlayerAttackPlayer(index, a, GetPlayerDamageAgainstPlayer(index, a))
            End If
        End If
    Next
End Sub

Sub CheckGodAttack(ByVal index As Long)
    If GPE(index) Then
        UnblockAllPlayerActions index
        
        EarthQuake index
    End If
    
End Sub



Sub ComputePlayerAttackTimer(ByVal index As Long)
    SetPlayerAttackTimer index, GetRealTickCount
End Sub

Function CanPlayerAttackTimer(ByVal index As Long) As Boolean
    Dim Timer As Long, ItemNum As Long
    Timer = GetPlayerAttackTimer(index)
    ItemNum = GetPlayerEquipment(index, Weapon)
    If ItemNum > 0 Then
        If GetRealTickCount > Timer + GetItemSpeed(ItemNum) Then
            CanPlayerAttackTimer = True
        End If
    Else
        If GetRealTickCount > Timer + 1000 Then
            CanPlayerAttackTimer = True
        End If
    End If
End Function

Function GetPlayerAttackTimer(ByVal index As Long) As Long
    GetPlayerAttackTimer = TempPlayer(index).AttackTimer
End Function

Sub SetPlayerAttackTimer(ByVal index As Long, ByVal Time As Long)
    TempPlayer(index).AttackTimer = Time
End Sub





