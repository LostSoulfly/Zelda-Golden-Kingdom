Attribute VB_Name = "modMoveSystem"
Option Explicit

Private RunningSprites() As clsPair

Public Const RUNNING_SPRITES_PATH As String = "/Data/Scripts"
Private Const RUNNING_SPRITES_FILE As String = "/RunningSprites.ini"

Sub ReadRunningSprites()
    Dim NSprites As Long
    
    
    'ReadNSprites
    NSprites = ReadNumberOfRunningSprites
    If NSprites = 0 Then Exit Sub
    
    ReDim RunningSprites(1 To NSprites)
    Dim i As Long
    For i = 1 To NSprites
        Dim WalkSprite As Long, RunSprite As Long
        'read walk 'n run
        WalkSprite = ReadWalkSprite(i)
        RunSprite = ReadRunSprite(i)
        Set RunningSprites(i) = New clsPair
        With RunningSprites(i)
        .SetFirst WalkSprite
        .SetSecond RunSprite
        End With
    Next
End Sub


Function ReadNumberOfRunningSprites() As Long
    Dim s As String
    s = GetVar(App.Path & RUNNING_SPRITES_PATH & RUNNING_SPRITES_FILE, "Total", "Total")
    If IsNumeric(s) Then
        ReadNumberOfRunningSprites = CLng(s)
    End If
End Function

Function ReadRunSprite(ByVal header As Long) As Long
    Dim s As String
    s = GetVar(App.Path & RUNNING_SPRITES_PATH & RUNNING_SPRITES_FILE, CStr(header), "Run")
    If IsNumeric(s) Then
        ReadRunSprite = CLng(s)
    End If
End Function

Function ReadWalkSprite(ByVal header As Long) As Long
    Dim s As String
    s = GetVar(App.Path & RUNNING_SPRITES_PATH & RUNNING_SPRITES_FILE, CStr(header), "Walk")
    If IsNumeric(s) Then
        ReadWalkSprite = CLng(s)
    End If
End Function

Sub SendRunningSprites(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SRunningSprites
    Buffer.WriteLong UBound(RunningSprites)
    Dim i As Long
    For i = 1 To UBound(RunningSprites)
        With RunningSprites(i)
        Buffer.WriteLong .GetFirst
        Buffer.WriteLong .GetSecond
        End With
    Next
    
    SendDataTo index, Buffer.ToArray
    Set Buffer = Nothing
End Sub


Function GetTileType(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long) As Byte
    GetTileType = map(mapnum).Tile(X, Y).Type
End Function

Sub WarpPlayerByTile(ByVal index As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    With map(mapnum).Tile(X, Y)
    PlayerWarpByEvent index, .Data1, .Data2, .Data3
    End With
End Sub

Sub CheckTilePlayerMove(ByVal index As Long, ByVal X As Integer, ByVal Y As Integer, ByRef Moved As Boolean, ByRef Teleported As Boolean)

    Dim mapnum As Long
    mapnum = GetPlayerMap(index)
    If OutOfBoundries(X, Y, mapnum) Then
        Exit Sub
    End If
     
    Teleported = False
    Moved = True

    Dim TileType As Long, VitalType As Long, colour As Long, amount As Long, scriptnum As Long
    
    With map(mapnum).Tile(X, Y)
    
    Select Case GetTileType(mapnum, X, Y)
    
    Case TILE_TYPE_WALKABLE
    Case TILE_TYPE_BLOCKED
        Moved = False
    Case TILE_TYPE_RESOURCE
        Moved = isWalkableResource(mapnum, X, Y)
    Case TILE_TYPE_WARP
        mapnum = .Data1
        X = .Data2
        Y = .Data3
        Call PlayerWarpByEvent(index, mapnum, X, Y)
        Teleported = True
    Case TILE_TYPE_DOOR
        Dim TempDoorNum As Long
        Dim DoorNum As Long
        DoorNum = .Data1
        TempDoorNum = GetTempDoorNumberByDoorNum(mapnum, DoorNum)
        If TempDoorNum > 0 Then
            If Not IsTempDoorWalkable(mapnum, TempDoorNum) Then
                Moved = False
            End If
        End If
        
        If IsDoorOpened(mapnum, TempDoorNum) Then
            mapnum = Doors(DoorNum).WarpMap
            If mapnum > 0 Then
                X = Doors(DoorNum).WarpX
                Y = Doors(DoorNum).WarpY
                Call PlayerWarpByEvent(index, mapnum, X, Y)
                Teleported = True
            End If
        Else
            If GetDoorType(DoorNum) = DOOR_TYPE_WEIGHTSWITCH Then
                Call CheckWeightSwitch(mapnum, TempDoorNum)
            Else
                Moved = False
            End If
        End If

        
    Case TILE_TYPE_KEYOPEN
            X = .Data1
            Y = .Data2
            Dim KeyToOpen As Long
            KeyToOpen = GetTempDoorNumberByTile(mapnum, X, Y)
            If KeyToOpen > 0 Then
                If map(mapnum).Tile(X, Y).Type = TILE_TYPE_KEY And Not TempTile(GetPlayerMap(index)).Door(KeyToOpen).state Then
                    TempTile(mapnum).Door(KeyToOpen).state = True
                    TempTile(mapnum).Door(KeyToOpen).DoorTimer = GetRealTickCount + 60000
                    'Send to all players on the map
                    SendMapKeyToMap mapnum, X, Y, 1
                    'SendMapKey index, X, Y, 1
                    Call MapMsg(mapnum, "La puerta se ha abierto.", White)
                    SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSwitchFloor, 1
                End If
            End If
            
    Case TILE_TYPE_SHOP
    
        X = .Data1
        If X > 0 Then ' shop exists?
            If Len(Trim$(Shop(X).Name)) > 0 Then ' name exists?
                SendOpenShop index, X
                TempPlayer(index).InShop = X ' stops movement and the like
            End If
        End If
    Case TILE_TYPE_HEAL
    
        VitalType = .Data1
        amount = .Data2
        If GetPlayerVital(index, VitalType) < GetPlayerMaxVital(index, VitalType) Then
            If VitalType = Vitals.HP Then
                colour = BrightGreen
            Else
                colour = Cyan
            End If
            SendActionMsg mapnum, "+" & amount, colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
            SetPlayerVital index, VitalType, GetPlayerVital(index, VitalType) + amount
            'PlayerMsg index, "Sientes unas fuerzas que rejuvenecen tu cuerpo.", BrightGreen
            Call SendVital(index, VitalType)
            ' send vitals to party if in one
            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        End If
    Case TILE_TYPE_TRAP
    
        amount = .Data1
        SendActionMsg GetPlayerMap(index), "-" & amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
        If GetPlayerVital(index, HP) - amount <= 0 Then
            KillPlayer index
            PlayerMsg index, "Has Muerto.", BrightRed
            Teleported = True
            'Kill Counter
            player(index).EnviroDead = player(index).EnviroDead + 1
        Else
            SetPlayerVital index, HP, GetPlayerVital(index, HP) - amount
            PlayerMsg index, "Has sido dañado.", BrightRed
            Call SendVital(index, HP)
            ' send vitals to party if in one
            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        End If
    Case TILE_TYPE_BANK
        SendBank index
        TempPlayer(index).InBank = True
    Case TILE_TYPE_ICE
        If (Not player(index).onIce) Then
            'player is now on ice, lets send him to the ice
            player(index).onIce = True
            Call SendOnIce(index, player(index).onIce)
            player(index).IceDir = player(index).dir
            Call SendIceDir(index, player(index).dir)
        End If
    Case Else
    End Select
    
    If player(index).onIce And .Type <> TILE_TYPE_ICE Then
        'stop sliding
        player(index).onIce = False
        Call SendOnIce(index, player(index).onIce)
    End If
           
    End With
End Sub
Sub CheckTileEventsBeforeMoving(ByVal index As Long, ByVal X As Long, ByVal Y As Long)
    If OutOfBoundries(X, Y, GetPlayerMap(index)) Then Exit Sub
    With map(GetPlayerMap(index)).Tile(X, Y)
    Select Case .Type
    Case TILE_TYPE_SCRIPT
        Call ScriptTileLeave(index, .Data1)
    End Select
    End With
        
End Sub

Sub CheckTileEventsAfterMoving(ByVal index As Long, ByVal X As Integer, ByVal Y As Integer)
    If OutOfBoundries(X, Y, GetPlayerMap(index)) Then Exit Sub
    With map(GetPlayerMap(index)).Tile(X, Y)
    Select Case .Type
    Case TILE_TYPE_SLIDE
        ForcePlayerMove index, MOVING_WALKING, .Data1
    Case TILE_TYPE_SCRIPT
        Dim scriptnum As Long
        scriptnum = .Data1
        Call ScriptTilePresses(index, scriptnum)
    End Select
    End With
End Sub


Sub PlayerMove(ByVal index As Long, ByVal dir As Byte, ByVal Movement As Long, Optional ByVal sendToSelf As Boolean = False)
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or dir < DIR_UP Or dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    'If IsPlayerOverWeight(index) Then
        'SendPlayerXY (index)
        'Exit Sub
    'End If
    
    
    Dim Moved As Boolean
    Dim mapnum As Long
    mapnum = GetPlayerMap(index)
    
    Dim CurX As Long, CurY As Long
    CurX = GetPlayerX(index)
    CurY = GetPlayerY(index)
    
    
    
    If (Not OutOfBoundries(CurX, CurY, mapnum)) Then
        If isDirBlocked(map(mapnum).Tile(CurX, CurY).DirBlock, dir + 1) Then
            SendPlayerXY (index)
            Exit Sub
        End If
    End If
    
    Call SetPlayerDir(index, dir)

    Dim X As Long, Y As Long

    GetNextPosition CurX, CurY, dir, X, Y
    
    If (Not OutOfBoundries(X, Y, mapnum)) Then
        Dim Teleported As Boolean
        Call CheckTilePlayerMove(index, X, Y, Moved, Teleported)
        If Not Teleported Then
            If Not Moved Then
                Call SendPlayerXY(index)
            Else
                Call SetPlayerX(index, X)
                Call SetPlayerY(index, Y)
                Call SendPlayerMove(index, Movement, dir, sendToSelf)
                Call CheckBladeNPCMatch(index, mapnum)
                Call CheckTileEventsAfterMoving(index, X, Y)
            End If
        End If
    Else
        If HasMapWarpByDir(dir, mapnum) > 0 Then
            PlayerWarpByMapLimits index, dir
            Call ClearPlayerTarget(index)
        Else
            Call SendPlayerXY(index)
        End If
    End If
    

End Sub


Function GetOppositeDir(ByVal dir As Byte) As Byte
    Select Case dir
    Case DIR_UP
        GetOppositeDir = DIR_DOWN
    Case DIR_LEFT
        GetOppositeDir = DIR_RIGHT
    Case DIR_RIGHT
        GetOppositeDir = DIR_LEFT
    Case DIR_DOWN
        GetOppositeDir = DIR_UP
    End Select
End Function


Sub BlockPlayerDirection(ByVal index As Long, ByVal dir As Byte)

End Sub

Sub UnBlockPlayerDirection(ByVal index As Long, ByVal dir As Byte)

End Sub
