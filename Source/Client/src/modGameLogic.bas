Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub GameLoop()
Dim FrameTime As Long
Dim tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim i As Long
Dim WalkTimer As Long
Dim tmr25 As Long
Dim tmr100 As Long
Dim tmr10000 As Long, RainTick As Long, RainUpdateTick As Long
Dim tmr250 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' *** Start GameLoop ***
    Do While InGame
        tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = tick                               ' Set the time second loop time to the first.

    'Send Speed
    If SpeedHack_Timer > 0 And SpeedHack_Timer < tick Then
        SendSpeedAck GetTickCount
        SpeedHack_Timer = GetTickCount + SpeedHack_Lapse * 1000
    End If
' update rain
     If RainTick < tick And map.Weather > 0 Then
         GenerateRainWave
        
         If map.Weather = 1 Then
             RainTick = tick + 100 ' reset time
         ElseIf map.Weather = 2 Then
             RainTick = tick + 1000
         ElseIf map.Weather = 3 Then
             RainTick = tick + 1
         End If
     End If
    
     If RainUpdateTick < tick And map.Weather > 0 Then
         UpdateRainDrops ' update
        
         If map.Weather = 1 Then
             RainUpdateTick = tick + 10
         ElseIf map.Weather = 2 Then
             RainUpdateTick = tick + 50
         ElseIf map.Weather = 3 Then
             RainUpdateTick = tick + 1
         End If
     End If

    If GettingMap = True Then
        If frmMain.Visible = True Then
            frmMain.picLoad.Visible = True
        End If
    End If

        ' * Check surface timers *
        ' Sprites
        If tmr10000 < tick Then

            ' characters
            If NumCharacters > 0 Then
                For i = 1 To NumCharacters    'Check to unload surfaces
                    If CharacterTimer(i) > 0 Then 'Only update surfaces in use
                        If CharacterTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Character(i)), LenB(DDSD_Character(i)))
                            Set DDS_Character(i) = Nothing
                            CharacterTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            ' Paperdolls
            If NumPaperdolls > 0 Then
                For i = 1 To NumPaperdolls    'Check to unload surfaces
                    If PaperdollTimer(i) > 0 Then 'Only update surfaces in use
                        If PaperdollTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Paperdoll(i)), LenB(DDSD_Paperdoll(i)))
                            Set DDS_Paperdoll(i) = Nothing
                            PaperdollTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' animations
            If NumAnimations > 0 Then
                For i = 1 To NumAnimations    'Check to unload surfaces
                    If AnimationTimer(i) > 0 Then 'Only update surfaces in use
                        If AnimationTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Animation(i)), LenB(DDSD_Animation(i)))
                            Set DDS_Animation(i) = Nothing
                            AnimationTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' Items
            If NumItems > 0 Then
                For i = 1 To NumItems    'Check to unload surfaces
                    If ItemTimer(i) > 0 Then 'Only update surfaces in use
                        If ItemTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i)))
                            Set DDS_Item(i) = Nothing
                            ItemTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' Resources
            If NumResources > 0 Then
                For i = 1 To NumResources    'Check to unload surfaces
                    If ResourceTimer(i) > 0 Then 'Only update surfaces in use
                        If ResourceTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Resource(i)), LenB(DDSD_Resource(i)))
                            Set DDS_Resource(i) = Nothing
                            ResourceTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            ' spell icons
            If NumSpellIcons > 0 Then
                For i = 1 To NumSpellIcons    'Check to unload surfaces
                    If SpellIconTimer(i) > 0 Then 'Only update surfaces in use
                        If SpellIconTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_SpellIcon(i)), LenB(DDSD_SpellIcon(i)))
                            Set DDS_SpellIcon(i) = Nothing
                            SpellIconTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            ' faces
            If NumFaces > 0 Then
                For i = 1 To NumFaces    'Check to unload surfaces
                    If FaceTimer(i) > 0 Then 'Only update surfaces in use
                        If FaceTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Face(i)), LenB(DDSD_Face(i)))
                            Set DDS_Face(i) = Nothing
                            FaceTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            ' check ping
            Call GetPing
            Call DrawPing
            tmr10000 = tick + 10000
        End If

        If tmr25 < tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMain.hwnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            ' check if we need to end the CD icon
            If NumSpellIcons > 0 Then
                For i = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(i) > 0 Then
                        If SpellCD(i) > 0 Then
                            If SpellCD(i) + (Spell(PlayerSpells(i)).CDTime * 1000) < tick Then
                                SpellCD(i) = 0
                                BltPlayerSpells
                                BltHotbar
                            End If
                        End If
                    End If
                Next
            End If
            
            ' check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If SpellBufferTimer + (Spell(PlayerSpells(SpellBuffer)).CastTime * 1000) < tick Then
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If

            If CanMoveNow Then
                Call CheckMovement ' Check if player is trying to move
                Call ProcessAttack 'Check attack
                Call CheckIceMovement 'and ice movement
            End If

            ' Change map animation every 250 milliseconds
            If MapAnimTimer < tick Then
                MapAnim = Not MapAnim
                MapAnimTimer = tick + 250
                If RecordingActived Then
                    PushFrame
                End If
            End If
            
            ' Update inv animation
            If NumItems > 0 Then
                If tmr100 < tick Then
                    BltAnimatedInvItems
                    tmr100 = tick + 100
                End If
            End If
            
            For i = 1 To MAX_BYTE
                CheckAnimInstance i
            Next
            
            tmr25 = tick + 50
        End If
        

        If GrowthTimer < tick Then
            CheckStaminaGrowth MyIndex
            CheckIncreaseRideStamina (MyIndex)
            GrowthTimer = GetTickCount + CHECK_GROWTH_TIME * 1000
        End If

        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < tick Then
            
            
            
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    Call ProcessLagMovement(i)
                    Call ProcessMovementSprite(i)
                    Call ProcessMovement(i)
                End If
            Next i

            ' Process npc movements (actually move them)
            For i = 1 To Npc_HighIndex
                If map.NPC(i) > 0 Then
                    Call ProcessNpcMovement(i)
                End If
            Next i

            WalkTimer = tick + 30 ' edit this value to change WalkTimer
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Graphics
        'DoEvents
        
        If wtf = True Then
            If tmr100 < tick Then
            
                Select Case Rand(0, 3)
                
                    Case Is = 0
                        DirUp = True
                    Case Is = 1
                        DirDown = True
                    Case Is = 2
                        DirLeft = True
                    Case Is = 3
                        DirRight = True
                    
                End Select
            
            CheckMovement
            ControlDown = True
            ProcessAttack
        End If
        End If
        
        
        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < tick + 15
                DoEvents
                Sleep 1
            Loop
        End If
        
        ' Calculate fps
        If TickFPS < tick Then
            GameFPS = FPS
            TickFPS = tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If

    Loop

    frmMain.Visible = False

    If isLogging Then
        isLogging = False
        frmMain.picScreen.Visible = False
        frmMenu.Visible = True
        GettingMap = True
        StopMidi
        PlayMidi Options.MenuMusic
    Else
        ' Shutdown the game
        frmLoad.Visible = True
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If
      
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessLagMovement(ByVal index As Long)
    'If Index <> MyIndex Then
        If Player(index).Moving = 0 Then
            If Not (Player(index).LagDirections Is Nothing) Then
                If Not Player(index).LagDirections.IsEmpty Then
                    PlayerMove index, Player(index).LagDirections.Front, Player(index).LagMovements.Front
                    Player(index).automatizedmove = True
                End If
            End If
        End If
    'End If
End Sub

Sub ProcessMovementSprite(ByVal index As Long)
    Dim sprite As Long
    If Player(index).Moving = MOVING_RUNNING And Not Player(index).MovementSprite Then
        sprite = GetRunningSprite(GetPlayerSprite(index))
        If sprite > 0 Then
            Player(index).PreviousSprite = GetPlayerSprite(index)
            SetPlayerSprite index, sprite
            Player(index).MovementSprite = True
        End If
    ElseIf Player(index).Moving <> MOVING_RUNNING And Player(index).MovementSprite Then
        sprite = GetWalkingSprite(index)
        If sprite > 0 Then
            SetPlayerSprite index, sprite
            Player(index).MovementSprite = False
        End If
    End If
End Sub

Sub ProcessMovement(ByVal index As Long)
Dim MovementSpeed As Long



    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if player is walking, and if so process moving them over
    Select Case Player(index).Moving
        Case MOVING_WALKING:
            MovementSpeed = ((ElapsedTime / 1000) * (GetPlayerSpeed(index, MOVING_WALKING) * SIZE_X))
        Case MOVING_RUNNING:
            MovementSpeed = ((ElapsedTime / 1000) * (GetPlayerSpeed(index, MOVING_RUNNING) * SIZE_X))
        Case Else: Exit Sub
    End Select
    If Player(index).step = 0 Then Player(index).step = 1
    
    Select Case GetPlayerDir(index)
        Case DIR_UP
        If Player(index).YOffset <= 0 Then
            'AddText "DIRUP offset for " & index & " Moving = 0 now.", White
            'Player(index).Moving = False
        End If
            Player(index).YOffset = Player(index).YOffset - MovementSpeed
            If Player(index).YOffset < 0 Then Player(index).YOffset = 0
        Case DIR_DOWN
        If Player(index).YOffset = 0 Then
           ' AddText "DIRDOWN offset for " & index & " Moving = 0 now.", White
           ' Player(index).Moving = False
        End If
            Player(index).YOffset = Player(index).YOffset + MovementSpeed
            If Player(index).YOffset > 0 Then Player(index).YOffset = 0
        Case DIR_LEFT
        If Player(index).XOffset = 0 Then
            'AddText "DIRLEFT offset for " & index & " Moving = 0 now.", White
            'Player(index).Moving = False
        End If
            Player(index).XOffset = Player(index).XOffset - MovementSpeed
                        If Player(index).XOffset < 0 Then Player(index).XOffset = 0
        Case DIR_RIGHT
        If Player(index).XOffset = 0 Then
            'AddText "DIRRIGHT offset for " & index & " Moving = 0 now.", White
            'Player(index).Moving = False
        End If
            Player(index).XOffset = Player(index).XOffset + MovementSpeed
            If Player(index).XOffset > 0 Then Player(index).XOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(index).Moving > 0 Then
        If GetPlayerDir(index) = DIR_RIGHT Or GetPlayerDir(index) = DIR_DOWN Then
        
            If (Player(index).XOffset >= 0) And (Player(index).YOffset >= 0) Then
                Player(index).Moving = 0
                
                If Not Player(index).LagDirections Is Nothing Then
                    If Not Player(index).LagDirections.IsEmpty Then
                        If Not (index = MyIndex And Not Player(index).automatizedmove) Then
                            Player(index).LagDirections.Pop
                            Player(index).LagMovements.Pop
                        End If
                    End If
                End If
                
                If Player(index).step = 1 Then
                    Player(index).step = 3
                Else
                    Player(index).step = 1
                End If
                
                If Player(index).LagDirections.IsEmpty Then
                    Player(index).automatizedmove = False
                    Player(index).Started = False
                End If
                
            End If
        Else
            If (Player(index).XOffset <= 0) And (Player(index).YOffset <= 0) Then
                Player(index).Moving = 0
                
                If Not Player(index).LagDirections Is Nothing Then
                    If Not Player(index).LagDirections.IsEmpty Then
                        If Not (index = MyIndex And Not Player(index).automatizedmove) Then
                            Player(index).LagDirections.Pop
                            Player(index).LagMovements.Pop
                        End If
                    End If
                End If
                If Player(index).step = 1 Then
                    Player(index).step = 3
                Else
                    Player(index).step = 1
                End If
                
                If Player(index).LagDirections.IsEmpty Then
                    Player(index).automatizedmove = False
                    Player(index).Started = False
                End If

            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessNpcMovement(ByVal mapnpcnum As Long)
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler
If MapNpc(mapnpcnum).num = 0 Then Exit Sub
Dim NPC_WALK_SPEED As Single

' Check if NPC is walking, and if so process moving them over
If MapNpc(mapnpcnum).Moving = MOVING_WALKING Then
    NPC_WALK_SPEED = 12 - CSng(NPC(MapNpc(mapnpcnum).num).Speed) / 125
    'NPC_WALK_SPEED = 4
    Select Case MapNpc(mapnpcnum).dir
        Case DIR_UP
        MapNpc(mapnpcnum).YOffset = MapNpc(mapnpcnum).YOffset - ((ElapsedTime / 1000 * NPC_WALK_SPEED * SIZE_X))
            If MapNpc(mapnpcnum).YOffset < 0 Then MapNpc(mapnpcnum).YOffset = 0
        
        Case DIR_DOWN
        MapNpc(mapnpcnum).YOffset = MapNpc(mapnpcnum).YOffset + ((ElapsedTime / 1000 * NPC_WALK_SPEED * SIZE_X))
            If MapNpc(mapnpcnum).YOffset > 0 Then MapNpc(mapnpcnum).YOffset = 0
        
        Case DIR_LEFT
        MapNpc(mapnpcnum).XOffset = MapNpc(mapnpcnum).XOffset - ((ElapsedTime / 1000 * NPC_WALK_SPEED * SIZE_X))
            If MapNpc(mapnpcnum).XOffset < 0 Then MapNpc(mapnpcnum).XOffset = 0
        
        Case DIR_RIGHT
        MapNpc(mapnpcnum).XOffset = MapNpc(mapnpcnum).XOffset + ((ElapsedTime / 1000 * NPC_WALK_SPEED * SIZE_X))
            If MapNpc(mapnpcnum).XOffset > 0 Then MapNpc(mapnpcnum).XOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If MapNpc(mapnpcnum).Moving > 0 Then
        If MapNpc(mapnpcnum).dir = DIR_RIGHT Or MapNpc(mapnpcnum).dir = DIR_DOWN Then
            If (MapNpc(mapnpcnum).XOffset >= 0) And (MapNpc(mapnpcnum).YOffset >= 0) Then
            MapNpc(mapnpcnum).Moving = 0
                If MapNpc(mapnpcnum).step = 1 Then
                    MapNpc(mapnpcnum).step = 3
                Else
                    MapNpc(mapnpcnum).step = 1
                End If
            End If
        Else
            If (MapNpc(mapnpcnum).XOffset <= 0) And (MapNpc(mapnpcnum).YOffset <= 0) Then
                MapNpc(mapnpcnum).Moving = 0
                If MapNpc(mapnpcnum).step = 1 Then
                    MapNpc(mapnpcnum).step = 3
                Else
                    MapNpc(mapnpcnum).step = 1
                End If
            End If
        End If
    End If
End If

' Error handler
Exit Sub
errorhandler:
HandleError "ProcessNpcMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Sub CheckMapGetItem()
Dim buffer As New clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer

    If GetTickCount > Player(MyIndex).MapGetTimer + 250 Then
        If Trim$(MyText) = vbNullString Then
            Player(MyIndex).MapGetTimer = GetTickCount
            buffer.WriteLong CMapGetItem
            SendData buffer.ToArray()
        End If
    End If

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMapGetItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAttack()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
              
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        If Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic > 0 Then
            ' projectile
            Set buffer = New clsBuffer
                buffer.WriteLong CProjecTileAttack
                SendData buffer.ToArray()
                Set buffer = Nothing
                Exit Sub
        End If
    End If
                        
    ' non projectile
    Set buffer = New clsBuffer
    buffer.WriteLong CAttack
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAttack", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function IsTryingToMove() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTryingToMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CanMove() As Boolean
Dim d As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    If Not (Player(MyIndex).LagDirections Is Nothing) Then
        If Player(MyIndex).LagDirections.IsEmpty Then
            Player(MyIndex).automatizedmove = False
            Player(MyIndex).Started = False
        End If
    End If
    If Player(MyIndex).automatizedmove Or Player(MyIndex).Started Then
        CanMove = False
        Exit Function
    End If
    
    ' Make sure they haven't just casted a spell
    'If SpellBuffer > 0 Then
    '    CanMove = False
    '    Exit Function
    'End If
    
    ' make sure they're not stunned
    If BlockedActions(aMove) Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' not in bank
    If InBank Then
        'CanMove = False
        'Exit Function
        InBank = False
        frmMain.picBank.Visible = False
    End If

    d = GetPlayerDir(MyIndex)

    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.Left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CanMoveOnIce(ByRef StopSliding As Boolean) As Boolean
Dim d As Long
Dim X As Long
Dim Y As Long
    
    'If Player(MyIndex).OnIce <> True Then Exit Function
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    StopSliding = False
    CanMoveOnIce = True

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMoveOnIce = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If SpellBuffer > 0 Then
        CanMoveOnIce = False
        Exit Function
    End If
    
    ' make sure they're not stunned
    If BlockedActions(aMove) Then
        CanMoveOnIce = False
        Exit Function
    End If
    
    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMoveOnIce = False
        Exit Function
    End If
    
    
    If Not IsPlayerSliding Then
        CanMoveOnIce = False
        StopSliding = True
        Exit Function
    End If

    d = Player(MyIndex).IceDir

    If d = DIR_UP Then
        Call SetPlayerDir(MyIndex, DIR_UP)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMoveOnIce = False
                StopSliding = True

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

                

        Else

            ' Check if they can warp to a new map
            If map.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMoveOnIce = False
            Exit Function
        End If
    End If

    If d = DIR_DOWN Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMoveOnIce = False
                StopSliding = True
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If


        Else

            ' Check if they can warp to a new map
            If map.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMoveOnIce = False
            Exit Function
        End If
    End If

    If d = DIR_LEFT Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMoveOnIce = False
                StopSliding = True
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If


        Else

            ' Check if they can warp to a new map
            If map.Left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMoveOnIce = False
            Exit Function
        End If
    End If

    If d = DIR_RIGHT Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMoveOnIce = False
                StopSliding = True
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If


        Else

            ' Check if they can warp to a new map
            If map.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMoveOnIce = False
            Exit Function
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanMoveOnIce", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function



Function CheckDirection(ByVal direction As Byte) As Boolean
Dim X As Long
Dim Y As Long
Dim i As Long

If MyIndex = 0 Then Exit Function

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    CheckDirection = False
    
    If map.MaxX < GetPlayerX(MyIndex) Or map.MaxY < GetPlayerY(MyIndex) Then
        Exit Function
    End If
    ' check directional blocking
    If isDirBlocked(map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, direction + 1) Then
        CheckDirection = True
        Exit Function
    End If

    Select Case direction
        Case DIR_UP
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex)
    End Select

    ' Check to see if the map tile is blocked or not
    If map.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        For i = 0 To UBound(MapResource)
            If MapResource(i).X = X And MapResource(i).Y = Y Then
                If Resource(map.Tile(X, Y).Data1).WalkableNormal = True And MapResource(i).ResourceState = 0 Then
                    CheckDirection = False
                    Exit For
                ElseIf Resource(map.Tile(X, Y).Data1).WalkableExhausted = True And MapResource(i).ResourceState = 1 Then
                    CheckDirection = False
                    Exit For
                Else
                    CheckDirection = True
                    Exit Function
                End If
            End If
        Next
 
    End If

    ' Check to see if the key door is open or not
    If map.Tile(X, Y).Type = TILE_TYPE_KEY Then

        ' This actually checks if its open or not
        If TempTile(X, Y).DoorOpen = NO Then
            CheckDirection = True
            Exit Function
        End If
    End If
    
    ' Check to see if a npc is already on that tile
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If MapNpc(i).X = X Then
                If MapNpc(i).Y = Y Then
                    If NPC(MapNpc(i).num).Behaviour <> NPC_BEHAVIOUR_BLADE Then
                        If MapNpc(i).petData.Owner = MyIndex Then
                            CheckDirection = False
                        Else
                            CheckDirection = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    
    ' Check to see if a player is already on that tile
    For i = 1 To Player_HighIndex
        If i <> MyIndex Then
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        If Not CanPlayerCrossPlayer(i) Then
                            CheckDirection = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    



    ' Error handler
    Exit Function
errorhandler:
    HandleError "checkDirection", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CanPlayerCrossPlayer(ByVal index As Long) As Boolean
    CanPlayerCrossPlayer = True
    Select Case map.moral
    Case MAP_MORAL_SAFE
        If GetPlayerPK(MyIndex) = YES Or GetPlayerPK(index) = YES Then
            CanPlayerCrossPlayer = Not (CanPlayerAttackPlayer(MyIndex, index) Or CanPlayerAttackPlayer(index, MyIndex))
        End If
    Case MAP_MORAL_PK_SAFE
        If GetPlayerPK(MyIndex) = NO Or GetPlayerPK(MyIndex) = HERO_PLAYER Or GetPlayerPK(index) = NO Or GetPlayerPK(index) = HERO_PLAYER Then
            CanPlayerCrossPlayer = Not (CanPlayerAttackPlayer(MyIndex, index) Or CanPlayerAttackPlayer(index, MyIndex))
        End If
    Case MAP_MORAL_PACIFIC
        CanPlayerCrossPlayer = True
    Case Else
        CanPlayerCrossPlayer = Not (CanPlayerAttackPlayer(MyIndex, index) Or CanPlayerAttackPlayer(index, MyIndex))
    End Select
End Function
Function CheckMapMorals(ByVal mapnum As Long, ByVal attacker As Long, ByVal victim As Long) As Boolean
    ' Check if map is attackable
    Dim moral As Byte
    moral = map.moral
    Select Case moral
    Case MAP_MORAL_NONE
        CheckMapMorals = True
    Case MAP_MORAL_SAFE
        If IsPlayerNeutral(victim) Then
        Else
            CheckMapMorals = True
        End If
    Case MAP_MORAL_ARENA
        CheckMapMorals = True
    Case MAP_MORAL_PK_SAFE
        If Not IsPlayerNeutral(victim) Then
            CheckMapMorals = False
        Else
            CheckMapMorals = True
        End If
    Case MAP_MORAL_PACIFIC
        CheckMapMorals = False
    End Select
    
    
End Function

Public Function IsPlayerNeutral(ByVal index As Long) As Boolean

IsPlayerNeutral = Not (Player(index).PK = YES)

End Function

Public Function CheckLevels(ByVal attacker As Long, ByVal victim As Long, Optional ByVal sendmsg As Boolean = True) As Boolean
CheckLevels = False
If Not IsPlaying(attacker) Or Not IsPlaying(victim) Then Exit Function

' Make sure attacker is high enough level
If GetPlayerLevel(attacker) < 15 Then
    Exit Function
End If

If GetPlayerLevel(victim) < 15 Then
    Exit Function
End If

CheckLevels = True
End Function

Function CanPlayerAttackByJustice(ByVal attacker As Long, ByVal victim As Long, Optional ByVal sendmsg As Boolean = True) As Boolean
    CanPlayerAttackByJustice = True
    If (GetPlayerPK(attacker) = PK_PLAYER And GetPlayerPK(victim) = NONE_PLAYER) Or (GetPlayerPK(attacker) = NONE_PLAYER And GetPlayerPK(victim) = PK_PLAYER) Then
        If Abs(GetPlayerLevel(attacker) - GetPlayerLevel(victim)) > 20 Then
             CanPlayerAttackByJustice = False
        End If
    End If
End Function
Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal IsProjectile As Boolean = False) As Boolean



    If Not CheckMapMorals(GetPlayerMap(attacker), attacker, victim) Then Exit Function
    
    If map.moral = MAP_MORAL_ARENA Then
        CanPlayerAttackPlayer = True
        Exit Function
    End If
    
    'Safe Mode

    ' Make sure they have more then 0 hp

    ' Check to make sure that they dont have access
    If GetPlayerAccess(attacker) > ADMIN_MONITOR Then
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Exit Function
    End If

    'Check if levels are correct
    If Not CheckLevels(attacker, victim, False) Then
        Exit Function
    End If
    
    If Not CanPlayerAttackByJustice(attacker, victim) Then
        Exit Function
    End If

    
    CanPlayerAttackPlayer = True
    
End Function



Sub CheckMovement()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If IsTryingToMove Then
        If CanMove Then
            ' Check if player has the shift key down for running
            Dim UsedMovement As Byte
            Dim UsedSpeed As Long
            If ShiftDown Then
                Call GetPlayerRunSpeed(MyIndex, UsedMovement, UsedSpeed)
            Else
                UsedMovement = MOVING_WALKING
                UsedSpeed = GetPlayerSpeed(MyIndex, UsedMovement)
            End If
            
            Player(MyIndex).Moving = UsedMovement

            Select Case GetPlayerDir(MyIndex)
                Case DIR_UP
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select
            
            If Not Player(MyIndex).onIce Then
                If GetMapTileType(GetPlayerX(MyIndex), GetPlayerY(MyIndex)) = TILE_TYPE_ICE Then
                    Player(MyIndex).onIce = True
                    Player(MyIndex).IceDir = GetPlayerDir(MyIndex)
                End If
            End If
            

            If Player(MyIndex).XOffset = 0 Then
                If Player(MyIndex).YOffset = 0 Then
                    If map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                        GettingMap = True
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isInBounds()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If (CurX >= 0) Then
        If (CurX <= map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isInBounds", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub UpdateDrawMapName()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    DrawMapNameX = Camera.Left + ((MAX_MAPX + 1) * PIC_X / 2) - getWidth(TexthDC, (map.Name))
    DrawMapNameY = Camera.Top + 1

    Select Case map.moral
        Case MAP_MORAL_NONE
            DrawMapNameColor = QBColor(BrightRed)
        Case MAP_MORAL_SAFE
            DrawMapNameColor = QBColor(White)
        Case MAP_MORAL_ARENA
            DrawMapNameColor = QBColor(Yellow)
        Case MAP_MORAL_PK_SAFE
            DrawMapNameColor = QBColor(Red)
        Case MAP_MORAL_PACIFIC
            DrawMapNameColor = QBColor(Cyan)
        Case Else
            DrawMapNameColor = QBColor(White)
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateDrawMapName", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UseItem()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UseItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ForgetSpell(ByVal spellslot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If SpellCD(spellslot) > 0 Then
        AddText "You can't forget an ability while recharging!", BrightRed, True
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellslot Then
        AddText "You can't forget a skill while using it!", BrightRed, True
        Exit Sub
    End If
    
    If PlayerSpells(spellslot) > 0 Then
        Set buffer = New clsBuffer
        buffer.WriteLong CForgetSpell
        buffer.WriteLong spellslot
        SendData buffer.ToArray()
        Set buffer = Nothing
    Else
        AddText "Vacío", BrightRed, True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ForgetSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CastSpell(ByVal spellslot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    If SpellCD(spellslot) > 0 Then
        'AddText "Magic recharging.", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellslot) = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.mp) < Spell(PlayerSpells(spellslot)).MPCost Then
        'Call AddText("You don't have enough MP to pitch" & Trim$(Spell(PlayerSpells(spellslot)).Name) & ".", BrightRed)
        CreateActionMsg "Without enough magic points!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(MyIndex) * 32, GetPlayerY(MyIndex) * 32, True
        Exit Sub
    End If

    If PlayerSpells(spellslot) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set buffer = New clsBuffer
                buffer.WriteLong CCast
                buffer.WriteLong spellslot
                SendData buffer.ToArray()
                Set buffer = Nothing
                SpellBuffer = spellslot
                SpellBufferTimer = GetTickCount
            Else
                Call AddText("You can't summon while running.", BrightRed, True)
            End If
        End If
    Else
        Call AddText("Gap", BrightRed, True)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "castSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearTempTile()
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ReDim TempTile(0 To map.MaxX, 0 To map.MaxY)

    For X = 0 To map.MaxX
        For Y = 0 To map.MaxY
            TempTile(X, Y).DoorOpen = NO
        Next
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearTempTile", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DevMsg(ByVal text As String, ByVal Color As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(text, Color)
        End If
    End If

    Debug.Print text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DevMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "TwipsToPixels", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "PixelsToTwips", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertCurrency(ByVal amount As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Int(amount) < 10000 Then
        ConvertCurrency = amount
    ElseIf Int(amount) < 999999 Then
        ConvertCurrency = Int(amount / 1000) & "k"
    ElseIf Int(amount) < 999999999 Then
        ConvertCurrency = Int(amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(amount / 1000000000) & "b"
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertCurrency", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub DrawPing()
Dim PingToDraw As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PingToDraw = Ping

    Select Case Ping
        Case -1
            PingToDraw = "Syncing"
        Case 0 To 5
            PingToDraw = "Local"
    End Select

    frmMain.lblPing.Caption = "Ping: " & PingToDraw
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPing", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateSpellWindow(ByVal spellnum As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check for off-screen
    If Y + frmMain.picSpellDesc.Height > frmMain.ScaleHeight Then
        Y = frmMain.ScaleHeight - frmMain.picSpellDesc.Height
    End If
    
    With frmMain
        .picSpellDesc.Top = Y
        .picSpellDesc.Left = X
        .picSpellDesc.Visible = True
        
        If LastSpellDesc = spellnum Then Exit Sub
        
        .lblSpellName.Caption = Trim$(Spell(spellnum).Name)
        .lblSpellDesc.Caption = Trim$(Spell(spellnum).Desc)
        BltSpellDesc spellnum
    End With
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdteSpellWindow", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateDescWindow(ByVal ItemNum As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Long
Dim FirstLetter As String * 1
Dim Name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    FirstLetter = LCase$(Left$(Trim$(Item(ItemNum).Name), 1))
   
    If FirstLetter = "$" Then
        Name = (Mid$(Trim$(Item(ItemNum).Name), 2, Len(Trim$(Item(ItemNum).Name)) - 1))
    Else
        Name = Trim$(Item(ItemNum).Name)
    End If
    
    ' check for off-screen
    If Y + frmMain.picItemDesc.Height > frmMain.ScaleHeight Then
        Y = frmMain.ScaleHeight - frmMain.picItemDesc.Height
    End If
    
    ' set z-order
    frmMain.picItemDesc.ZOrder (0)

    With frmMain
        .picItemDesc.Top = Y
        .picItemDesc.Left = X
        .picItemDesc.Visible = True

        If LastItemDesc = ItemNum Then Exit Sub ' exit out after setting x + y so we don't reset values

        ' set the name
        Select Case Item(ItemNum).Rarity
            Case 0 ' white
                .lblItemName.ForeColor = RGB(255, 255, 255)
            Case 1 ' green
                .lblItemName.ForeColor = RGB(117, 198, 92)
            Case 2 ' blue
                .lblItemName.ForeColor = RGB(103, 140, 224)
            Case 3 ' maroon
                .lblItemName.ForeColor = RGB(205, 34, 0)
            Case 4 ' purple
                .lblItemName.ForeColor = RGB(193, 104, 204)
            Case 5 ' orange
                .lblItemName.ForeColor = RGB(217, 150, 64)
        End Select
        
        ' set captions
        
        .lblItemName.Caption = Name
        
        .lblItemDesc.Caption = Item(ItemNum).Desc

        .lblItemWeight.Caption = "Weight: " & Item(ItemNum).Weight
        ' render the item
        BltItemDesc ItemNum
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateDescWindow", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CacheResources()
Dim X As Long, Y As Long, Resource_Count As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource_Count = 0

    For X = 0 To map.MaxX
        For Y = 0 To map.MaxY
            If map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).X = X
                MapResource(Resource_Count).Y = Y
            End If
        Next
    Next

    Resource_Index = Resource_Count
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CreateActionMsg(ByVal message As String, ByVal Color As Integer, ByVal MsgType As Byte, ByVal X As Long, ByVal Y As Long, Optional blTranslate As Boolean = False)
Dim i As Long

    If Not IsNumeric(message) Then
        If blTranslate = True Then message = message
    Else
        If message = 0 Then Exit Sub
    End If

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .message = message
        .Color = Color
        .Type = MsgType
        .Created = GetTickCount
        .Scroll = 1
        .X = X
        .Y = Y
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).Y = ActionMsg(ActionMsgIndex).Y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).X = ActionMsg(ActionMsgIndex).X + Rand(-8, 8)
    End If
    
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CreateActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearActionMsg(ByVal index As Byte)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionMsg(index).message = vbNullString
    ActionMsg(index).Created = 0
    ActionMsg(index).Type = 0
    ActionMsg(index).Color = 0
    ActionMsg(index).Scroll = 0
    ActionMsg(index).X = 0
    ActionMsg(index).Y = 0
    
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimInstance(ByVal index As Long)
Dim looptime As Long
Dim layer As Long
Dim FrameCount As Long
Dim lockindex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if doesn't exist then exit sub
    If AnimInstance(index).Animation <= 0 Then Exit Sub
    If AnimInstance(index).Animation >= MAX_ANIMATIONS Then Exit Sub
    
    For layer = 0 To 1
        If AnimInstance(index).Used(layer) Then
            looptime = Animation(AnimInstance(index).Animation).looptime(layer)
            FrameCount = Animation(AnimInstance(index).Animation).Frames(layer)
            
            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(index).FrameIndex(layer) = 0 Then AnimInstance(index).FrameIndex(layer) = 1
            If AnimInstance(index).LoopIndex(layer) = 0 Then AnimInstance(index).LoopIndex(layer) = 1
            
            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(index).Timer(layer) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimInstance(index).FrameIndex(layer) >= FrameCount Then
                    AnimInstance(index).LoopIndex(layer) = AnimInstance(index).LoopIndex(layer) + 1
                    If AnimInstance(index).LoopIndex(layer) > Animation(AnimInstance(index).Animation).LoopCount(layer) Then
                        AnimInstance(index).Used(layer) = False
                    Else
                        AnimInstance(index).FrameIndex(layer) = 1
                    End If
                Else
                    AnimInstance(index).FrameIndex(layer) = AnimInstance(index).FrameIndex(layer) + 1
                End If
                AnimInstance(index).Timer(layer) = GetTickCount
            End If
        End If
    Next
    
    ' if neither layer is used, clear
    If AnimInstance(index).Used(0) = False And AnimInstance(index).Used(1) = False Then ClearAnimInstance (index)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "checkAnimInstance", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenShop(ByVal Shopnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InShop = Shopnum
    ShopAction = 0
    frmMain.picShop.Visible = True
    BltShop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "OpenShop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemNum(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If bankslot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    If bankslot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    GetBankItemNum = Bank.Item(bankslot).num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal ItemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).num = ItemNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemValue(ByVal index As Long, ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If index > MAX_PLAYERS Then Exit Function
    GetBankItemValue = Bank.Item(bankslot).value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockVar As Byte, ByRef dir As Byte, ByVal block As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If block Then
        blockVar = blockVar Or (2 ^ dir)
    Else
        blockVar = blockVar And Not (2 ^ dir)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "setDirBlock", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isDirBlocked(ByRef blockVar As Byte, ByRef dir As Byte) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not blockVar And (2 ^ dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsHotbarSlot(ByVal X As Single, ByVal Y As Single) As Long
Dim Top As Long, Left As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    IsHotbarSlot = 0

    For i = 1 To MAX_HOTBAR
        Top = HotbarTop
        Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
        If X >= Left And X <= Left + PIC_X Then
            If Y >= Top And Y <= Top + PIC_Y Then
                IsHotbarSlot = i
                Exit Function
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsHotbarSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub PlayMapSound(ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim soundName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If entityNum <= 0 Then Exit Sub
    
    ' find the sound
    Select Case entityType
        ' animations
        Case SoundEntity.seAnimation
            If entityNum > MAX_ANIMATIONS Then Exit Sub
            soundName = Trim$(Animation(entityNum).Sound)
            
        ' items
        Case SoundEntity.seItem
            If entityNum > MAX_ITEMS Then Exit Sub
            soundName = Trim$(Item(entityNum).Sound)
        ' npcs
        Case SoundEntity.seNpc
            If entityNum > MAX_NPCS Then Exit Sub
            soundName = Trim$(NPC(entityNum).Sound)
        ' resources
        Case SoundEntity.seResource
            If entityNum > MAX_RESOURCES Then Exit Sub
            soundName = Trim$(Resource(entityNum).Sound)
        ' spells
        Case SoundEntity.seSpell
            If entityNum > MAX_SPELLS Then Exit Sub
            soundName = Trim$(Spell(entityNum).Sound)
        ' LevelUp
            Case SoundEntity.seLevelUp
            soundName = Trim$(LEVEL_SOUND)
        ' Switch
            Case SoundEntity.seSwitch
            soundName = Trim$(Switch)
        ' SwitchFloor
            Case SoundEntity.seSwitchFloor
            soundName = Trim$(SwitchFloor)
        ' Sandstorm
            Case SoundEntity.seSandstorm
            soundName = Trim$(Sandstorm)
        ' Slide
            Case SoundEntity.seSlide
            soundName = Trim$(Slide)
        ' class sounds
            Case SoundEntity.seAttack To SoundEntity.seDie
            soundName = GetClassSound(entityNum, entityType)
        ' Reset
            Case SoundEntity.seReset
            soundName = Trim$(Reset)
        ' Error
            Case SoundEntity.seError
            soundName = Trim$(Error)
        Case Else
            Exit Sub
    End Select
    
    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub

    ' play the sound
    PlaySound soundName
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayMapSound", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetClassSound(ByVal entityNum As Long, ByVal entityType As SoundEntity) As String
Dim ClassName As String
Select Case entityNum
Case 1 To 4 'hylian
    ClassName = "Hylian"
Case 5 To 6 'zora
    ClassName = "Goron"
Case 7 To 8 'goron
    ClassName = "Zora"
Case 9 To 10
    ClassName = "Gerudo"
End Select

GetClassSound = ClassName & GetEntitySoundTypeName(entityType) & ".wav"
End Function

Function GetEntitySoundTypeName(ByVal entityType As SoundEntity) As String
Select Case entityType
Case seAttack
    GetEntitySoundTypeName = "Attack"
Case seCritical
    GetEntitySoundTypeName = "Critical"
Case seHit
    GetEntitySoundTypeName = "Hit"
Case seDie
    GetEntitySoundTypeName = "Die"
End Select
End Function

Public Sub Dialogue(ByVal diTitle As String, ByVal diText As String, ByVal diIndex As Long, Optional ByVal isYesNo As Boolean = False, Optional ByVal Data1 As Long = 0)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' exit out if we've already got a dialogue open
    If dialogueIndex > 0 Then Exit Sub
    
    ' set global dialogue index
    dialogueIndex = diIndex
    
    ' set the global dialogue data
    dialogueData1 = Data1

    ' set the captions
    frmMain.lblDialogue_Title.Caption = diTitle
    frmMain.lblDialogue_Text.Caption = diText
    
    ' show/hide buttons
    If Not isYesNo Then
        frmMain.lblDialogue_Button(1).Visible = True ' Okay button
        frmMain.lblDialogue_Button(2).Visible = False ' Yes button
        frmMain.lblDialogue_Button(3).Visible = False ' No button
    Else
        frmMain.lblDialogue_Button(1).Visible = False ' Okay button
        frmMain.lblDialogue_Button(2).Visible = True ' Yes button
        frmMain.lblDialogue_Button(3).Visible = True ' No button
    End If
    
    ' show the dialogue box
    frmMain.picDialogue.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Dialogue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub dialogueHandler(ByVal index As Long)
    ' find out which button
    If index = 1 Then ' okay button
        ' dialogue index
        Select Case dialogueIndex
        
        End Select
    ElseIf index = 2 Then ' yes button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendAcceptTradeRequest
            Case DIALOGUE_TYPE_FORGET
                ForgetSpell dialogueData1
            Case DIALOGUE_TYPE_PARTY
                SendAcceptParty
            Case DIALOGUE_TYPE_QUESTION
                SendQuestionAnswer True
        End Select
    ElseIf index = 3 Then ' no button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendDeclineTradeRequest
            Case DIALOGUE_TYPE_PARTY
                SendDeclineParty
            Case DIALOGUE_TYPE_QUESTION
                SendQuestionAnswer False
        End Select
    End If
End Sub

Public Sub AddChatBubble(ByVal target As Long, ByVal targetType As Byte, ByVal msg As String, ByVal colour As Long)
Dim i As Long, index As Long

    ' set the global index
    chatBubbleIndex = chatBubbleIndex + 1
    If chatBubbleIndex < 1 Or chatBubbleIndex > MAX_BYTE Then chatBubbleIndex = 1
    
    ' default to new bubble
    index = chatBubbleIndex
    
    ' loop through and see if that player/npc already has a chat bubble
    For i = 1 To MAX_BYTE
        If chatBubble(i).targetType = targetType Then
            If chatBubble(i).target = target Then
                ' reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                ' we use this one now, yes?
                index = i
                Exit For
            End If
        End If
    Next
    
    ' set the bubble up
    With chatBubble(index)
        .target = target
        .targetType = targetType
        .msg = msg
        .colour = colour
        .Timer = GetTickCount
        .active = True
    End With
End Sub

Public Sub GenerateRainDrop(ByVal RainIndex As Long)
' generate raindrop
RainDrop(RainIndex).X = Rand(1, frmMain.picScreen.Width)
RainDrop(RainIndex).Y = 0
RainDrop(RainIndex).InMotion = 1
End Sub

Public Sub GenerateRainWave()
Dim i As Long

' generate a wave
For i = 1 To Rand(1, 3)
     GenerateRainDrop FindRainSlot
    
     If map.Weather = 2 Then Exit Sub
Next

End Sub

Public Function FindRainSlot()
Dim i As Long, RainMax As Long

If map.Weather = 1 Then RainMax = 25
If map.Weather = 2 Then RainMax = 15
If map.Weather = 3 Then RainMax = 100

' check if in motion
For i = 1 To RainMax
     If RainDrop(i).InMotion = 0 Then
         Exit For
     End If
Next

' replace rain
FindRainSlot = i
End Function

Public Sub UpdateRainDrops()
Dim i As Long, RainMax As Long

If map.Weather = 1 Then RainMax = 25
If map.Weather = 2 Then RainMax = 15
If map.Weather = 3 Then RainMax = 100

' loop through all raindrops
For i = 1 To RainMax
     If Not RainDrop(i).InMotion = 0 Then
         RainDrop(i).Y = RainDrop(i).Y + 5
         If map.Weather = 3 Then
         RainDrop(i).Y = RainDrop(i).Y + 20
         End If
         ' if it's past the screen, reset.
         If RainDrop(i).Y > frmMain.picScreen.Height Then
             RainDrop(i).InMotion = 0
         End If
     End If
Next
End Sub

Public Function BTI(ByVal Var As Boolean) As Long

If Var Then
    BTI = 1
Else
    BTI = 0
End If

End Function

Public Function ITB(ByVal Var As Long) As Boolean

If Var = 1 Then
    ITB = True
Else
    ITB = False
End If

End Function

Public Function STB(ByVal Var As String) As Boolean

If Var = "1" Then
    STB = True
Else
   STB = False
End If

End Function

Public Function isItemStackable(ByVal numitem As Long) As Boolean

Dim ItemType As Byte
    'Return True if item is stackable
    If numitem > 0 And numitem <= MAX_ITEMS Then
        ItemType = Item(numitem).Type
        If ItemType = ITEM_TYPE_CURRENCY Or ItemType = ITEM_TYPE_CONSUME Then
            isItemStackable = True
            Exit Function
        End If
    End If

isItemStackable = False

End Function

Public Function DirtoStr(ByVal dir As Byte) As String
Select Case dir
Case 0
    DirtoStr = "Up"
Case 1
    DirtoStr = "Down"
Case 2
    DirtoStr = "Left"
Case 3
    DirtoStr = "Right"
Case Else
    DirtoStr = "None"
End Select
End Function

Public Function IsEmptyList(ByRef list As MovementsListRec)
Dim a As Variant
If UBound(list.vect) > 0 Then
IsEmptyList = False
Else
IsEmptyList = True
End If
End Function


Sub SpawnPet(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSpawnPet
    SendData buffer.ToArray()
    Set buffer = Nothing

End Sub

Sub PetFollow(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPetFollowOwner
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub PetAttack(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPetAttackTarget
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub PetWander(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPetWander
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub PetDisband(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPetDisband
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Function CheckFreePetSlots(ByVal index As Long) As Integer
'-1: There aren't free slots, >= 1: Index of first free slot
Dim i As Byte
If index <> MyIndex Then Exit Function

CheckFreePetSlots = -1
For i = 1 To MAX_PLAYER_PETS

    If Player(MyIndex).Pet(i).NumPet = 0 Then
        CheckFreePetSlots = i
        Exit Function
    End If
    
Next

End Function

Public Sub RefreshPetVitals(ByVal index As Long)
    'mp
    With frmMain
    
    If GetPlayerPetMapNPCNum(index) > 0 Then
        .lblPetMPNum.Caption = MapNpc(GetPlayerPetMapNPCNum(index)).vital(Vitals.mp) & "/" & GetNpcMaxVital(MapNpc(GetPlayerPetMapNPCNum(index)).num, Vitals.mp, index)
    Else
        .lblPetMPNum.Caption = "0/0"
    End If
    
    End With
        
End Sub

Public Sub RefreshPetData(ByVal index As Long)
Dim ActualPet As Byte
Dim X As Long
ActualPet = Player(index).ActualPet
If ActualPet < 1 Or ActualPet > MAX_PLAYER_PETS Then Exit Sub


    With frmMain
        'Name
        Dim PetNum As Byte
        PetNum = Player(index).Pet(ActualPet).NumPet
        
        Dim NPCNum As Long
        NPCNum = 0
        If PetNum > 0 And PetNum <= MAX_PETS Then
           NPCNum = Pet(PetNum).NPCNum
        End If
        
        If NPCNum > 0 Then
            .lblChoosePet.Caption = "Name: " & Trim$(NPC(NPCNum).Name)
            .lblPetName.Caption = Trim$(NPC(NPCNum).Name)
        Else
            .lblChoosePet.Caption = "Name:"
            .lblPetName.Caption = "Empty"
        End If
        'Lvl
        If Player(index).Pet(ActualPet).Level > 0 Then
            .lblPetLvlNum.Caption = Player(index).Pet(ActualPet).Level
        Else
            .lblPetLvlNum.Caption = ""
        End If
        'Experience
        .lblPetExpNum = Player(index).Pet(ActualPet).Experience & "%" '& "/" & GetPlayerPetNextLevel(MyIndex)
        'Stats
        'Points
        For X = 1 To Stats.Stat_Count - 1
            frmMain.lblCharStat(X + Stats.Stat_Count - 1).Caption = GetPlayerPetStat(MyIndex, X)
        Next
        
        ' Set training label visiblity depending on points
        .lblPetPoints.Caption = GetPlayerPetPOINTS(MyIndex)
        If GetPlayerPetPOINTS(MyIndex) > 0 Then
            For X = 1 To Stats.Stat_Count - 1
                If GetPlayerPetStat(index, X) < 255 Then
                    frmMain.lblTrainStat(X + Stats.Stat_Count - 1).Visible = True
                Else
                    frmMain.lblTrainStat(X + Stats.Stat_Count - 1).Visible = False
                End If
            Next
        Else
            For X = 1 To Stats.Stat_Count - 1
                frmMain.lblTrainStat(X + Stats.Stat_Count - 1).Visible = False
            Next
        End If
          
    End With
    
        
    
End Sub

Public Function GetPlayerPetNextLevel(ByVal index As Long) As Long
    If Player(index).Pet(Player(index).ActualPet).NumPet < 1 Or Player(index).Pet(Player(index).ActualPet).NumPet > MAX_PETS Then Exit Function
    'GetPlayerPetNextLevel = (50 / 6) * ((GetPlayerPetLevel(Index) + Pet(Player(Index).Pet(Player(Index).ActualPet).NumPet).ExpProgression) ^ 3 - (6 * (GetPlayerPetLevel(Index) + Pet(Player(Index).Pet(Player(Index).ActualPet).NumPet).ExpProgression) ^ 2) + 17 * (GetPlayerPetLevel(Index) + Pet(Player(Index).Pet(Player(Index).ActualPet).NumPet).ExpProgression) - 12) + 50

End Function

Public Function GetPlayerPetLevel(ByVal index As Long) As Long
    GetPlayerPetLevel = Player(index).Pet(Player(index).ActualPet).Level
End Function

Public Function GetPlayerPetMapNPCNum(ByVal index As Long) As Long
Dim i As Long

    GetPlayerPetMapNPCNum = 0

    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).petData.Owner = index Then
                GetPlayerPetMapNPCNum = i
                Exit Function
        End If
    Next

End Function

Public Function GetPlayerTriforce(ByVal index As Long, ByVal triforce As TriforceType) As Boolean
Dim i As Byte

GetPlayerTriforce = False

If triforce > 0 And triforce < TriforceType_Count Then
    GetPlayerTriforce = Player(index).triforce(triforce)
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

Public Function GetPlayerNameColorByTriforce(ByVal index As Long) As Long

Dim Red As Byte
Dim Blue As Byte
Dim Green As Byte

'Normal Color
If HasPlayerAnyTriforce(index) = False Then
    
    GetPlayerNameColorByTriforce = RGB(255, 170, 70)
    Exit Function
End If

If GetPlayerTriforce(index, TRIFORCE_WISDOM) Then
    Blue = 127
End If
If GetPlayerTriforce(index, TRIFORCE_COURAGE) Then
    Green = 127
End If
If GetPlayerTriforce(index, TRIFORCE_POWER) Then
    Red = 127
End If

GetPlayerNameColorByTriforce = RGB(Red, Green, Blue)

End Function

Public Function GetPlayerNameColorByJustice(ByVal index As Long) As Long

Select Case GetPlayerPK(index)
Case NONE_PLAYER
    GetPlayerNameColorByJustice = NoneColor
Case PK_PLAYER
    GetPlayerNameColorByJustice = QBColor(PKColor)
Case HERO_PLAYER
    GetPlayerNameColorByJustice = HeroColor
Case Else
    GetPlayerNameColorByJustice = NoneColor
End Select

End Function

Public Function GetNameColorByJustice(ByVal PK As Byte) As Long

Select Case PK
Case NONE_PLAYER
    GetNameColorByJustice = NoneColor
Case PK_PLAYER
    GetNameColorByJustice = QBColor(PKColor)
Case HERO_PLAYER
    GetNameColorByJustice = HeroColor
Case Else
    GetNameColorByJustice = NoneColor
End Select

End Function


Public Sub ForceMovement(ByVal dir As Byte)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    ' Check if player has the shift key down for running

    Player(MyIndex).Moving = MOVING_WALKING

    Select Case dir
        Case DIR_UP
            Call SendPlayerForcedMove(dir)
            Player(MyIndex).YOffset = PIC_Y
            Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
        Case DIR_DOWN
            Call SendPlayerForcedMove(dir)
            Player(MyIndex).YOffset = PIC_Y * -1
            Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
        Case DIR_LEFT
            Call SendPlayerForcedMove(dir)
            Player(MyIndex).XOffset = PIC_X
            Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
        Case DIR_RIGHT
            Call SendPlayerForcedMove(dir)
            Player(MyIndex).XOffset = PIC_X * -1
            Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
    End Select

    If Player(MyIndex).XOffset = 0 Then
        If Player(MyIndex).YOffset = 0 Then
            If map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                GettingMap = True
            End If
        End If
    End If

    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ForcedMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Sub CheckIceMovement()
    Dim StopSliding As Boolean
    Dim StartSliding As Boolean
    
    If Player(MyIndex).onIce Then
        
        If CanMoveOnIce(StopSliding) Then
            ForceMovement (Player(MyIndex).IceDir)
        Else
            If (StopSliding) Then
                Player(MyIndex).onIce = False
                SendOnIce MyIndex, False
            End If
        End If
    End If


End Sub

Function IsPlayerSliding()
    If GetMapTileType(GetPlayerX(MyIndex), GetPlayerY(MyIndex)) = TILE_TYPE_ICE Then
        IsPlayerSliding = Player(MyIndex).onIce
    Else
        IsPlayerSliding = False
    End If
End Function

Public Function GetPlayerMaxMoney(ByVal index As Long) As Long
    GetPlayerMaxMoney = GetMaxMoneyByBag(GetPlayerBags(index))
End Function

Public Function GetPlayerBags(ByVal index As Long) As Byte
    GetPlayerBags = Player(index).RupeeBags
End Function

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

Public Function GetMaxMoneyByBag(ByVal bags As Byte) As Long
    If (bags >= MAX_RUPEE_BAGS) Then
        GetMaxMoneyByBag = bags * BAG_CAPACITY - 1
    Else
        GetMaxMoneyByBag = bags * BAG_CAPACITY
    End If
End Function

Public Function CheckAttackNPC() As Boolean
'OUT: True if there is 1 NPC to attack
    Dim N As Long
    
    N = GetPlayerEquipment(MyIndex, Weapon)
    If N > 0 Then
        If Item(N).ProjecTile.Pic > 0 Then
            Exit Function
        End If
    End If
                        
    ' non projectile
    Dim i As Long
    For i = 1 To MAX_MAP_NPCS
        If CanPlayerAttackNpc(i) Then
            SendAttackNPC (i)
            CheckAttackNPC = True
            Exit Function
        End If
    Next

End Function

Public Function CheckResource() As Boolean
'OUT: True if there is 1 NPC to attack
    Dim X As Long
    Dim Y As Long
    X = GetPlayerX(MyIndex)
    Y = GetPlayerY(MyIndex)
    
    If GetNextPositionByRef(Player(MyIndex).dir, X, Y) Then Exit Function
    
    If GetMapTileType(X, Y) <> TILE_TYPE_RESOURCE Then Exit Function
                    
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        If Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic > 0 Then
            Exit Function
        End If
    End If
                        
    ' non projectile
    Dim i As Long
    For i = 0 To UBound(MapResource)
        If MapResource(i).X = X And MapResource(i).Y = Y Then
            If MapResource(i).ResourceState = 0 Then 'non cut
                SendCheckResource (i)
                CheckResource = True
                Exit Function
            End If
        End If
    Next

End Function

Public Function CanPlayerAttackNpc(ByVal mapnpcnum As Long) As Boolean
    Dim mapnum As Long
    Dim NPCNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(mapnpcnum).num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(MyIndex)
    NPCNum = MapNpc(mapnpcnum).num

    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnpcnum).vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    If Not CanNPCBeAttacked(NPCNum) Then Exit Function
    
    Select Case GetPlayerDir(MyIndex)
        Case DIR_UP
            NpcX = MapNpc(mapnpcnum).X
            NpcY = MapNpc(mapnpcnum).Y + 1
        Case DIR_DOWN
            NpcX = MapNpc(mapnpcnum).X
            NpcY = MapNpc(mapnpcnum).Y - 1
        Case DIR_LEFT
            NpcX = MapNpc(mapnpcnum).X + 1
            NpcY = MapNpc(mapnpcnum).Y
        Case DIR_RIGHT
            NpcX = MapNpc(mapnpcnum).X - 1
            NpcY = MapNpc(mapnpcnum).Y
    End Select
    
     ' Check if at same coordinates
    If Not NpcX = GetPlayerX(MyIndex) And NpcY = GetPlayerY(MyIndex) Then Exit Function
    
    If NPCNum > 0 And GetTickCount > Player(MyIndex).AttackTimer + attackspeed Then
        CanPlayerAttackNpc = True
    End If


End Function

Public Function CanNPCBeAttacked(ByVal NPCNum As Long) As Boolean
If Not (NPCNum > 0 And NPCNum < MAX_NPCS) Then Exit Function

CanNPCBeAttacked = True

Select Case NPC(NPCNum).Behaviour

    Case NPC_BEHAVIOUR_FRIENDLY
        CanNPCBeAttacked = False
    Case NPC_BEHAVIOUR_BLADE
        CanNPCBeAttacked = False
    Case NPC_BEHAVIOUR_SHOPKEEPER
        CanNPCBeAttacked = False
    Case NPC_BEHAVIOUR_SLIDE
        CanNPCBeAttacked = False
End Select

End Function


Sub ProcessAttack(Optional blResourceOnly As Boolean = False)
Dim attackspeed As Long

    If ControlDown Or blResourceOnly Then
        
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If BlockedActions(aAttack) Then Exit Sub  ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            SendFSpellActiEmptyn (MyIndex)
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).Speed
        Else
            attackspeed = 1000
        End If
        
        If Player(MyIndex).AttackTimer + attackspeed < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then
                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = GetTickCount
                End With
                
                If blResourceOnly = True Then
                    CheckResource
                    'CheckAttack
                    Exit Sub
                End If
                
                If Not CheckAttackNPC Then
                    If Not CheckResource Then
                        CheckAttack
                    End If
                End If
                
            End If

        End If
    End If

End Sub

Function GetMapTileType(ByVal X As Long, ByVal Y As Long) As Byte
    If X < 0 Or Y < 0 Or X > map.MaxX Or Y > map.MaxY Then Exit Function
    GetMapTileType = map.Tile(X, Y).Type
End Function

Public Function OutOfBoundries(ByVal X As Long, ByVal Y As Long) As Boolean
OutOfBoundries = False

If (X > map.MaxX Or X < 0 Or Y > map.MaxY Or Y < 0) Then
    OutOfBoundries = True
End If
End Function

Public Function GetNextPositionByRef(ByVal dir As Byte, ByRef X As Long, ByRef Y As Long) As Boolean
    Select Case dir
    Case DIR_UP
        Y = Y - 1
    Case DIR_DOWN
        Y = Y + 1
    Case DIR_LEFT
        X = X - 1
    Case DIR_RIGHT
        X = X + 1
    End Select
    'Check if out of map
    GetNextPositionByRef = OutOfBoundries(X, Y)
End Function




Public Function CalculateItemWeight(ByVal ItemNum As Long) As Long
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function
    
    If IsItemUnitWeight(ItemNum) Then
        CalculateItemWeight = 1
        Exit Function
    End If
    
    Dim Weight As Long
    Weight = 0
    With Item(ItemNum)
        If Item(ItemNum).Type = ITEM_TYPE_CONSUME Then
            Weight = Weight + .AddEXP
            Weight = Weight + .AddHP * ONE_VITAL_WEIGHT
            Weight = Weight + .AddMP * ONE_VITAL_WEIGHT
            Weight = Weight + CalculateItemWeight(.ConsumeItem)
        Else 'sword, shield...
            Weight = Weight + GetEquipableItemStatsSum(ItemNum) * ONE_STAT_ADD_WEIGHT
        End If
    End With
    
    CalculateItemWeight = Weight
            
End Function

Public Function GetEquipableItemStatsSum(ByVal ItemNum As Long) As Long
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function
    
    Dim ret As Long
    ret = 0
    With Item(ItemNum)
    Dim i As Byte
    For i = 1 To Stats.Stat_Count - 1
        ret = ret + .Add_Stat(i)
    Next
    
    ret = ret + .Data2
    
    GetEquipableItemStatsSum = ret
    
    End With
End Function

Public Function GetItemWeight(ByVal ItemNum As Long) As Long
If ItemNum > 0 And ItemNum < MAX_ITEMS Then
    GetItemWeight = Item(ItemNum).Weight
End If
End Function

Public Function IsItemUnitWeight(ByVal ItemNum As Long) As Boolean
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function
    
    Dim ItemType As Long
    ItemType = Item(ItemNum).Type
    
    If ItemType = ITEM_TYPE_KEY Or ItemType = ITEM_TYPE_CURRENCY Or ItemType = ITEM_TYPE_SPELL Or ItemType = ITEM_TYPE_RESET_POINTS Or ItemType = ITEM_TYPE_TRIFORCE Or ItemType = ITEM_TYPE_REDEMPTION Or ItemType = ITEM_TYPE_CONTAINER Or ItemType = ITEM_TYPE_BAG Or ItemType = ITEM_TYPE_NONE Then
        IsItemUnitWeight = True
    End If
End Function

Public Function GetPlayerInvWeight(ByVal index As Long) As Long
If index > 0 And index < MAX_PLAYERS Then
    Dim i As Byte
    Dim res As Currency
    res = 0
    For i = 1 To MAX_INV
        Dim ItemNum As Long
        ItemNum = GetPlayerInvItemNum(index, i)
        If isItemStackable(ItemNum) Then
            res = res + GetItemWeight(ItemNum) * GetPlayerInvItemValue(index, i)
        Else
            res = res + GetItemWeight(ItemNum)
        End If
    Next
    If res <= 0 Then
        res = 1
    End If

    GetPlayerInvWeight = res
End If
End Function
Public Function GetPlayerWeight(ByVal index As Long) As Long
    GetPlayerWeight = GetPlayerEquipmentWeight(index) + GetPlayerInvWeight(index)
End Function
Public Function GetPlayerEquipmentWeight(ByVal index As Long) As Long
If index > 0 And index < MAX_PLAYERS Then
    Dim i As Byte
    Dim res As Long
    For i = 1 To Equipment.Equipment_Count - 1
        res = res + GetItemWeight(GetPlayerEquipment(index, i))
    Next
    
    GetPlayerEquipmentWeight = res
End If
End Function
Public Sub SetPlayerMaxWeight(ByVal index As Long, ByVal Weight As Long)
If index > 0 And index < MAX_PLAYERS Then
   Player(index).MaxWeight = Weight
End If
End Sub
Public Function GetPlayerMaxWeight(ByVal index As Long) As Long
If index > 0 And index < MAX_PLAYERS Then
    GetPlayerMaxWeight = Player(index).MaxWeight
End If
End Function

Public Sub DisplayWeightPercent()
    If GetPlayerMaxWeight(MyIndex) > 0 Then
        frmMain.lblInvWeight.Caption = "Weight: " & CLng(GetPlayerWeight(MyIndex) / GetPlayerMaxWeight(MyIndex) * 100) & "%"
    End If
End Sub

Sub InitChatRooms()
    Dim i As Byte
    For i = 1 To ChatType_Count - 1
       Call ListCreate(ChatRooms(i))
    Next
End Sub

Function StatToStr(ByVal stat As Stats) As String
    Select Case stat
    Case Stats.Strength
        StatToStr = "Strength"
    Case Stats.Endurance
        StatToStr = "Endurance"
    Case Stats.Agility
        StatToStr = "Agility"
    Case Stats.Intelligence
        StatToStr = "Intelligence"
    Case Stats.Willpower
        StatToStr = "WillPower"
    Case Else
        StatToStr = "None"
    End Select
End Function


Function ActionToStr(ByVal PlayerAction As PlayerActionsType) As String
    Select Case PlayerAction
    Case aAttack
        ActionToStr = "Attack"
    Case aMove
        ActionToStr = "Move"
    Case aSpell
        ActionToStr = "Spell"
    Case aUseItem
        ActionToStr = "UseItem"
    Case aTeleport
        ActionToStr = "Teleport"
    Case Else
        ActionToStr = "None"
    End Select
End Function





Function IsMapNPCaPet(ByVal mapnpcnum As Long) As Boolean
    If Not (0 < mapnpcnum <= MAX_MAP_NPCS) Then Exit Function
    
    If MapNpc(mapnpcnum).petData.Owner > 0 Then
        IsMapNPCaPet = True
    End If
End Function


Public Sub CheckMPScroll(ByVal setvalue As Boolean)
Dim prevvalue As Long
If setvalue Then
    If Not (frmEditor_Spell.scrlMP.value <= 100 And 0 <= Spell(EditorIndex).MPCost And Spell(EditorIndex).MPCost <= 100) Then
        frmEditor_Spell.scrlMP.value = 0
        frmEditor_Spell.scrlMP.Min = 0
        frmEditor_Spell.scrlMP.Max = 100
    Else
        prevvalue = Spell(EditorIndex).MPCost
        frmEditor_Spell.scrlMP.Min = 0
        frmEditor_Spell.scrlMP.Max = 100
        frmEditor_Spell.scrlMP.value = prevvalue
    End If
Else
    prevvalue = Spell(EditorIndex).MPCost
    frmEditor_Spell.scrlMP.Min = 0
    frmEditor_Spell.scrlMP.Max = 2000
    frmEditor_Spell.scrlMP.value = prevvalue
End If
End Sub


Function ShopTypeToStr(ByVal ShopType As Byte) As String
    Select Case ShopType
    Case SHItem
        ShopTypeToStr = "Item"
    Case SHHeroKillPoints
        ShopTypeToStr = "Hero Points"
    Case SHPKKillPoints
        ShopTypeToStr = "Assassin Points"
    Case SHQuestPoints
        ShopTypeToStr = "Quest Points"
    Case SHNPCPoints
        ShopTypeToStr = "NPC Points"
    Case SHBonusPoints
        ShopTypeToStr = CURRENCY_NAME
    End Select
End Function

Function GetShopPriceName(ByVal Shopnum As Long, ByVal shopslot As Long) As String
    If Shop(Shopnum).PriceType = SHItem Then
        GetShopPriceName = Item(Shop(Shopnum).TradeItem(shopslot).CostItem).Name
    Else
        GetShopPriceName = ShopTypeToStr(Shop(Shopnum).PriceType)
    End If
    GetShopPriceName = GetShopPriceName
    
End Function


Sub InitCmbPriceType()
    Dim i As Byte
    With frmEditor_Shop
    For i = 0 To ShopPricesTypeCount - 1
        .cmbPriceType.AddItem ShopTypeToStr(i)
    Next
    End With
End Sub

Sub ClearCmbPricetype()
    With frmEditor_Shop
    .cmbPriceType.Clear
    End With
End Sub


Sub SetKillPoints(ByVal Status As Byte, ByVal points As Long)
    With frmMain.lblKillPoints
    Select Case Status
    Case PK_PLAYER
        .Caption = "PK: " & points
        .ForeColor = QBColor(BrightRed)
    Case HERO_PLAYER
        .Caption = "HERO: " & points
        .ForeColor = QBColor(Yellow)
    Case NONE_PLAYER
        .Caption = "Neutral"
        .ForeColor = QBColor(White)
    End Select
    End With
End Sub

Sub SetBonusPoints(ByVal index As Long, ByVal points As Long)
    Player(index).BonusPoints = points
End Sub

Function GetBonusPoints(ByVal index As Long) As Long
    GetBonusPoints = Player(index).BonusPoints
End Function


Sub CheckCustomSpritePosition(ByVal X As Single, ByVal Y As Single)
    If CanSetSpriteByClick Then
        With CustomSprites(EditorIndex)
            .Layers(1).CentersPositions(1).X = Y
            .Layers(1).CentersPositions(1).Y = X
        End With
    
    End If
End Sub

Function CanSetSpriteByClick() As Boolean
    If frmEditor_CustomSprites.Visible = True Then
        
        CanSetSpriteByClick = True
    End If
End Function

Function GetCustomSpriteCenterLayerSprite(ByVal CustomSprite As Byte) As Long
    Dim i As Long
    For i = 1 To CustomSprites(CustomSprite).NLayers
        With CustomSprites(CustomSprite).Layers(i)
            If .UseCenterPosition = False Then
                If .UsePlayerSprite Then
                    GetCustomSpriteCenterLayerSprite = 0
                Else
                    GetCustomSpriteCenterLayerSprite = .sprite
                End If
                Exit Function
            End If
        End With
    Next
End Function

Function GetPlayerCenterSprite(ByVal index As Long)
    Dim sprite As Long
    If GetPlayerCustomSprite(index) > 0 Then
        sprite = GetCustomSpriteCenterLayerSprite(GetPlayerCustomSprite(index))
        If sprite = 0 Then: sprite = GetPlayerSprite(index)
    Else
        sprite = GetPlayerSprite(index)
    End If
    
    GetPlayerCenterSprite = sprite
End Function
Function HasMaxGold() As Boolean
    Dim i As Long, value As Long
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) = 1 Then
            value = value + GetPlayerInvItemValue(MyIndex, i)
        End If
    Next
    
    Dim bags As Long
    bags = Player(MyIndex).RupeeBags
    If bags = MAX_RUPEE_BAGS Then
        HasMaxGold = (value >= bags * BAG_CAPACITY - 1)
    Else
        HasMaxGold = (value >= bags * BAG_CAPACITY)

    End If
End Function
