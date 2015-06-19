Attribute VB_Name = "modDoor"

Public Type DoorRec
    Name As String * NAME_LENGTH
    DoorType As Long
    
    WarpMap As Long
    WarpX As Long
    WarpY As Long
    
    UnlockType As Long
    key As Long
    Switch As Long
    
    Time As Long
    
    InitialState As Boolean
End Type


Public Doors(1 To MAX_DOORS) As DoorRec



Sub CheckDoor(ByVal index As Long, ByVal X As Long, ByVal Y As Long)
    Dim Door_Num As Long
    Dim i As Long
    Dim n As Long
    Dim key As Long
    Dim tmpIndex As Long
    
    If OutOfBoundries(X, Y, GetPlayerMap(index)) Then Exit Sub
    
    If map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_DOOR Then
        Door_Num = map(GetPlayerMap(index)).Tile(X, Y).Data1
        Dim TempDoorNum As Long
        TempDoorNum = GetTempDoorNumberByTile(GetPlayerMap(index), X, Y)
        
        If Door_Num > 0 Then
            If Doors(Door_Num).DoorType = 0 Then
                If Not IsDoorOpened(GetPlayerMap(index), TempDoorNum) Then
                    If Doors(Door_Num).UnlockType = 0 Then
                        For i = 1 To MAX_INV
                            key = GetPlayerInvItemNum(index, i)
                            If Doors(Door_Num).key = key Then
                                SetAllMapDoorNum GetPlayerMap(index), Door_Num
                                PlayerMsg n, "Se ha desbloqueado algo", Cyan
                                SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSwitch, 1
                                ' End
                                Exit Sub
                            End If
                        Next
                        PlayerMsg index, "No posees la llave adecuada para abrir la puerta.", Cyan
                    ElseIf Doors(Door_Num).UnlockType = 1 Then
                        If Not TempTile(GetPlayerMap(index)).Door(TempDoorNum).state Then
                            PlayerMsg index, "No has pulsado el interruptor adecuado para abrir la puerta.", Cyan
                        End If
                    ElseIf Doors(Door_Num).UnlockType = 2 Or Doors(Door_Num).UnlockType = 3 Then
                        PlayerMsg index, "Ésta puerta está cerrada.", Cyan
                    End If
                    
                Else
                    PlayerMsg index, "Ésta puerta ya está abierta.", Cyan
                End If
            ElseIf Doors(Door_Num).DoorType = 1 Then
                Dim Switch As Long
                If Not IsDoorOpened(GetPlayerMap(index), TempDoorNum) Then 'checking if switch is activated
                    TempTile(GetPlayerMap(index)).Door(TempDoorNum).state = True
                    
                    If (Doors(Door_Num).Switch) > 0 Then
                        Switch = GetTempDoorNumberByDoorNum(GetPlayerMap(index), Doors(Door_Num).Switch)
                        If Switch > 0 Then
                            SetAllMapDoorNum GetPlayerMap(index), Doors(Door_Num).Switch
                            SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSwitch, 1
                        End If
                        TempTile(GetPlayerMap(index)).Door(TempDoorNum).DoorTimer = GetDoorLockTime(Doors(Door_Num).Switch)
                    End If
                Else
                    TempTile(GetPlayerMap(index)).Door(TempDoorNum).state = False
                    If (Doors(Door_Num).Switch) > 0 Then
                        Switch = GetTempDoorNumberByDoorNum(GetPlayerMap(index), Doors(Door_Num).Switch)
                        If Switch > 0 Then
                            SetAllMapDoorNum GetPlayerMap(index), Doors(Door_Num).Switch
                        End If
                        PlayerMsg n, "El interruptor ha sido desactivado", Cyan
                        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSwitch, 1
                    End If
                End If
            End If
        End If
    End If
End Sub
Sub SetAllMapDoorNum(ByVal mapnum As Long, ByVal Door_Num As Long)
    If Door_Num = 0 Then Exit Sub
    Dim i As Long
    For i = 1 To TempTile(mapnum).NumDoors
        If TempTile(mapnum).Door(i).doornum = Door_Num Then
            TempTile(mapnum).Door(i).state = Not TempTile(mapnum).Door(i).state
            If TempTile(mapnum).Door(i).state Then
                TempTile(mapnum).Door(i).DoorTimer = GetDoorLockTime(Door_Num)
            Else
                TempTile(mapnum).Door(i).DoorTimer = 0
            End If
            SendMapKeyToMap mapnum, TempTile(mapnum).Door(i).X, TempTile(mapnum).Door(i).Y, TempTile(mapnum).Door(i).state
        End If
    Next
End Sub


Function GetDoorLockTime(ByVal doornum As Long) As Long
    If doornum < 1 Or doornum > MAX_DOORS Then Exit Function
    
    If Doors(doornum).Time = 0 Then
        GetDoorLockTime = 0
    Else
        GetDoorLockTime = GetRealTickCount + Doors(doornum).Time * 1000
    End If
End Function
Function GetTempDoorNumberByTile(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long) As Integer
    Dim i As Integer
    
    If OutOfBoundries(X, Y, mapnum) Then Exit Function
    
    If map(mapnum).Tile(X, Y).Type <> TILE_TYPE_DOOR And map(mapnum).Tile(X, Y).Type <> TILE_TYPE_KEY Then Exit Function
    
    i = BinarySearchDoor(mapnum, 1, TempTile(mapnum).NumDoors, X, Y)
    If i > 0 Then
        If TempTile(mapnum).Door(i).X = X And TempTile(mapnum).Door(i).Y = Y Then
            GetTempDoorNumberByTile = i
            Exit Function
        End If
    End If
End Function

Public Function BinarySearchDoor(ByVal mapnum As Long, ByVal left As Long, ByVal right As Long, ByVal X As Long, ByVal Y As Long) As Long
    If right < left Then
        BinarySearchDoor = 0
    Else
        Dim meddle As Integer
        meddle = (left + right) \ 2
        
        With TempTile(mapnum).Door(meddle)
        
        Dim Ordenation As Integer
        Ordenation = PosOrdenation(X, Y, .X, .Y)
        If Ordenation = 1 Then
            BinarySearchDoor = BinarySearchDoor(mapnum, left, meddle - 1, X, Y)
        ElseIf Ordenation = -1 Then
            BinarySearchDoor = BinarySearchDoor(mapnum, meddle + 1, right, X, Y)
        Else
            BinarySearchDoor = meddle
        End If
        
        End With
    End If
        
        
End Function

Function GetTempDoorNumberByDoorNum(ByVal mapnum As Long, ByVal Door_Num As Long) As Long
    Dim i As Integer
    For i = 1 To TempTile(mapnum).NumDoors
        If TempTile(mapnum).Door(i).doornum = Door_Num Then
            GetTempDoorNumberByDoorNum = i
            Exit Function
        End If
    Next
            
End Function


Function IsDoorOpened(ByVal mapnum As Long, ByVal TempDoorNum As Long) As Boolean
    If TempDoorNum < 1 Or TempDoorNum > TempTile(mapnum).NumDoors Then Exit Function
    
    If TempTile(mapnum).Door(TempDoorNum).state Then
        IsDoorOpened = True
    End If
End Function

Function CanRenderTempDoor(ByVal mapnum As Long, ByVal TempDoorNum As Long) As Boolean
    If mapnum = 0 Or TempDoorNum = 0 Then Exit Function
    Dim doornum As Long
    doornum = TempTile(mapnum).Door(TempDoorNum).doornum
    If doornum > 0 Then
        If Doors(doornum).DoorType = DOOR_TYPE_DOOR Then
            CanRenderTempDoor = True
        End If
    ElseIf doornum = -1 Then
        CanRenderTempDoor = True
    End If
End Function

Function GetInitialDoorState(ByVal doornum As Long) As Byte
    If doornum < 1 Or doornum > MAX_DOORS Then Exit Function
    GetInitialDoorState = Doors(doornum).InitialState
End Function

Function GetDoorType(ByVal doornum As Long) As Byte
    If doornum < 1 Or doornum > MAX_DOORS Then Exit Function
    GetDoorType = Doors(doornum).DoorType
End Function

Sub ChangeAllMapDoorNum(ByVal mapnum As Long, ByVal doornum As Long)
    Dim i As Long
    For i = 1 To TempTile(mapnum).NumDoors
        If TempTile(mapnum).Door(i).doornum = doornum Then
            TempTile(mapnum).Door(i).state = Not (TempTile(mapnum).Door(i).state)
            TempTile(mapnum).Door(i).DoorTimer = 0
            SendMapKeyToMap mapnum, TempTile(mapnum).Door(i).X, TempTile(mapnum).Door(i).Y, TempTile(mapnum).Door(i).state
        End If
    Next
End Sub

Sub ChangeWeightSwitchState(ByVal mapnum As Long, ByVal TempDoorNum As Long)
    With TempTile(mapnum).Door(TempDoorNum)
    .state = Not (.state)
    Dim Switch As Long
    Switch = Doors(.doornum).Switch
    Call ChangeAllMapDoorNum(mapnum, Switch)
    SendSoundToMap mapnum, .X, .Y, seSwitch, 1
    End With
End Sub

Sub CheckWeightSwitch(ByVal mapnum As Long, ByVal TempDoorNum As Long)
    If TempDoorNum > 0 Then
        Dim doornum As Long
        doornum = TempTile(mapnum).Door(TempDoorNum).doornum
        If GetDoorType(doornum) = DOOR_TYPE_WEIGHTSWITCH Then
            Call ChangeWeightSwitchState(mapnum, TempDoorNum)
        End If
    End If
End Sub

Function IsSomebodyOnSwitch(ByVal mapnum As Long, ByVal TempDoorNum As Long) As Boolean

If mapnum = 0 Or TempDoorNum = 0 Then Exit Function
With TempTile(mapnum).Door(TempDoorNum)

If GetMapRefNPCNumByTile(GetMapRef(mapnum), .X, .Y) > 0 Then
    IsSomebodyOnSwitch = True
    Exit Function
End If


If FindPlayerByPos(mapnum, .X, .Y) > 0 Then
    IsSomebodyOnSwitch = True
    Exit Function
End If


End With

End Function

Function IsTempDoorWalkable(ByVal mapnum As Long, ByVal TempDoorNum As Long) As Boolean
    If Not mapnum > 0 And TempDoorNum > 0 Then Exit Function
    With TempTile(mapnum).Door(TempDoorNum)
    Select Case GetDoorType(.doornum)
    Case DOOR_TYPE_DOOR
        IsTempDoorWalkable = IsDoorOpened(mapnum, TempDoorNum)
    Case DOOR_TYPE_SWITCH
        IsTempDoorWalkable = False
    Case DOOR_TYPE_WEIGHTSWITCH
        IsTempDoorWalkable = True
    End Select
    End With
End Function


Function YEStoNO(ByVal i As Byte) As Byte
    If i = YES Then
        YEStoNO = NO
    Else
        YEStoNO = YES
    End If
End Function

