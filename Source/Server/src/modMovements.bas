Attribute VB_Name = "modMovements"
Option Explicit


Public Function GetNextMovementDir(ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long) As Byte
Dim MovementNum As Byte
Dim ListActual As Byte
Dim i As Byte
Dim Absolute As Boolean
Dim Inverse As Boolean

MovementNum = npcs.Movement
ListActual = MapNpc(mapnum).NPC(mapnpcnum).Actual
Inverse = MapNpc(mapnum).NPC(mapnpcnum).Inverse

GetNextMovementDir = 4 'Null movement

If MovementNum = 0 Then Exit Function

If ListActual = 0 Then Exit Function



Select Case GetMovementType(MovementNum)

Case MovementType.Random
    i = RAND(0, 4) ' 4 is used for stopping movement
    If Not TestNPCMove(npcs, mapnum, mapnpcnum, i) Then
        i = 4
    End If
    
Case MovementType.OnlyDirectional
    If Movements(MovementNum).MovementsTable.nelem <= 0 Then Exit Function
    i = GetMovementActualDir(mapnum, mapnpcnum, MovementNum)
    If Inverse Then InvertDir i
    Call ProcessOnlyDirectionalMovement(npcs, mapnum, mapnpcnum, i)
    
Case MovementType.Bydirection
    If Movements(MovementNum).MovementsTable.nelem <= 0 Then Exit Function
    i = GetMovementActualDir(mapnum, mapnpcnum, MovementNum)
    If Inverse Then InvertDir i
    Call ProcessByDirectionMovement(npcs, mapnum, mapnpcnum, i)
    
Case MovementType.ByTile
    If Movements(MovementNum).MovementsTable.nelem <= 0 Then Exit Function
    i = GetMovementActualDir(mapnum, mapnpcnum, MovementNum)
    If Inverse Then InvertDir i
    Call ProcessByTileMovement(npcs, mapnum, mapnpcnum, i, MovementNum)
End Select

GetNextMovementDir = i

End Function

Private Function GetMovementRepeat(ByRef MovementNum As Byte) As Boolean
    If Movements(MovementNum).Repeat = True Then
        GetMovementRepeat = True
    End If
End Function

Private Sub ProcessOnlyDirectionalMovement(ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long, ByRef dir As Byte)
    Dim CanMove As Boolean
    CanMove = TestNPCMove(npcs, mapnum, mapnpcnum, dir)
    If Not CanMove Then
        dir = 4
        Call InvertNpcList(npcs, mapnum, mapnpcnum)
    End If
End Sub

Private Sub ProcessByDirectionMovement(ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long, ByRef dir As Byte)
    Dim CanMove As Boolean
    CanMove = TestNPCMove(npcs, mapnum, mapnpcnum, dir)
    If Not CanMove Then
        If EndOfMovement(npcs, mapnum, mapnpcnum, npcs.Movement) Then
            If GetMovementRepeat(npcs.Movement) Then
                Call NPCSfirst(Movements(npcs.Movement).MovementsTable, npcs, mapnum, mapnpcnum)
            Else
                Call InvertNpcList(npcs, mapnum, mapnpcnum)
            End If
        Else
           Call NextMovement(npcs, mapnum, mapnpcnum)
        End If
        dir = 4
    End If
End Sub

Private Sub ProcessByTileMovement(ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long, ByRef dir As Byte, ByVal MovementNum As Long)
    Dim CanMove As Boolean
    If GetRemainingTiles(Movements(MovementNum).MovementsTable, npcs, mapnum, mapnpcnum) > 0 Then
        CanMove = TestNPCMove(npcs, mapnum, mapnpcnum, dir)
        If CanMove Then
            MakeStep MovementNum, mapnum, mapnpcnum
        Else
            InvertNpcList npcs, mapnum, mapnpcnum
            dir = 4
        End If
    Else
        If EndOfMovement(npcs, mapnum, mapnpcnum, MovementNum) Then
            dir = 4
            If GetMovementRepeat(npcs.Movement) Then
                Call NPCSfirst(Movements(npcs.Movement).MovementsTable, npcs, mapnum, mapnpcnum)
            Else
                Call InvertNpcList(npcs, mapnum, mapnpcnum)
            End If
        Else
            NextMovement npcs, mapnum, mapnpcnum
            dir = 4
        End If
    End If
End Sub

Private Function TestNPCMove(ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal dir As Byte)

    If NPC(MapNpc(mapnum).NPC(mapnpcnum).Num).Behaviour = NPC_BEHAVIOUR_BLADE Then
        Dim CanBladeMove As Integer
        CanBladeMove = CanBladeNpcMove(mapnum, mapnpcnum, dir)
        If CanBladeMove = 0 Then
            TestNPCMove = True
        ElseIf CanBladeMove > 0 Then
            TestNPCMove = True
            Call ParseAction(CanBladeMove, npcs.Action) 'Computing action on tilematch moment
        End If
    Else
        TestNPCMove = CanNpcMove(mapnum, mapnpcnum, dir)
    End If
    
End Function

Public Sub InvertNpcList(ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long)

    MapNpc(mapnum).NPC(mapnpcnum).Inverse = Not (MapNpc(mapnum).NPC(mapnpcnum).Inverse)
    MapNpc(mapnum).NPC(mapnpcnum).Count = GetRemainingTiles(Movements(npcs.Movement).MovementsTable, npcs, mapnum, mapnpcnum)

End Sub



Public Sub NextMovement(ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long)

Select Case MapNpc(mapnum).NPC(mapnpcnum).Inverse

Case False 'Normal Moviment, we have to stop when list end
        Call NPCSnext(Movements(npcs.Movement).MovementsTable, npcs, mapnum, mapnpcnum)
Case True 'Inverted movement, stop when first
        Call NPCSprevious(Movements(npcs.Movement).MovementsTable, npcs, mapnum, mapnpcnum)
End Select

If Movements(npcs.Movement).Type = MovementType.ByTile Then
    MapNpc(mapnum).NPC(mapnpcnum).Count = 0
End If

    

End Sub

Public Sub NPCSnext(ByRef List As MovementsListRec, ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long)
If MapNpc(mapnum).NPC(mapnpcnum).Actual >= List.nelem Then Exit Sub

MapNpc(mapnum).NPC(mapnpcnum).Actual = MapNpc(mapnum).NPC(mapnpcnum).Actual + 1

End Sub

Public Sub NPCSprevious(ByRef List As MovementsListRec, ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long)
If MapNpc(mapnum).NPC(mapnpcnum).Actual <= 1 Then Exit Sub

MapNpc(mapnum).NPC(mapnpcnum).Actual = MapNpc(mapnum).NPC(mapnpcnum).Actual - 1

End Sub

Public Sub NPCSfirst(ByRef List As MovementsListRec, ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long)


MapNpc(mapnum).NPC(mapnpcnum).Actual = 1
MapNpc(mapnum).NPC(mapnpcnum).Count = 0
MapNpc(mapnum).NPC(mapnpcnum).Inverse = False

End Sub

Public Function EndOfMovement(ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal MovementNum As Long) As Boolean

EndOfMovement = False

Select Case GetMovementType(MovementNum)
Case MovementType.OnlyDirectional
    EndOfMovement = True
Case MovementType.Bydirection
    If NPCListEnd(Movements(MovementNum).MovementsTable, npcs, mapnum, mapnpcnum) Then
        EndOfMovement = True
    End If
Case MovementType.ByTile
    If NPCListEnd(Movements(MovementNum).MovementsTable, npcs, mapnum, mapnpcnum) Then
        If GetRemainingTiles(Movements(MovementNum).MovementsTable, npcs, mapnum, mapnpcnum) <= 0 Then
            EndOfMovement = True
        End If
    End If
End Select

End Function

Public Function NPCListEnd(ByRef List As MovementsListRec, ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long) As Boolean

NPCListEnd = False

Select Case MapNpc(mapnum).NPC(mapnpcnum).Inverse
Case False
    'Normal Moviment return true when list ends
    If List.nelem = MapNpc(mapnum).NPC(mapnpcnum).Actual Then
        NPCListEnd = True
    End If
Case True
    If MapNpc(mapnum).NPC(mapnpcnum).Actual = 1 Then
        NPCListEnd = True
    End If
End Select


End Function

Public Function GetRemainingTiles(ByRef List As MovementsListRec, ByRef npcs As MapNPCPropertiesRec, ByVal mapnum As Long, ByVal mapnpcnum As Long) As Byte
GetRemainingTiles = List.vect(MapNpc(mapnum).NPC(mapnpcnum).Actual).Data.NumberOfTiles - MapNpc(mapnum).NPC(mapnpcnum).Count
End Function

Public Sub MakeStep(ByVal MovementNum As Byte, ByVal mapnum As Long, ByVal mapnpcnum As Long)

MapNpc(mapnum).NPC(mapnpcnum).Count = MapNpc(mapnum).NPC(mapnpcnum).Count + 1

End Sub

Public Function GetMovementType(ByVal MovementNum As Long) As MovementType
    If MovementNum > 0 And MovementNum < MAX_MOVEMENTS Then
        GetMovementType = Movements(MovementNum).Type
    End If
End Function

Private Function GetMovementActualDir(ByVal mapnum As Long, ByVal mapnpcnum As Long, ByVal MovementNum As Long)
    GetMovementActualDir = Movements(MovementNum).MovementsTable.vect(MapNpc(mapnum).NPC(mapnpcnum).Actual).Data.Direction
End Function

Private Sub InvertDir(ByRef dir As Byte)
    If dir = DIR_UP Then
        dir = DIR_DOWN
    ElseIf dir = DIR_RIGHT Then
        dir = DIR_LEFT
    ElseIf dir = DIR_DOWN Then
        dir = DIR_UP
    ElseIf dir = DIR_LEFT Then
        dir = DIR_RIGHT
    End If
End Sub



Sub ResetMapNPCSProperties(ByVal Movement As Byte)
Dim i, N As Long
For i = 1 To MAX_MAPS
    For N = 1 To MAX_MAP_NPCS
        If map(i).NPCSProperties(N).Movement = Movement Then
            MapNpc(i).NPC(N).Actual = 1
            MapNpc(i).NPC(N).Count = 0
            MapNpc(i).NPC(N).Inverse = False
        End If
    Next
Next
End Sub

Sub ResetMapNPCMovement(ByVal mapnum As Long, ByVal mapnpcnum As Long)
If mapnum > 0 And mapnum <= MAX_MAPS And mapnpcnum > 0 And mapnpcnum <= MAX_MAP_NPCS Then
    MapNpc(mapnum).NPC(mapnpcnum).Actual = 1
    MapNpc(mapnum).NPC(mapnpcnum).Count = 0
    MapNpc(mapnum).NPC(mapnpcnum).Inverse = False
End If

End Sub



