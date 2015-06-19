Attribute VB_Name = "modMovements"


Public Function ComputeActualMovement(ByRef NPCS As MapNPCPropertiesRec, ByVal MapNum As Long, ByVal MapNPCNum As Long) As Byte
Dim MovementNum As Byte, ListActual As Byte, i As Byte
Dim Absolute As Boolean
MovementNum = NPCS.Movement
ListActual = mapnpc(MapNum).NPC(MapNPCNum).Actual

ComputeActualMovement = 4 'Null movement


If MovementNum <> 0 Then

    
    
    If Movements(MovementNum).Type = 4 Then 'random
        i = RAND(0, 4) ' 4 is used for stopping movement
        If CanRandNPCMove(NPCS, MapNum, MapNPCNum, i) = True Then
            ComputeActualMovement = i
            Exit Function
        End If
        Exit Function
    End If
    
    If ListActual = 0 Then
        Exit Function
    End If
        
    If Movements(MovementNum).MovementsTable.nelem <= 0 Then Exit Function
    
    ComputeActualMovement = GetMovementActualDirection(MovementNum, MapNum, MapNPCNum)
    
    ' This have to be above of another conditional, Order matters
    If EndOfMovement(NPCS, MapNum, MapNPCNum) = True Then
        'Only used for ByTile Movement
        Absolute = AbsoluteEndOfMovement(NPCS, MapNum, MapNPCNum)
        If Absolute = True Then
            Call InvertNpcList(NPCS, MapNum, MapNPCNum)
        Else
            Call NextMovement(NPCS, MapNum, MapNPCNum)
        End If
        ComputeActualMovement = 4
        Exit Function
    ElseIf BlockingEndOfMovement(NPCS, MapNum, MapNPCNum) = True Then
        'Stop Movement, we decide this for lowering  proces complexity
        Absolute = AbsoluteEndOfMovement(NPCS, MapNum, MapNPCNum)
        If Absolute = True Then
            Call InvertNpcList(NPCS, MapNum, MapNPCNum)
        Else
            Call NextMovement(NPCS, MapNum, MapNPCNum)
        End If
        ComputeActualMovement = 4
        Exit Function
    End If
        
End If



End Function

Public Sub InvertNpcList(ByRef NPCS As MapNPCPropertiesRec, ByVal MapNum As Long, ByVal MapNPCNum As Long)

    mapnpc(MapNum).NPC(MapNPCNum).Inverse = Not (mapnpc(MapNum).NPC(MapNPCNum).Inverse)
    mapnpc(MapNum).NPC(MapNPCNum).Count = GetRemainingTiles(Movements(NPCS.Movement).MovementsTable, NPCS, MapNum, MapNPCNum)

End Sub



Public Sub NextMovement(ByRef NPCS As MapNPCPropertiesRec, ByVal MapNum As Long, ByVal MapNPCNum As Long)

Select Case mapnpc(MapNum).NPC(MapNPCNum).Inverse

Case False 'Normal Moviment, we have to stop when list end
        Call NPCSnext(Movements(NPCS.Movement).MovementsTable, NPCS, MapNum, MapNPCNum)
Case True 'Inverted movement, stop when first
        Call NPCSprevious(Movements(NPCS.Movement).MovementsTable, NPCS, MapNum, MapNPCNum)
End Select

If Movements(NPCS.Movement).Type = 3 Then
    mapnpc(MapNum).NPC(MapNPCNum).Count = 0
End If

    

End Sub

Public Sub NPCSnext(ByRef list As MovementsListRec, ByRef NPCS As MapNPCPropertiesRec, ByVal MapNum As Long, ByVal MapNPCNum As Long)
If mapnpc(MapNum).NPC(MapNPCNum).Actual >= list.nelem Then Exit Sub

mapnpc(MapNum).NPC(MapNPCNum).Actual = mapnpc(MapNum).NPC(MapNPCNum).Actual + 1

End Sub

Public Sub NPCSprevious(ByRef list As MovementsListRec, ByRef NPCS As MapNPCPropertiesRec, ByVal MapNum As Long, ByVal MapNPCNum As Long)
If mapnpc(MapNum).NPC(MapNPCNum).Actual <= 1 Then Exit Sub

mapnpc(MapNum).NPC(MapNPCNum).Actual = mapnpc(MapNum).NPC(MapNPCNum).Actual - 1

End Sub

Public Function AbsoluteEndOfMovement(ByRef NPCS As MapNPCPropertiesRec, ByVal MapNum As Long, ByVal MapNPCNum As Long) As Boolean

Dim CanBladeMove As Integer
Dim canmove As Boolean
Dim MovementNum As Byte

MovementNum = NPCS.Movement
AbsoluteEndOfMovement = False

Select Case Movements(MovementNum).Type

Case 1 To 2 'OnlyDirectional and bymovements


        If NPCListEnd(Movements(MovementNum).MovementsTable, NPCS, MapNum, MapNPCNum) = True Then
                AbsoluteEndOfMovement = True
        End If

Case 3 'By tile

    If NPCListEnd(Movements(MovementNum).MovementsTable, NPCS, MapNum, MapNPCNum) = True And GetRemainingTiles(Movements(MovementNum).MovementsTable, NPCS, MapNum, MapNPCNum) <= 0 Then
            AbsoluteEndOfMovement = True
    End If
          
End Select

End Function

Public Function BlockingEndOfMovement(ByRef NPCS As MapNPCPropertiesRec, ByVal MapNum As Long, ByVal MapNPCNum As Long) As Boolean
Dim CanBladeMove As Integer
Dim canmove As Boolean
Dim MovementNum As Byte

MovementNum = NPCS.Movement
BlockingEndOfMovement = False

Select Case Movements(MovementNum).Type

Case 1 To 3

    If NPC(Map(MapNum).NPC(MapNPCNum)).Behaviour = NPC_BEHAVIOUR_BLADE Then
            CanBladeMove = CanBladeNpcMove(MapNum, MapNPCNum, GetMovementActualDirection(MovementNum, MapNum, MapNPCNum))
            If CanBladeMove = -1 Then
                BlockingEndOfMovement = True
            ElseIf CanBladeMove > 0 Then
                'Call action
                Call ParseAction(CanBladeMove, NPCS.Action, 0) 'Computing action on tilematch moment
            End If

    Else
            canmove = CanNpcMove(MapNum, MapNPCNum, GetMovementActualDirection(MovementNum, MapNum, MapNPCNum))
            If canmove = False Then
                BlockingEndOfMovement = True
            End If
    End If

    
       
End Select

End Function

Public Function EndOfMovement(ByRef NPCS As MapNPCPropertiesRec, ByVal MapNum As Long, ByVal MapNPCNum As Long) As Boolean

Dim CanBladeMove As Integer
Dim canmove As Boolean
Dim MovementNum As Byte

MovementNum = NPCS.Movement
EndOfMovement = False

Select Case Movements(MovementNum).Type

Case 3

    If GetRemainingTiles(Movements(MovementNum).MovementsTable, NPCS, MapNum, MapNPCNum) <= 0 Then
            EndOfMovement = True
    Else
            mapnpc(MapNum).NPC(MapNPCNum).Count = mapnpc(MapNum).NPC(MapNPCNum).Count + 1
    End If
          
End Select

End Function

Public Function NPCListEnd(ByRef list As MovementsListRec, ByRef NPCS As MapNPCPropertiesRec, ByVal MapNum As Long, ByVal MapNPCNum As Long) As Boolean

NPCListEnd = False

Select Case mapnpc(MapNum).NPC(MapNPCNum).Inverse
Case False
    'Normal Moviment return true when list ends
    If list.nelem = mapnpc(MapNum).NPC(MapNPCNum).Actual Then
        NPCListEnd = True
    End If
Case True
    If mapnpc(MapNum).NPC(MapNPCNum).Actual = 1 Then
        NPCListEnd = True
    End If
End Select


End Function

Public Function GetRemainingTiles(ByRef list As MovementsListRec, ByRef NPCS As MapNPCPropertiesRec, ByVal MapNum As Long, ByVal MapNPCNum As Long) As Byte
GetRemainingTiles = list.vect(mapnpc(MapNum).NPC(MapNPCNum).Actual).Data.NumberOfTiles - mapnpc(MapNum).NPC(MapNPCNum).Count
End Function

Public Function CanRandNPCMove(ByRef NPCS As MapNPCPropertiesRec, ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal i As Byte) As Boolean
Dim CanBladeMove As Integer
Dim canmove As Boolean

    If NPC(Map(MapNum).NPC(MapNPCNum)).Behaviour = NPC_BEHAVIOUR_BLADE Then
        CanBladeMove = CanBladeNpcMove(MapNum, MapNPCNum, i)
        If CanBladeMove = -1 Then
            CanRandNPCMove = False
        ElseIf CanBladeMove > 0 Then
            'Compute player action (canblademove)
            Call ParseAction(CanBladeMove, NPCS.Action, 0) 'Computing action on tilematch moment
            CanRandNPCMove = True
        ElseIf CanBladeMove = 0 Then
            CanRandNPCMove = True
        End If
    Else
        canmove = CanNpcMove(MapNum, MapNPCNum, i)
        If canmove = False Then
            CanRandNPCMove = False
        Else
            CanRandNPCMove = True
        End If
    End If
        
End Function



Public Function GetMovementActualDirection(ByVal MovementNum As Byte, ByVal MapNum As Long, ByVal MapNPCNum As Long) As Byte
Dim i As Integer
Dim j As Byte

i = BTI(mapnpc(MapNum).NPC(MapNPCNum).Inverse)
j = Movements(MovementNum).MovementsTable.vect(mapnpc(MapNum).NPC(MapNPCNum).Actual).Data.Direction

If j >= 0 And j <= 3 Then
    If j = 1 Or j = 3 Then
        j = j - i
    Else
        j = j + i
    End If
    
    GetMovementActualDirection = j
    Exit Function
End If

GetMovementActualDirection = 4

End Function

Public Function EndOfMovementTiles(ByVal MovementNum As Byte, ByVal MapNum As Long, ByVal MapNPCNum As Long) As Boolean

EndOfMovementTiles = False
If mapnpc(MapNum).NPC(MapNPCNum).Count > Movements(MovementNum).MovementsTable.vect(Movements(MovementNum).MovementsTable.Actual).Data.NumberOfTiles Then
    EndOfMovementTiles = True
    Exit Function
End If

End Function

Public Sub InvertMovementTiles(ByVal MovementNum As Byte, ByVal MapNum As Long, ByVal MapNPCNum As Long)

mapnpc(MapNum).NPC(MapNPCNum).Count = Movements(MovementNum).MovementsTable.vect(Movements(MovementNum).MovementsTable.Actual).Data.NumberOfTiles - mapnpc(MapNum).NPC(MapNPCNum).Count

End Sub
Public Sub NextMovementTiles(ByVal MovementNum As Byte, ByVal MapNum As Long, ByVal MapNPCNum As Long)

mapnpc(MapNum).NPC(MapNPCNum).Count = mapnpc(MapNum).NPC(MapNPCNum).Count + 1

End Sub

Public Function BTI(ByVal Var As Boolean) As Long

If Var Then
    BTI = 1
Else
    BTI = 0
End If

End Function

Sub ResetMapNPCSProperties(ByVal Movement As Byte)
Dim i, n As Long
For i = 1 To MAX_MAPS
    For n = 1 To MAX_MAP_NPCS
        If Map(i).NPCSProperties(n).Movement = Movement Then
            mapnpc(i).NPC(n).Actual = 1
            mapnpc(i).NPC(n).Count = 0
            mapnpc(i).NPC(n).Inverse = False
        End If
    Next
Next
End Sub

Sub ResetMapNPCMovement(ByVal MapNum As Long, ByVal MapNPCNum As Long)
If MapNum > 0 And MapNum <= MAX_MAPS And MapNPCNum > 0 And MapNPCNum <= MAX_MAP_NPCS Then
    mapnpc(MapNum).NPC(MapNPCNum).Actual = 1
    mapnpc(MapNum).NPC(MapNPCNum).Count = 0
    mapnpc(MapNum).NPC(MapNPCNum).Inverse = False
End If

End Sub


