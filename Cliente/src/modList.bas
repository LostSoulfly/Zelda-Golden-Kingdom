Attribute VB_Name = "modMovementsList"
Public Sub LMCreate(ByRef list As MovementsListRec)
Call LMClear(list)
'Create a list: erase previous vector (redimensionate to 1) and set 1 element to it (bad programming method but necessary with the choosen interface)
list.Actual = 0
list.nelem = 0
End Sub

Public Sub LMAdd(ByRef list As MovementsListRec, ByRef element As SingularMovementRec)
Dim i As Byte


If LMredimensionate(list, 1) = False Then Exit Sub

i = list.nelem
'Change the position of the elements
Do While i > list.Actual
    list.vect(i + 1) = list.vect(i)
    i = i - 1
Loop

'add the element after actual, set actual to new element
list.vect(list.Actual + 1).Data = element
list.nelem = list.nelem + 1

End Sub


Public Sub LMDelete(ByRef list As MovementsListRec)
Dim i As Byte

If list.Actual > list.nelem Or LMempty(list) = True Then Exit Sub

'move the elements
For i = list.Actual To list.nelem - 1
    list.vect(i) = list.vect(i + 1)
Next

'change array size, table will be redimensionated
list.nelem = list.nelem - 1
If LMredimensionate(list, 0) = False Then Exit Sub



End Sub

Public Function LMempty(ByRef list As MovementsListRec) As Boolean

If list.nelem = 0 Then
    LMempty = True
Else
    LMempty = False
End If

End Function

Public Function LMGet(ByRef list As MovementsListRec) As SingularMovementRec

If list.Actual > list.nelem Then
    Exit Function
Else
    LMGet = list.vect(list.Actual).Data
End If

End Function

Public Function LMGetByPosition(ByRef list As MovementsListRec, ByVal pos As Byte) As SingularMovementRec

If pos > list.nelem Or pos < 1 Then Exit Function

LMGetByPosition = list.vect(pos).Data

End Function

Public Sub LMFirst(ByRef list As MovementsListRec)
If LMempty(list) = True Then Exit Sub
list.Actual = 1
End Sub
Public Sub LMNext(ByRef list As MovementsListRec)
If Not (LMempty(list)) Then
    list.Actual = list.Actual + 1
End If
End Sub

Public Sub LMPrevious(ByRef list As MovementsListRec)
If Not (list.Actual = 1) Or Not (LMempty(list)) Then
    list.Actual = list.Actual - 1
End If
End Sub

Public Function LMredimensionate(ByRef list As MovementsListRec, ByVal i As Integer) As Boolean

If list.nelem + i < 1 Or list.nelem + i > MAX_MOVEMENT_MOVEMENTS Then
    LMredimensionate = False
Else
   ReDim Preserve list.vect(1 To list.nelem + i)
   LMredimensionate = True
End If

End Function

Public Sub LMClear(ByRef list As MovementsListRec)
Call LMFirst(list)
Do While Not (LMEnd(list))
    Call LMDelete(list)
Loop


End Sub

Public Function LMEnd(ByRef list As MovementsListRec)
    LMEnd = (list.Actual > list.nelem Or LMempty(list) = True)
End Function

Public Sub LMModify(ByRef list As MovementsListRec, ByRef element As SingularMovementRec)

If list.Actual > list.nelem Or LMempty(list) = True Then
    Exit Sub
Else
    list.vect(list.Actual).Data = element
End If

End Sub

Public Sub MergeElements(ByRef list As MovementsListRec, ByVal elements As Byte)


'actual counts for the merging
Dim i As Byte
Dim AuxiliarElement As SingularMovementRec
i = 1

Do While i < elements
    AuxiliarElement = LMGet(list)
    Call LMPrevious(list)
    AuxiliarElement = CombineElements(AuxiliarElement, LMGet(list))
    Call LMDelete(list)
    Call LMModify(list, AuxiliarElement)
i = i + 1
Loop

End Sub

Public Sub LMOptimize(ByRef list As MovementsListRec)
Dim i As Byte, Comparing As Boolean
Dim AuxiliarElement As SingularMovementRec

Comparing = False
i = 0
Call LMFirst(list)

Do While Not (LMEnd(list))

    AuxiliarElement = LMGet(list)
    i = 1
    Call LMNext(list)
    
    Do While AuxiliarElement.direction = LMGet(list).direction And Not (LMEnd(list))
        i = i + 1
        AuxiliarElement = LMGet(list)
        Call LMNext(list)
    Loop
    If i > 1 Then
        Call LMPrevious(list)
        Call MergeElements(list, i)
    End If

Loop


End Sub

Public Function CombineElements(ByRef element1 As SingularMovementRec, ByRef element2 As SingularMovementRec) As SingularMovementRec

If element1.direction <> element2.direction Or element1.NumberOfTiles + element2.NumberOfTiles > 255 Then Exit Function

CombineElements.direction = element1.direction
CombineElements.NumberOfTiles = element1.NumberOfTiles + element2.NumberOfTiles

End Function

Public Sub LMDeleteNulls(ByRef list As MovementsListRec)

Call LMFirst(list)

Do While Not (LMEnd(list))
    
    If LMGet(list).direction = 4 Then 'NONE DIRECTION
        Call LMDelete(list)
    Else
        Call LMNext(list)
    End If

Loop

End Sub



