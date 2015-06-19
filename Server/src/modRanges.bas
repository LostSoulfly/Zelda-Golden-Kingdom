Attribute VB_Name = "modRanges"
Private Type TilePosRec
    X As Byte
    Y As Byte
End Type

Public Function GetAngleByPos(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Single
Dim m As Double
If (x2 - x1) = 0 Then
    If -(y2 - y1) < 0 Then
        GetAngleByPos = 270
    Else
        GetAngleByPos = 90
    End If
Else
    m = -(y2 - y1) / (x2 - x1)
    GetAngleByPos = CSng(Atn(m) * 180 / 3.14159265358979)
    
    If (x2 - x1) < 0 Then
        GetAngleByPos = GetAngleByPos + 180
    End If
    If GetAngleByPos < 0 Then
        GetAngleByPos = GetAngleByPos + 360
    End If
End If


End Function

Function GetDirByCollindantPos(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Byte
    Dim DifX As Long, DifY As Long
    DifX = x1 - x2
    DifY = y1 - y2
    
    If DifX = 0 Then
        If DifY = 1 Then
            GetDirByCollindantPos = DIR_UP
        ElseIf DifY = -1 Then
            GetDirByCollindantPos = DIR_DOWN
        End If
    ElseIf DifY = 0 Then    'difx = -1 or difx = 1
        If DifX = 1 Then
            GetDirByCollindantPos = DIR_LEFT
        ElseIf DifX = -1 Then
            GetDirByCollindantPos = DIR_RIGHT
        End If
    End If
End Function

Function GetDirByPos(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Byte
    GetDirByPos = GetDirByAngle(GetAngleByPos(x1, y1, x2, y2))
End Function

Function GetSecondChanceDir(ByVal FirstDir As Byte, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Byte
    If x1 = x2 Then
        If RAND(0, 1) = 1 Then
            GetSecondChanceDir = DIR_RIGHT
        Else
            GetSecondChanceDir = DIR_LEFT
        End If
    ElseIf y1 = y2 Then
        If RAND(0, 1) = 1 Then
            GetSecondChanceDir = DIR_UP
        Else
            GetSecondChanceDir = DIR_DOWN
        End If
    Else
        
    End If
End Function

Public Function IsAngleInDir(ByVal dir As Byte, ByVal Angle As Single) As Boolean
Select Case dir
Case DIR_UP
    If Angle >= 45 And Angle <= 135 Then
        IsAngleInDir = True
    End If
Case DIR_DOWN
    If Angle >= 225 And Angle <= 315 Then
        IsAngleInDir = True
    End If
Case DIR_LEFT
    If Angle >= 135 And Angle <= 225 Then
        IsAngleInDir = True
    End If
Case DIR_RIGHT
    If Angle >= 315 Or Angle <= 45 Then
        IsAngleInDir = True
    End If
End Select
End Function

Public Function GetDirByAngle(ByVal Angle As Single) As Byte
    If Angle >= 45 And Angle <= 135 Then
        GetDirByAngle = DIR_UP
    ElseIf Angle >= 225 And Angle <= 315 Then
         GetDirByAngle = DIR_DOWN
    ElseIf Angle >= 135 And Angle <= 225 Then
         GetDirByAngle = DIR_LEFT
    ElseIf Angle >= 315 Or Angle <= 45 Then
         GetDirByAngle = DIR_RIGHT
    End If
End Function

Function IsTileAnObstacle(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    IsTileAnObstacle = map(mapnum).Tile(X, Y).Type = TILE_TYPE_BLOCKED
End Function


Sub GetSecondLinearPath(ByVal FirstDir As Byte, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByRef NextX As Long, ByRef NextY As Long)
    Select Case FirstDir
    
    Case DIR_UP, DIR_DOWN
        NextY = y1
        If x1 < x2 Then
            NextX = x1 + 1
        ElseIf x1 > x2 Then
            NextX = x1 - 1
        Else
            NextX = x1 + RAND2(1, -1)
        End If
    Case DIR_LEFT, DIR_RIGHT
        NextX = x1
        If y1 < y2 Then
            NextY = y1 + 1
        ElseIf y2 > y1 Then
            NextY = y1 - 1
        Else
            NextY = y1 + RAND2(1, -1)
        End If
    End Select
End Sub

Function GetSecondDir(ByVal FirstDir As Byte, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Byte
    Select Case FirstDir
    
    Case DIR_UP, DIR_DOWN
        If x1 < x2 Then
            GetSecondDir = DIR_RIGHT
        ElseIf x1 > x2 Then
            GetSecondDir = DIR_LEFT
        Else
            GetSecondDir = RAND2(DIR_RIGHT, DIR_LEFT)
        End If
    Case DIR_LEFT, DIR_RIGHT
        If y1 < y2 Then
            GetSecondDir = DIR_DOWN
        ElseIf y2 > y1 Then
            GetSecondDir = DIR_UP
        Else
            GetSecondDir = RAND2(DIR_DOWN, DIR_UP)
        End If
    End Select

End Function

Function GetDirPriority(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Byte()

    Dim Data(3) As Byte
    Data(0) = GetFirstDir(x1, y1, x2, y2)
    Data(1) = GetSecondDir(Data(0), x1, y1, x2, y2)
    Data(2) = GetOppositeDir(Data(1))
    Data(3) = GetOppositeDir(Data(0))
    GetDirPriority = Data
    
End Function

Function GetFirstDir(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Byte
    Dim DifX As Long, DifY As Long
    DifX = x1 - x2
    DifY = y1 - y2
    
    Dim NextX As Long, NextY As Long
    If Abs(DifX) = Abs(DifY) Then
        If Rnd * 2 < 1 Then
            NextX = x1 + (-1 * DifX \ Abs(DifX))
            NextY = y1
        Else
            NextY = y1 + (-1 * DifY \ Abs(DifY))
            NextX = x1
        End If
    Else
        If Abs(DifX) > Abs(DifY) Then
            ' must move x
            NextX = x1 + (-1 * DifX \ Abs(DifX))
            NextY = y1
        Else
            ' must move y
            NextY = y1 + (-1 * DifY \ Abs(DifY))
            NextX = x1
        End If
    End If
    
    GetFirstDir = GetDirByCollindantPos(x1, y1, NextX, NextY)
End Function
Function GetThirdDir(ByVal SecondDir As Byte) As Byte
    GetThirdDir = GetOppositeDir(SecondDir)
End Function

Function GetFourthDir(ByVal FirstDir As Byte) As Byte
    GetFourthDir = GetOppositeDir(FirstDir)
End Function

Sub GetLinearPath(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByRef NextX As Long, ByRef NextY As Long)
    If x1 = x2 And y1 = y2 Then
        NextX = x1
        NextY = y1
    Else
        Dim DifX As Long, DifY As Long
        DifX = x1 - x2
        DifY = y1 - y2
        
        If Abs(DifX) = Abs(DifY) Then
                If Rnd * 2 < 1 Then
                    NextX = x1 + (-1 * DifX \ Abs(DifX))
                    NextY = y1
                Else
                    NextY = y1 + (-1 * DifY \ Abs(DifY))
                    NextX = x1
                End If
        Else
            If Abs(DifX) > Abs(DifY) Then
                ' must move x
                NextX = x1 + (-1 * DifX \ Abs(DifX))
                NextY = y1
            Else
                ' must move y
                NextY = y1 + (-1 * DifY \ Abs(DifY))
                NextX = x1
            End If
        End If
    End If
End Sub

Function IsFreeObstacle(ByVal mapnum As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal dir As Byte)
    If IsTileAnObstacle(mapnum, x1, y1) Then
        IsFreeObstacle = False
    Else
        If x1 = x2 And y1 = y2 Then
            IsFreeObstacle = True
        Else
            Dim DifX As Long, DifY As Long
            DifX = x1 - x2
            DifY = y1 - y2
            
            If Abs(DifX) = Abs(DifY) Then
                Dim side1 As Boolean, side2 As Boolean
                Dim NextX As Long, NextY As Long
                Dim ForbiddenX As Long, ForbiddenY As Long
                ForbiddenX = x1
                ForbiddenY = y1
                If GetNextPositionByRef(dir, mapnum, ForbiddenX, ForbiddenY) Then
                    IsFreeObstacle = False
                Else
                    
                
                    NextX = x1 + (-1 * DifX \ Abs(DifX))
                    NextY = y1 + (-1 * DifY \ Abs(DifY))
                    
                    If ForbiddenX <> NextX And ForbiddenY <> y1 Then
                        side1 = IsTileAnObstacle(mapnum, NextX, y1)
                    End If
                    If ForbiddenY <> NextY And ForbiddenX <> x1 Then
                        side2 = IsTileAnObstacle(mapnum, x1, NextY)
                    End If
                    
                    If side1 And side2 Then 'can't view
                        IsFreeObstacle = False
                    ElseIf side1 And Not side2 Then
                        IsFreeObstacle = IsFreeObstacle(mapnum, NextX, y1, x2, y2, dir)
                    ElseIf Not side1 And side2 Then
                        IsFreeObstacle = IsFreeObstacle(mapnum, x1, NextY, x2, y2, dir)
                    Else
                        IsFreeObstacle = IsFreeObstacle(mapnum, NextX, NextY, x2, y2, dir)
                    End If
                End If
            Else
                If Abs(DifX) > Abs(DifY) Then
                    ' must move x
                    x1 = x1 + (-1 * DifX \ Abs(DifX))
                Else
                    ' must move y
                    y1 = y1 + (-1 * DifY \ Abs(DifY))
                End If
                
                IsFreeObstacle = IsFreeObstacle(mapnum, x1, y1, x2, y2, dir)
            End If
        End If
    End If
End Function

Function IsInDirection(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal dir As Byte) As Boolean
    IsInDirection = IsAngleInDir(dir, GetAngleByPos(x1, y1, x2, y2))
End Function

Public Function IsXonYRangeVision(ByVal mapnum As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal range As Long, ByVal dir As Byte) As Boolean
    IsXonYRangeVision = (x1 = x2 And y1 = y2)
    
    If Not IsXonYRangeVision Then
        If IsInDirection(x1, y1, x2, y2, dir) Then
            If IsinRange(range, x1, y1, x2, y2) Then
                If IsFreeObstacle(mapnum, x1, y1, x2, y2, dir) Then
                    IsXonYRangeVision = True
                End If
            End If
        End If
    End If


End Function


Function CanNPCViewThoughTile(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long) As Boolean

If OutOfBoundries(X, Y, mapnum) Then Exit Function

Dim TileType As Byte
TileType = map(mapnum).Tile(X, Y).Type

If TileType <> TILE_TYPE_BLOCKED Then
    CanNPCViewThoughTile = True
End If

End Function

Function GetOptimusTile(ByVal Angle As Single, ByVal x0 As Long, ByVal y0 As Long, ByRef X As Long, ByRef Y As Long) As Boolean

Dim CenterX As Long
Dim CenterY As Long

CenterX = X
CenterY = Y


Dim i As Integer
Dim j As Integer
Dim OptimousAngle As Single
OptimousAngle = 181
For i = CenterX - 1 To CenterX + 1
    For j = CenterY - 1 To CenterY + 1
        If i <> CenterX Or j <> CenterY Then
            Dim AngleDirection As Single
            AngleDirection = GetAngleByCenterDistance(i - CenterX, j - CenterY)
            If GetAngleDifference(AngleDirection, Angle) < 45 Then
                Dim AngleDifference As Single
                Dim ActualAngle As Single
                ActualAngle = GetAngleByPos(x0, y0, i, j)
                AngleDifference = GetAngleDifference(ActualAngle, Angle)
                If AngleDifference < OptimousAngle Then
                    OptimousAngle = AngleDifference
                    X = i
                    Y = j
                End If
            End If
        End If
    Next
Next

If OptimousAngle = 181 Then
    GetOptimusTile = False
Else
    GetOptimusTile = True
End If

End Function

Function GetAngleByCenterDistance(ByVal distx As Integer, ByVal disty As Integer) As Single
'-1 <= distx,disty <= 1
If distx = -1 And disty = -1 Then
    GetAngleByCenterDistance = 135
ElseIf distx = -1 And disty = 0 Then
    GetAngleByCenterDistance = 180
ElseIf distx = -1 And disty = 1 Then
    GetAngleByCenterDistance = 225
ElseIf distx = 0 And disty = -1 Then
    GetAngleByCenterDistance = 90
ElseIf distx = 0 And disty = 0 Then
    GetAngleByCenterDistance = 360
ElseIf distx = 0 And disty = 1 Then
    GetAngleByCenterDistance = 270
ElseIf distx = 1 And disty = -1 Then
    GetAngleByCenterDistance = 45
ElseIf distx = 1 And disty = 0 Then
    GetAngleByCenterDistance = 0
ElseIf distx = 1 And disty = 1 Then
    GetAngleByCenterDistance = 315
Else
    GetAngleByCenterDistance = 360
End If

End Function



Function GetAngleDifference(ByVal angle1 As Single, ByVal angle2 As Single)

If Abs(angle1 - angle2) <= 180 Then
    GetAngleDifference = Abs(angle1 - angle2)
Else
    GetAngleDifference = 360 - Abs(angle1 - angle2)
End If


End Function


Sub CheckNPCVision(ByVal mapnum As Long, ByVal mapnpcnum As Long)
    Dim ActionNum As Long
    ActionNum = GetMapNPCAction(mapnum, mapnpcnum)
    If ActionNum > 0 Then
        If Actions(ActionNum).Moment = InFrontRange Then
            Dim a As Variant
            For Each a In GetMapPlayerCollection(mapnum)
                If IsXonYRangeVision(mapnum, MapNpc(mapnum).NPC(mapnpcnum).X, MapNpc(mapnum).NPC(mapnpcnum).Y, GetPlayerX(a), GetPlayerY(a), NPC(MapNpc(mapnum).NPC(mapnpcnum).Num).range, MapNpc(mapnum).NPC(mapnpcnum).dir) Then
                    Call ParseAction(a, ActionNum)
                End If
            Next
        End If
    End If
End Sub

Sub IsPlayerOnNPCVision(ByVal index As Long)
    Dim i As Long, mapnum As Long
    mapnum = GetPlayerMap(index)
    For i = 1 To GetMapNpcHighIndex(mapnum)
        ActionNum = GetMapNPCAction(mapnum, i)
        If ActionNum > 0 Then
            If Actions(ActionNum).Moment = InFrontRange Then
                If IsXonYRangeVision(mapnum, MapNpc(mapnum).NPC(i).X, MapNpc(mapnum).NPC(i).Y, GetPlayerX(index), GetPlayerY(index), NPC(MapNpc(mapnum).NPC(i).Num).range, MapNpc(mapnum).NPC(i).dir) Then
                    Call ParseAction(index, ActionNum)
                End If
            ElseIf Actions(ActionNum).Moment = TileMatch Then
                If GetPlayerX(index) = MapNpc(mapnum).NPC(i).X And GetPlayerY(index) = MapNpc(mapnum).NPC(i).Y Then
                    Call ParseAction(index, ActionNum)
                End If
            End If
        End If
    Next
End Sub


