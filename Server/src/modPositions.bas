Attribute VB_Name = "modPositions"
Option Explicit

Private Visited() As Integer
Public Function Calculate2PointsVector(ByVal x1 As Long, ByVal x2 As Long, ByVal y1 As Long, ByVal y2 As Long) As clsDirVector
    Dim p1 As clsPoint, p2 As clsPoint
    Set p1 = New clsPoint
    Set p2 = New clsPoint
    
    p1.SetX x1
    p1.SetY y1
    
    p2.SetX y1
    p2.SetY y2
    
    Set Calculate2PointsVector = New clsVector
    Calculate2PointsVector.SetVector p1, p2
End Function



Sub SearchPath(ByVal index As Long, ByVal mapnum As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal y1 As Long, ByVal y2 As Long)
    Dim p1 As clsPosition, p2 As clsPosition
    Dim Point1 As clsPoint, Point2 As clsPoint
    Set Point1 = New clsPoint
    Point1.SetX x1
    Point1.SetY y1
    
    Set p1 = New clsPosition
    Set p2 = New clsPosition
    
    p1.SetPoint Point1
    p1.SetMap mapnum
    
    Set Point2 = New clsPoint
    Point2.SetX x2
    Point2.SetY y2
    
    p2.SetPoint Point2
    p2.SetMap mapnum
    
    Dim s As clsStack
    Set s = New clsStack
    Dim calls As Long
    ReDim Visited(0 To map(mapnum).MaxX, 0 To map(mapnum).MaxY)
    ' 1: alredy visited
    calls = 0
    'Call Search_Path_Rec(p1, p2, s, calls)
    Dim d As clsDirection
    Dim a As Long
    a = GetTickCount
    Call Search_Path_Rec_2(p1, p2, d, s, calls, 0)
    PlayerMsg index, GetTickCount - a, 1, , False
    
    'Dim a As Long
    'a = GetTickCount
    'Call Invert(s)
    
    'Dim V As clsVector
   ' Set V = s.ToVector
    'Call DeleteTrash(V)
    
    'Dim q As clsQueue
    'Set q = New clsQueue
    'Dim i As Long
    'For i = 0 To V.GetSize - 1
       ' q.Push V.GetElem(i)
   ' Next
    
    'Set s = V.ToStack
    GenerateSpaces GetPlayerMap(index)
        
    Set TempPlayer(index).MovementsStack = s
End Sub

Function DirToPos(ByRef DirVector As clsVector, Optional ByRef StartPoint As clsPoint) As clsVector
    If StartPoint Is Nothing Then
        Set StartPoint = New clsPoint
        StartPoint.SetX 0
        StartPoint.SetY 0
    End If
    
    Set DirToPos = New clsVector
    DirToPos.SetSize DirVector.GetSize
    Call DirToPos.SetElem(StartPoint, 0)
    
    Dim i As Long
    For i = 0 To DirVector.GetSize - 1
    
        Call StartPoint.GetNextPoint(DirVector.GetElem(i))
        Call DirToPos.SetElem(StartPoint, i + 1)
    Next
End Function

Sub DeleteTrash(ByRef v As clsVector)
    ' Vector of points
    Dim i As Long
    i = 0
    While i < v.GetSize
        Dim j As Long
        j = i + 2
        Dim AtLeastOne As Boolean
        While j < v.GetSize
            Dim d As clsDirection
            Set d = Colindant(v, i, j)
            If d.GetDir <> 4 Then
                Call v.SetElem(d, i)
                Dim k As Long
                For k = i + 1 To j
                    v.DeleteElem (k)
                Next
                
                j = i + 2
                
                v.Unificate
            Else
                j = j + 1
            End If
        Wend
        i = i + 1
    Wend
    
    
        
        
End Sub

Function Colindant(ByRef DirVec As clsVector, ByVal StartI As Long, ByVal EndI As Long) As clsDirection
    'returns direction: 4: no colindant
    Dim StartPoint As clsPoint
    Set StartPoint = New clsPoint
    StartPoint.SetX 0
    StartPoint.SetY 0
    Set Colindant = New clsDirection
    Colindant.SetDir 4
    Dim i As Long
    For i = StartI To EndI
        StartPoint.GetNextPoint DirVec.GetElem(i)
    Next
    
    If Abs(StartPoint.GetX) + Abs(StartPoint.GetY) <= 1 Then
        Dim X As Long, Y As Long
        X = StartPoint.GetX
        Y = StartPoint.GetY
        
        If X = 0 And Y = 1 Then
            Colindant.SetDir DIR_DOWN
        ElseIf X = 1 And Y = 0 Then
            Colindant.SetDir DIR_RIGHT
        ElseIf X = -1 And Y = 0 Then
            Colindant.SetDir DIR_LEFT
        ElseIf X = 0 And Y = -1 Then
            Colindant.SetDir DIR_UP
        End If
    End If
End Function

Function Search_Path_Rec_2(ByRef p1 As clsPosition, ByRef p2 As clsPosition, ByRef ComeDir As clsDirection, ByRef Dir_Stack As clsStack, ByRef high As Long, ByVal calls As Long) As Integer
    If p1.Equals(p2) Then
        high = 1
        Search_Path_Rec_2 = 1
        Set Dir_Stack = New clsStack
    ElseIf p1.IsBlocked Or Visited(p1.GetPoint.GetX, p1.GetPoint.GetY) = -1 Then
        Search_Path_Rec_2 = -1
        Visited(p1.GetPoint.GetX, p1.GetPoint.GetY) = -1
    Else
        Dim Sol(0 To 3) As clsStack
        Dim height(0 To 3) As Long
        Dim state(0 To 3) As Integer
        Visited(p1.GetPoint.GetX, p1.GetPoint.GetY) = -1
        Dim v As clsDirVector
        Set v = New clsDirVector
        Call v.SetVector(p1.GetPoint, p2.GetPoint)
        Dim PriorityStack As clsStack
        Set PriorityStack = v.GetVectorAngle.GetPriorityStackDir(ComeDir)
        
        Dim i As Byte
        While Not PriorityStack.IsEmpty
            
                Dim NextPos As clsPosition
                i = PriorityStack.Front.GetDir
                If p1.GetNextPosition(PriorityStack.Front, NextPos) Then
                    state(i) = Search_Path_Rec_2(NextPos, p2, PriorityStack.Front, Sol(i), height(i), calls + 1)
                End If
                
                PriorityStack.Pop
        Wend
        Dim best As Byte
        best = 4
        Search_Path_Rec_2 = -1
        For i = 0 To 3
            If state(i) = 1 Then
                Search_Path_Rec_2 = 1
                If best = 4 Then
                    high = height(i)
                    best = i
                Else
                    If height(i) < high Then
                        high = height(i)
                        best = i
                    End If
                End If
            End If
        Next
        If best <> 4 Then
            Visited(p1.GetPoint.GetX, p1.GetPoint.GetY) = 0
            Dim d As clsDirection
            Set d = New clsDirection
            d.SetDir best
            Set Dir_Stack = Sol(best)
            Dir_Stack.Push d
            high = high + 1
        End If
    End If
    
End Function

Function Search_Path_Rec(ByRef p1 As clsPosition, ByRef p2 As clsPosition, ByRef Dir_Stack As clsStack, ByVal calls As Long) As Integer
    'blocked : returns -1; found: returns 1; nothing: returns 0
    If p1.Equals(p2) Then
        Search_Path_Rec = 1
    ElseIf p1.IsBlocked Or Visited(p1.GetPoint.GetX, p1.GetPoint.GetY) = -1 Then
        Search_Path_Rec = -1
        Visited(p1.GetPoint.GetX, p1.GetPoint.GetY) = -1
    Else
        Visited(p1.GetPoint.GetX, p1.GetPoint.GetY) = -1
        Dim v As clsDirVector
        Set v = New clsDirVector
        Call v.SetVector(p1.GetPoint, p2.GetPoint)
        Dim PriorityStack As clsStack
        Set PriorityStack = v.GetVectorAngle.GetPriorityStackDir(Dir_Stack.Front)
        
        While Not PriorityStack.IsEmpty
            
                Dim NextPos As clsPosition
                If p1.GetNextPosition(PriorityStack.Front, NextPos) Then
                    Dir_Stack.Push PriorityStack.Front
                    Dim state As Integer
                    state = Search_Path_Rec(NextPos, p2, Dir_Stack, calls + 1)
                    If state = -1 Then
                        Dir_Stack.Pop
                    ElseIf state = 1 Then
                        Search_Path_Rec = 1
                        PriorityStack.Clear
                    End If
                End If
                PriorityStack.Pop
        Wend
        
        If Search_Path_Rec = 0 Then
            Search_Path_Rec = -1
        End If
    End If
End Function

Function GenerateSpaces(ByVal mapnum As Long) As Integer()
    
    Dim Spaces() As Integer
    ReDim Spaces(0 To map(mapnum).MaxX, 0 To map(mapnum).MaxY)
    
    Dim counter As Integer
    counter = 1
    
    Dim X As Long, Y As Long
    Dim x2 As Long, y2 As Long
    For X = 0 To map(mapnum).MaxX
        For Y = 0 To map(mapnum).MaxY
            For x2 = 0 To map(mapnum).MaxX
                For y2 = 0 To map(mapnum).MaxY
                    If IsTileblocked(mapnum, X, Y) Then
                        Spaces(X, Y) = -2
                    ElseIf IsTileblocked(mapnum, x2, y2) Then
                        Spaces(x2, y2) = -2
                    Else
                        If Spaces(X, Y) > -1 And Spaces(x2, y2) > -1 Then
                            If X = x2 And Y = y2 And Spaces(X, Y) = 0 Then
                                Spaces(X, Y) = counter
                            Else
                                If CanWalkTo(mapnum, X, Y, x2, y2) Then
                                    If Spaces(X, Y) = Spaces(x2, y2) Then
                                        
                                    Else
                                        If Abs(x2 - X) + Abs(y2 - Y) Then
                                            If Spaces(X, Y) <> 0 And Spaces(x2, y2) <> 0 Then
                                                Spaces(X, Y) = -1
                                            End If
                                        Else
                                            
                                        End If
                                    End If
                                    
                                Else
                                    If Spaces(X, Y) = Spaces(x2, y2) Then
                                        Spaces(x2, y2) = Spaces(x2, y2) + 1
                                    ElseIf Spaces(x2, y2) = 0 Then
                                        Spaces(x2, y2) = Spaces(X, Y) + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            Next
        Next
    Next
    
    
    For X = 0 To map(mapnum).MaxX
        For Y = 0 To map(mapnum).MaxY
            If Spaces(X, Y) > 0 Then
                SendAnimation mapnum, Spaces(X, Y), X, Y
            End If
        Next
    Next
End Function

Function CanWalkTo(ByVal mapnum As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    Dim i As Long
    CanWalkTo = True
    If x1 = x2 Then
        For i = y1 To y2
            If IsTileblocked(mapnum, x1, i) Then
                CanWalkTo = False
            End If
        Next
    
    ElseIf y1 = y2 Then
        For i = x1 To x2
            If IsTileblocked(mapnum, i, y1) Then
                CanWalkTo = False
            End If
        Next
    
    Else
        Dim reach As Boolean
        Do
            reach = (x1 = x2) And (y1 = y2)
            If IsTileblocked(mapnum, x1, y1) Then
                CanWalkTo = False
                reach = True
            End If
            If Not reach Then
                If Abs(x1 - x2) > Abs(y1 - y2) Then
                    If x1 > x2 Then
                        x1 = x1 - 1
                    Else
                        x1 = x1 + 1
                    End If
                Else
                    If y1 > y2 Then
                        y1 = y1 - 1
                    Else
                        y1 = y1 + 1
                    End If
                End If
            End If
        Loop Until reach
    
    
    End If
End Function

Function IsTileblocked(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    If GetTileType(mapnum, X, Y) = TILE_TYPE_BLOCKED Then
        IsTileblocked = True
    End If
End Function


Public Sub Invert(ByRef s As clsStack)
    Dim aux As clsStack
    Set aux = New clsStack
    While Not s.IsEmpty
        aux.Push s.Front
        s.Pop
    Wend
    
    Set s = aux
    
End Sub
