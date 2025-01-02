Attribute VB_Name = "modPlayerMove"
Option Explicit

Private RunningSprites() As clsPair
Private N As Long

Private RunSpritesPointer() As Long
Private WalkSpritesPointer() As Long

Sub PlayerMove(ByVal index As Long, ByVal dir As Byte, ByVal movement As Byte)
    Dim X As Long, Y As Long
    X = Player(index).X
    Y = Player(index).Y
    If GetNextPositionByRef(dir, X, Y) Then
        Exit Sub
    End If
    SetPlayerX index, X
    SetPlayerY index, Y
    Call SetPlayerDir(index, dir)
    
    Select Case GetPlayerDir(index)
        Case DIR_UP
            Player(index).YOffset = PIC_Y
        Case DIR_DOWN
            Player(index).YOffset = PIC_Y * -1
        Case DIR_LEFT
            Player(index).XOffset = PIC_X
        Case DIR_RIGHT
            Player(index).XOffset = PIC_X * -1
    End Select
    
    Player(index).Moving = movement
End Sub

Function GetRunningSprite(ByVal sprite As Long) As Long
    If sprite > UBound(WalkSpritesPointer) Then Exit Function
    If WalkSpritesPointer(sprite) = 0 Then Exit Function
    GetRunningSprite = RunningSprites(WalkSpritesPointer(sprite)).GetSecond
End Function

Function GetWalkingSprite(ByVal index As Long) As Long
    GetWalkingSprite = Player(index).PreviousSprite
    Exit Function
    'If sprite > UBound(RunSpritesPointer) Then Exit Function
    'If RunSpritesPointer(sprite) = 0 Then Exit Function
    'GetWalkingSprite = RunningSprites(RunSpritesPointer(sprite)).GetFirst
End Function


Public Sub HandleRunningSprites(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Dim i As Long
    N = Buffer.ReadLong
    ReDim RunningSprites(1 To N)
    For i = 1 To N
        Set RunningSprites(i) = New clsPair
        With RunningSprites(i)
        .SetFirst Buffer.ReadLong
        .SetSecond Buffer.ReadLong
        End With
    Next
      
    CreatePointers
    Set Buffer = Nothing
End Sub

Private Sub CreatePointers()
    Dim WalkMax As Long, RunMax As Long
    FindMaxes WalkMax, RunMax
    
    If WalkMax = 0 Or RunMax = 0 Then Exit Sub
    
    ReDim WalkSpritesPointer(1 To WalkMax)
    ReDim RunSpritesPointer(1 To RunMax)
    
    Dim i As Long
    For i = 1 To N
        With RunningSprites(i)
        WalkSpritesPointer(.GetFirst) = i
        RunSpritesPointer(.GetSecond) = i
        End With
    Next
End Sub

Private Sub FindMaxes(ByRef WalkMax As Long, ByRef RunMax As Long)
    Dim i As Long
    WalkMax = 0
    RunMax = 0
    Dim WalkIndex As Long, RunIndex As Long
    For i = 1 To N
        With RunningSprites(i)
        If i = 1 Then
            WalkIndex = i
            RunIndex = i
            WalkMax = .GetFirst
            RunMax = .GetSecond
        Else
            If .GetFirst > WalkMax Then
                WalkIndex = i
                WalkMax = .GetFirst
            End If
            If .GetSecond > RunMax Then
                RunIndex = i
                RunMax = .GetSecond
            End If
        End If
        End With
    Next
End Sub


