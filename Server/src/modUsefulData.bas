Attribute VB_Name = "modUsefulData"
Function SendUsefulDataToPlayer(ByVal index As Long) As Boolean
    If GetPlayerAccess_Mode(index) = 0 Then
        SendUsefulDataToPlayer = True
    End If
End Function

Function GetSpellUsefulData(ByVal spellnum As Long) As Byte()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    With Spell(spellnum)
    buffer.WriteString Trim(.Name)
    buffer.WriteString Trim(.Desc)
    buffer.WriteString Trim(.Sound)
    
    buffer.WriteLong .MPCost
    buffer.WriteLong .CDTime
    buffer.WriteLong .Icon
    buffer.WriteLong .CastTime
    End With
    GetSpellUsefulData = buffer.ToArray
End Function

Function GetNPCUsefulData(ByVal npcnum As Long) As Byte()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    With NPC(npcnum)
    buffer.WriteString Trim(.Name)
    buffer.WriteString Trim(.Sound)
    buffer.WriteLong .Sprite
    
    buffer.WriteByte .Behaviour
    
    Dim i As Byte
    For i = 1 To Stats.Stat_Count - 1
        buffer.WriteByte .stat(i)
    Next
    buffer.WriteLong .HP
    buffer.WriteLong .questnum
    buffer.WriteLong .Speed
    buffer.WriteLong .level
    End With
    GetNPCUsefulData = buffer.ToArray
End Function


Function GetItemUsefulData(ByVal ItemNum As Long) As Byte()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    With item(ItemNum)
    buffer.WriteString Trim(.Name)
    buffer.WriteString Trim(.Desc)
    buffer.WriteString Trim(.Sound)
    
    buffer.WriteLong .Pic
    buffer.WriteByte .Type
    buffer.WriteLong .Price
    buffer.WriteByte .Rarity
    buffer.WriteLong .Speed
    buffer.WriteLong .Paperdoll
    buffer.WriteLong .ProjecTile.Pic
    buffer.WriteLong .weight

    End With
    GetItemUsefulData = buffer.ToArray
    Set buffer = Nothing
End Function

Function GetResourceUsefulData(ByVal ResourceNum As Long) As Byte()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    With Resource(ResourceNum)
    buffer.WriteLong .ResourceImage
    buffer.WriteLong .ExhaustedImage
    buffer.WriteByte .WalkableNormal
    buffer.WriteByte .WalkableExhausted
    End With
    GetResourceUsefulData = buffer.ToArray
    Set buffer = Nothing
End Function


Function GetAnimationUseFulData(ByVal AnimNum As Long) As Byte()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    With Animation(AnimNum)
    buffer.WriteString .Sound
    Dim i As Byte
    For i = 0 To 1
       buffer.WriteLong .Sprite(i)
       buffer.WriteLong .Frames(i)
       buffer.WriteLong .LoopCount(i)
       buffer.WriteLong .LoopTime(i)
    Next
    End With
    GetAnimationUseFulData = buffer.ToArray
    Set buffer = Nothing
End Function

Public Function MapToServerMap(ByRef map As MapRec) As ServerMapRec
    With MapToServerMap
        .Name = map.Name
        .Revision = map.Revision
        .moral = map.moral
        
        .Up = map.Up
        .Down = map.Down
        .left = map.left
        .right = map.right
        
        .BootMap = map.BootMap
        .BootX = map.BootX
        .BootY = map.BootY
        
        .MaxX = map.MaxX
        .MaxY = map.MaxY
        
        ReDim .Tile(0 To .MaxX, 0 To .MaxY)
        Dim X As Long, Y As Long, j As Long
        For X = 0 To .MaxX
            For Y = 0 To .MaxY
                .Tile(X, Y).Data1 = map.Tile(X, Y).Data1
                .Tile(X, Y).Data2 = map.Tile(X, Y).Data2
                .Tile(X, Y).Data3 = map.Tile(X, Y).Data3
                .Tile(X, Y).DirBlock = map.Tile(X, Y).DirBlock
                .Tile(X, Y).Type = map.Tile(X, Y).Type
            Next
        Next
        
        
        For X = 1 To MAX_MAP_NPCS
            .NPC(X) = map.NPC(X)
            .NPCSProperties(X) = map.NPCSProperties(X)
        Next
        
        .Weather = map.Weather
        
        For X = 1 To Max_States - 1
           .AllowedStates(X) = map.AllowedStates(X)
        Next
    
    
    End With
End Function
