Attribute VB_Name = "modUsefulData"
Function SendUsefulDataToPlayer(ByVal index As Long) As Boolean
    If GetPlayerAccess_Mode(index) = 0 Then
        SendUsefulDataToPlayer = True
    End If
End Function

Function GetSpellUsefulData(ByVal spellnum As Long) As Byte()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    With Spell(spellnum)
    Buffer.WriteString Trim$(.Name)
    Buffer.WriteString Trim$(.TranslatedName)
    Buffer.WriteString Trim$(.Desc)
    Buffer.WriteString Trim$(.Sound)
    
    Buffer.WriteLong .MPCost
    Buffer.WriteLong .CDTime
    Buffer.WriteLong .Icon
    Buffer.WriteLong .CastTime
    End With
    GetSpellUsefulData = Buffer.ToArray
End Function

Function GetNPCUsefulData(ByVal npcnum As Long) As Byte()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    With NPC(npcnum)
    Buffer.WriteString Trim$(.Name)
    Buffer.WriteString Trim$(.TranslatedName)
    Buffer.WriteString Trim$(.Sound)
    Buffer.WriteLong .Sprite
    
    Buffer.WriteByte .Behaviour
    
    Dim i As Byte
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteByte .stat(i)
    Next
    Buffer.WriteLong .HP
    Buffer.WriteLong .questnum
    Buffer.WriteLong .Speed
    Buffer.WriteLong .level
    End With
    GetNPCUsefulData = Buffer.ToArray
End Function


Function GetItemUsefulData(ByVal ItemNum As Long) As Byte()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    With item(ItemNum)
    Buffer.WriteString Trim$(.Name)
    Buffer.WriteString Trim$(.TranslatedName)
    Buffer.WriteString Trim$(.Desc)
    Buffer.WriteString Trim$(.Sound)
    
    Buffer.WriteLong .Pic
    Buffer.WriteByte .Type
    Buffer.WriteLong .Price
    Buffer.WriteByte .Rarity
    Buffer.WriteLong .Speed
    Buffer.WriteLong .Paperdoll
    Buffer.WriteLong .ProjecTile.Pic
    Buffer.WriteLong .weight

    End With
    GetItemUsefulData = Buffer.ToArray
    Set Buffer = Nothing
End Function

Function GetResourceUsefulData(ByVal ResourceNum As Long) As Byte()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    With Resource(ResourceNum)
    Buffer.WriteLong .ResourceImage
    Buffer.WriteLong .ExhaustedImage
    Buffer.WriteByte .WalkableNormal
    Buffer.WriteByte .WalkableExhausted
    End With
    GetResourceUsefulData = Buffer.ToArray
    Set Buffer = Nothing
End Function


Function GetAnimationUseFulData(ByVal AnimNum As Long) As Byte()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    With Animation(AnimNum)
    Buffer.WriteString .Sound
    Dim i As Byte
    For i = 0 To 1
       Buffer.WriteLong .Sprite(i)
       Buffer.WriteLong .Frames(i)
       Buffer.WriteLong .LoopCount(i)
       Buffer.WriteLong .LoopTime(i)
    Next
    End With
    GetAnimationUseFulData = Buffer.ToArray
    Set Buffer = Nothing
End Function

Public Function MapToServerMap(ByRef map As MapRec) As ServerMapRec
    With MapToServerMap
        .Name = map.Name
        .TranslatedName = GetTranslation(.Name)
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
