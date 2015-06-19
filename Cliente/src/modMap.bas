Attribute VB_Name = "modMap"



Public Function GetMapData(ByRef MapT As MapRec) As Byte()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    With MapT
        Buffer.WriteConstString "v.2"
        Buffer.WriteConstString .Name
        If Len(Trim$(Replace(.TranslatedName, vbNullChar, ""))) = 0 Then .TranslatedName = GetTranslation(.Name)
        Buffer.WriteConstString .TranslatedName
        Buffer.WriteConstString .Music
        
        Buffer.WriteLong .Revision
        Buffer.WriteByte .moral
        
        Buffer.WriteLong .Up
        Buffer.WriteLong .Down
        Buffer.WriteLong .Left
        Buffer.WriteLong .Right
        
        Buffer.WriteLong .BootMap
        Buffer.WriteByte .BootX
        Buffer.WriteByte .BootY
        
        Buffer.WriteByte .MaxX
        Buffer.WriteByte .MaxY
        Dim X As Byte, y As Byte
        For X = 0 To .MaxX
            For y = 0 To .MaxY
                Dim j As Byte
                For j = 1 To Layer_Count - 1
                    Buffer.WriteLong .Tile(X, y).layer(j).X
                    Buffer.WriteLong .Tile(X, y).layer(j).y
                    Buffer.WriteLong .Tile(X, y).layer(j).Tileset
                Next
                Buffer.WriteByte .Tile(X, y).Type
                Buffer.WriteLong .Tile(X, y).Data1
                Buffer.WriteLong .Tile(X, y).Data2
                Buffer.WriteLong .Tile(X, y).Data3
                Buffer.WriteByte .Tile(X, y).DirBlock
            Next
        Next

        For X = 1 To MAX_MAP_NPCS
            Buffer.WriteLong .NPC(X)
            Buffer.WriteByte .NPCSProperties(X).movement
            Buffer.WriteByte .NPCSProperties(X).Action
        Next

        Buffer.WriteLong .Weather
        
        For X = 1 To Max_States - 1
            Buffer.WriteByte .AllowedStates(X)
        Next
    End With
    
    GetMapData = Buffer.ToArray
    'Debug.Print Buffer.ToString
    Set Buffer = Nothing
End Function

Public Sub SetMapData(ByRef map As MapRec, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim newVer As Boolean
    Buffer.WriteBytes Data
    With map
        If Buffer.ReadConstString(3, False) = "v.2" Then newVer = True: Buffer.MoveReadHead 3
        .Name = Buffer.ReadConstString(NAME_LENGTH)
        'small patch.. not a safe fix!
         If newVer = True Then .TranslatedName = Buffer.ReadConstString(NAME_LENGTH)
         If newVer = False Then .TranslatedName = GetTranslation(.Name)
        .Music = Buffer.ReadConstString(NAME_LENGTH)
        .Revision = Buffer.ReadLong
        .moral = Buffer.ReadByte
        .Up = Buffer.ReadLong
        .Down = Buffer.ReadLong
        .Left = Buffer.ReadLong
        .Right = Buffer.ReadLong
        .BootMap = Buffer.ReadLong
        .BootX = Buffer.ReadByte
        .BootY = Buffer.ReadByte
        .MaxX = Buffer.ReadByte
        .MaxY = Buffer.ReadByte
        ReDim .Tile(0 To .MaxX, 0 To .MaxY)

        For X = 0 To .MaxX
            For y = 0 To .MaxY
                Dim j As Byte
                For j = 1 To Layer_Count - 1
                    .Tile(X, y).layer(j).X = Buffer.ReadLong
                    .Tile(X, y).layer(j).y = Buffer.ReadLong
                    .Tile(X, y).layer(j).Tileset = Buffer.ReadLong
                Next
                .Tile(X, y).Type = Buffer.ReadByte
                .Tile(X, y).Data1 = Buffer.ReadLong
                .Tile(X, y).Data2 = Buffer.ReadLong
                .Tile(X, y).Data3 = Buffer.ReadLong
                .Tile(X, y).DirBlock = Buffer.ReadByte
            Next
        Next

        For X = 1 To MAX_MAP_NPCS
            .NPC(X) = Buffer.ReadLong
            .NPCSProperties(X).movement = Buffer.ReadByte
            .NPCSProperties(X).Action = Buffer.ReadByte
        Next

        .Weather = Buffer.ReadLong
        
        For X = 1 To Max_States - 1
            .AllowedStates(X) = Buffer.ReadByte
        Next
    
    End With
    Set Buffer = Nothing
End Sub

