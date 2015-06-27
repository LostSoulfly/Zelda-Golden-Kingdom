Attribute VB_Name = "modMap"
Option Explicit



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
        Dim X As Byte, Y As Byte
        For X = 0 To .MaxX
            For Y = 0 To .MaxY
                Dim j As Byte
                For j = 1 To Layer_Count - 1
                    Buffer.WriteLong .Tile(X, Y).layer(j).X
                    Buffer.WriteLong .Tile(X, Y).layer(j).Y
                    Buffer.WriteLong .Tile(X, Y).layer(j).Tileset
                Next
                Buffer.WriteByte .Tile(X, Y).Type
                Buffer.WriteLong .Tile(X, Y).Data1
                Buffer.WriteLong .Tile(X, Y).Data2
                Buffer.WriteLong .Tile(X, Y).Data3
                Buffer.WriteByte .Tile(X, Y).DirBlock
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
    Dim X As Long, Y As Long
    Buffer.WriteBytes Data
    With map
        If Buffer.ReadConstString(3, False) = "v.2" Then newVer = True: Buffer.MoveReadHead 3
        .Name = Buffer.ReadConstString(NAME_LENGTH)
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
            For Y = 0 To .MaxY
                Dim j As Byte
                For j = 1 To Layer_Count - 1
                    .Tile(X, Y).layer(j).X = Buffer.ReadLong
                    .Tile(X, Y).layer(j).Y = Buffer.ReadLong
                    .Tile(X, Y).layer(j).Tileset = Buffer.ReadLong
                Next
                .Tile(X, Y).Type = Buffer.ReadByte
                .Tile(X, Y).Data1 = Buffer.ReadLong
                .Tile(X, Y).Data2 = Buffer.ReadLong
                .Tile(X, Y).Data3 = Buffer.ReadLong
                .Tile(X, Y).DirBlock = Buffer.ReadByte
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

