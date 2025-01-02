Attribute VB_Name = "modMap"
Option Explicit



Public Function GetMapData(ByRef MapT As MapRec) As Byte()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    With MapT
        buffer.WriteConstString "v.2"
        buffer.WriteConstString .Name
        If Len(Trim$(Replace(.Name, vbNullChar, ""))) = 0 Then .Name = .Name
        buffer.WriteConstString .Name
        buffer.WriteConstString .Music
        
        buffer.WriteLong .Revision
        buffer.WriteByte .moral
        
        buffer.WriteLong .Up
        buffer.WriteLong .Down
        buffer.WriteLong .Left
        buffer.WriteLong .Right
        
        buffer.WriteLong .BootMap
        buffer.WriteByte .BootX
        buffer.WriteByte .BootY
        
        buffer.WriteByte .MaxX
        buffer.WriteByte .MaxY
        Dim X As Byte, Y As Byte
        For X = 0 To .MaxX
            For Y = 0 To .MaxY
                Dim j As Byte
                For j = 1 To Layer_Count - 1
                    buffer.WriteLong .Tile(X, Y).layer(j).X
                    buffer.WriteLong .Tile(X, Y).layer(j).Y
                    buffer.WriteLong .Tile(X, Y).layer(j).Tileset
                Next
                buffer.WriteByte .Tile(X, Y).Type
                buffer.WriteLong .Tile(X, Y).Data1
                buffer.WriteLong .Tile(X, Y).Data2
                buffer.WriteLong .Tile(X, Y).Data3
                buffer.WriteByte .Tile(X, Y).DirBlock
            Next
        Next

        For X = 1 To MAX_MAP_NPCS
            buffer.WriteLong .NPC(X)
            buffer.WriteByte .NPCSProperties(X).movement
            buffer.WriteByte .NPCSProperties(X).Action
        Next

        buffer.WriteLong .Weather
        
        For X = 1 To Max_States - 1
            buffer.WriteByte .AllowedStates(X)
        Next
    End With
    
    GetMapData = buffer.ToArray
    'Debug.Print Buffer.ToString
    Set buffer = Nothing
End Function

Public Sub SetMapData(ByRef map As MapRec, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim newVer As Boolean
    Dim X As Long, Y As Long
    buffer.WriteBytes Data
    With map
        If buffer.ReadConstString(3, False) = "v.2" Then newVer = True: buffer.MoveReadHead 3
        .Name = buffer.ReadConstString(NAME_LENGTH)
         If newVer = True Then .Name = buffer.ReadConstString(NAME_LENGTH)
         If newVer = False Then .Name = .Name
        .Music = buffer.ReadConstString(NAME_LENGTH)
        .Revision = buffer.ReadLong
        .moral = buffer.ReadByte
        .Up = buffer.ReadLong
        .Down = buffer.ReadLong
        .Left = buffer.ReadLong
        .Right = buffer.ReadLong
        .BootMap = buffer.ReadLong
        .BootX = buffer.ReadByte
        .BootY = buffer.ReadByte
        .MaxX = buffer.ReadByte
        .MaxY = buffer.ReadByte
        ReDim .Tile(0 To .MaxX, 0 To .MaxY)

        For X = 0 To .MaxX
            For Y = 0 To .MaxY
                Dim j As Byte
                For j = 1 To Layer_Count - 1
                    .Tile(X, Y).layer(j).X = buffer.ReadLong
                    .Tile(X, Y).layer(j).Y = buffer.ReadLong
                    .Tile(X, Y).layer(j).Tileset = buffer.ReadLong
                Next
                .Tile(X, Y).Type = buffer.ReadByte
                .Tile(X, Y).Data1 = buffer.ReadLong
                .Tile(X, Y).Data2 = buffer.ReadLong
                .Tile(X, Y).Data3 = buffer.ReadLong
                .Tile(X, Y).DirBlock = buffer.ReadByte
            Next
        Next

        For X = 1 To MAX_MAP_NPCS
            .NPC(X) = buffer.ReadLong
            .NPCSProperties(X).movement = buffer.ReadByte
            .NPCSProperties(X).Action = buffer.ReadByte
        Next

        .Weather = buffer.ReadLong
        
        For X = 1 To Max_States - 1
            .AllowedStates(X) = buffer.ReadByte
        Next
    
    End With
    Set buffer = Nothing
End Sub

