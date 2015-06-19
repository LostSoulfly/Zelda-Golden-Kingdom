Attribute VB_Name = "modMap"
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Public FixWarpMap As Long
Public FixWarpMap_Enabled As Boolean

Private Type WaitingNPCSRec
    mapnpcnum As Long
    npcnum As Long
    Timer As Long
End Type
'used in random spawning
Private Type TempMapRec
    Exists As Boolean
    npc_highindex As Long
    Item_highindex As Long
    WaitingForSpawnItems() As WaitingItemsRec
    HasItems As Boolean
    
    WaitingForSpawnNPCS() As WaitingNPCSRec
    WaitingNPCS As Long
    
    mapref As Long
    
    PlayersOnMap As Collection
End Type

Private Type TileReferenceRec
    mapnpcnum As Byte
    ResourceIndex As Long
    LastMapNpcNum As Byte
End Type

Private Type MapReferenceRec
    Tiles() As TileReferenceRec
    MapNum As Long
    NumPlayers As Long
End Type

Public TempMap(1 To MAX_MAPS) As TempMapRec

Public MapReferences() As MapReferenceRec
Private NMaps As Long

Public CurrentMapIndex As Long

Function GetMapPlayerCollection(ByVal MapNum As Long) As Collection
    Dim i As Long
    i = GetMapRef(MapNum)
    If i > 0 Then
        Set GetMapPlayerCollection = TempMap(MapNum).PlayersOnMap
    Else
        Set GetMapPlayerCollection = New Collection
    End If
End Function


Function FindMapPlayerSlot(ByVal MapNum As Long, ByVal index As Long) As Long
    If MapNum = 0 Or index = 0 Then Exit Function
    Dim i As Long
        With TempMap(MapNum)
            For i = 1 To .PlayersOnMap.Count
                If .PlayersOnMap.item(i) = index Then
                    FindMapPlayerSlot = i
                    Exit Function
                End If
            Next
        End With
End Function



Sub AddWaitingNPC(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal Time As Long)
    If MapNum = 0 Or mapnpcnum = 0 Then Exit Sub
    Dim i As Long
    With TempMap(MapNum)
    .WaitingNPCS = .WaitingNPCS + 1
    
    
    Dim AuxNPC As WaitingNPCSRec
    AuxNPC.mapnpcnum = mapnpcnum
    AuxNPC.Timer = GetRealTickCount + 1000 * Time
    
    i = .WaitingNPCS
    ReDim Preserve .WaitingForSpawnNPCS(1 To i)
    
    
    If i > 1 Then
        While .WaitingForSpawnNPCS(i).Timer > AuxNPC.Timer And i > 0
            .WaitingForSpawnNPCS(i) = .WaitingForSpawnNPCS(i - 1)
            i = i - 1
        Wend
    End If
    .WaitingForSpawnNPCS(i) = AuxNPC
    End With
End Sub

Sub CheckWaitingNPCS(ByVal Tick As Long)
    Dim i As Long
    For i = 1 To Map_highindex
        With TempMap(i)
        If .WaitingNPCS > 0 Then
            If .WaitingForSpawnNPCS(1).Timer < Tick Then
                Call SpawnNpc(.WaitingForSpawnNPCS(1).mapnpcnum, i)
                Call PopWaitingNPC(i)
            End If
        End If
        End With
    Next
End Sub

Sub PopWaitingNPC(ByVal MapNum As Long)
    If MapNum = 0 Then Exit Sub

    With TempMap(MapNum)
    If .WaitingNPCS > 1 Then
        Dim i As Long
        For i = 1 To .WaitingNPCS - 1
            .WaitingForSpawnNPCS(i) = .WaitingForSpawnNPCS(i + 1)
        Next
        ReDim Preserve .WaitingForSpawnNPCS(1 To i - 1)
    End If
    
    .WaitingNPCS = .WaitingNPCS - 1
    End With
End Sub

Sub DeleteWaitingNPC(ByVal MapNum As Long, ByVal index As Long)

End Sub

Sub ClearMapWaitingNPCS(ByVal MapNum As Long)
    If MapNum = 0 Then Exit Sub
    
    ReDim TempMap(MapNum).WaitingForSpawnNPCS(1 To 1)
    TempMap(MapNum).WaitingNPCS = 0
    
End Sub



Function GetNumberOfMapsWithPlayers() As Long
    GetNumberOfMapsWithPlayers = NMaps
End Function

Function GetMapNumByMapReference(ByVal index As Long) As Long
    If index > 0 And index <= NMaps Then
        GetMapNumByMapReference = MapReferences(index).MapNum
    End If
End Function

Function ArePlayersOnMap(ByVal MapNum As Long) As Long
    'if map is on our structure, this means players are on it
    If MapNum = 0 Then Exit Function
    
    Dim i As Long
    i = GetMapRef(MapNum)
    If i > 0 And i <= NMaps Then
        ArePlayersOnMap = MapReferences(i).NumPlayers
    End If
End Function

Sub AddMapPlayer(ByVal index As Long, ByVal MapNum As Long)
    If index = 0 Or MapNum = 0 Then Exit Sub
    
    Dim i As Long
    i = GetMapRef(MapNum)
    If i > 0 Then
        MapReferences(i).NumPlayers = MapReferences(i).NumPlayers + 1
    Else 'map not created, do it
        i = InsertMapReference(MapNum)
        MapReferences(i).NumPlayers = 1
    End If
    
    
    If i > 0 Then
        With TempMap(MapNum)
            .PlayersOnMap.Add index
        End With
    End If
    'Dim j As Long
    'j = FindOpenMapPlayerSlot(mapnum)
    'If j > 0 Then 'slot exists
        'MapReferences(i).PlayersOnMap(j) = index
    'Else
        'MapReferences(i).PlayerSlots = MapReferences(i).PlayerSlots + 1
        'ReDim Preserve MapReferences(i).PlayersOnMap(1 To MapReferences(i).PlayerSlots)
        'MapReferences(i).PlayersOnMap(MapReferences(i).PlayerSlots) = index
    'End If
End Sub



Function FindPlayerByPos(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Long
    Dim i As Long, j As Variant
    If MapNum = 0 Then Exit Function
        With TempMap(MapNum)
            For Each j In .PlayersOnMap
                If GetPlayerX(j) = X And GetPlayerY(j) = Y Then
                    FindPlayerByPos = j
                    Exit Function
                End If
            Next
        End With
End Function

Sub DeleteMapPlayer(ByVal index As Long, ByVal MapNum As Long)
    Dim i As Long
    i = GetMapRef(MapNum)
    If i > 0 Then
        
        MapReferences(i).NumPlayers = MapReferences(i).NumPlayers - 1
        
        If MapReferences(i).NumPlayers <= 0 Then
            ClearMapReferenceByIndex i
        End If
        
        With TempMap(MapNum)
        
        Dim j As Long
        j = FindMapPlayerSlot(MapNum, index)
        
        If j > 0 Then
            .PlayersOnMap.Remove j
        End If
        
        
        End With
    End If

    
    
    
End Sub

Sub AddMapData(ByVal MapNum As Long)

End Sub

Sub DeleteMapData(ByVal MapNum As Long)
    'ZeroMemory map(mapnum), Len(map(mapnum))
End Sub


Function InsertMapReference(ByVal MapNum As Long) As Long
    If MapNum = 0 Then Exit Function
    
    'AddMapData mapnum
    Dim i As Long
    NMaps = NMaps + 1
    i = NMaps
    ReDim Preserve MapReferences(1 To i)
    
    Dim found As Boolean
    found = False
    
    If i > 1 Then
        While i > 1 And Not found
            If MapReferences(i - 1).MapNum > MapNum Then
                MapReferences(i) = MapReferences(i - 1)
                TempMap(MapReferences(i).MapNum).mapref = TempMap(MapReferences(i).MapNum).mapref + 1
                i = i - 1
            Else
                found = True
            End If
        Wend
    End If
    
    
    Call CreateMapTileReference(i, MapNum)
    InsertMapReference = i
    
    TempMap(MapNum).mapref = i
    
    Set TempMap(MapNum).PlayersOnMap = New Collection
    
End Function

Sub ClearMapReference(ByVal MapNum As Long)
    If MapNum = 0 Then Exit Sub
    Dim i As Long
    i = GetMapRef(MapNum)
    If i > 0 Then
        If NMaps > 1 Then
            Dim j As Long
            For j = i To NMaps - 1
                MapReferences(j) = MapReferences(j + 1)
                TempMap(MapReferences(j).MapNum).mapref = TempMap(MapReferences(j).MapNum).mapref - 1
            Next
            ReDim Preserve MapReferences(1 To NMaps - 1)
        End If
        
        NMaps = NMaps - 1
        
        TempMap(MapNum).mapref = 0
        
    End If
    
End Sub

Sub ClearMapReferenceByIndex(ByVal index As Long)
    If index > 0 And index <= NMaps Then
        Dim j As Long
        Dim MapNum As Long
        MapNum = MapReferences(index).MapNum
        If NMaps > 1 Then
            For j = index To NMaps - 1
                MapReferences(j) = MapReferences(j + 1)
                TempMap(MapReferences(j).MapNum).mapref = TempMap(MapReferences(j).MapNum).mapref - 1
            Next
            
            ReDim Preserve MapReferences(1 To NMaps - 1)
        End If
        
        NMaps = NMaps - 1
        
        TempMap(MapNum).mapref = 0
    End If
    
    Call DeleteMapData(MapNum)
End Sub

Public Function BinarySearchMapRef(ByVal MapNum As Long, ByVal left As Long, ByVal right As Long) As Long
    If right < left Then
        BinarySearchMapRef = 0
    Else
        Dim meddle As Integer
        meddle = (left + right) \ 2
        
        Dim CurMap As Long
        CurMap = MapReferences(meddle).MapNum
        If MapNum < CurMap Then
            BinarySearchMapRef = BinarySearchMapRef(MapNum, left, meddle - 1)
        ElseIf MapNum > CurMap Then
            BinarySearchMapRef = BinarySearchMapRef(MapNum, meddle + 1, right)
        Else
            BinarySearchMapRef = meddle
        End If

    End If
        
        
End Function
Sub CreateMapTileReference(ByVal index As Long, ByVal MapNum As Long)
    MapReferences(index).MapNum = MapNum
    Dim MaxX As Byte, MaxY As Byte
    MaxX = map(MapNum).MaxX
    MaxY = map(MapNum).MaxY
    With MapReferences(index)
    ReDim .Tiles(0 To MaxX, 0 To MaxY)
    
    Dim i As Long
    For i = 1 To GetMapNpcHighIndex(MapNum)
        MaxX = MapNpc(MapNum).NPC(i).X
        MaxY = MapNpc(MapNum).NPC(i).Y
        
        .Tiles(MaxX, MaxY).mapnpcnum = i
    Next
    
    'Create Resource Tile
    For MaxX = 0 To map(MapNum).MaxX
        For MaxY = 0 To map(MapNum).MaxY
            .Tiles(MaxX, MaxY).ResourceIndex = 0
        Next
    Next
    
    For i = 1 To ResourceCache(MapNum).Resource_Count
        MaxX = ResourceCache(MapNum).ResourceData(i).X
        MaxY = ResourceCache(MapNum).ResourceData(i).Y
        .Tiles(MaxX, MaxY).ResourceIndex = i
    Next
    
    
    
    End With
    
End Sub

Function GetMapNpcHighIndex(ByVal MapNum As Long) As Long
    If MapNum = 0 Then Exit Function
    GetMapNpcHighIndex = TempMap(MapNum).npc_highindex
End Function



Function GetMapNpcNumForClient(ByVal MapNum As Long, ByVal mapnpcnum As Long) As Long
    If MapNum = 0 Or mapnpcnum = 0 Then Exit Function
    
    'GetMapNpcNumForClient = mapnpc(mapnum).NPC(MapNPCNum).MapNPCNum
    GetMapNpcNumForClient = mapnpcnum
End Function




Function GetMapNPCX(ByVal MapNum As Long, ByVal mapnpcnum As Long) As Long
    GetMapNPCX = MapNpc(MapNum).NPC(mapnpcnum).X
End Function

Function GetMapNPCY(ByVal MapNum As Long, ByVal mapnpcnum As Long) As Long
    GetMapNPCY = MapNpc(MapNum).NPC(mapnpcnum).Y
End Function

Function SetMapNPCX(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal X As Long) As Long
    MapNpc(MapNum).NPC(mapnpcnum).X = X
End Function

Function SetMapNPCY(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal Y As Long) As Long
    MapNpc(MapNum).NPC(mapnpcnum).Y = Y
End Function


Function ComputeNPCSingleMovement(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal dir As Byte) As Boolean
    Dim mapref As Long
    mapref = GetMapRef(MapNum)
    
    If mapref < 1 Or mapref > NMaps Then Exit Function
    
    With MapReferences(mapref)
    If .MapNum <> MapNum Then Exit Function
    
    Dim X As Long
    Dim Y As Long
    X = GetMapNPCX(MapNum, mapnpcnum)
    Y = GetMapNPCY(MapNum, mapnpcnum)
    
    
    Dim Nx As Long, ny As Long
    GetNextPosition X, Y, dir, Nx, ny
    
    If OutOfBoundries(Nx, ny, MapNum) Then Exit Function
    
    MapNpc(MapNum).NPC(mapnpcnum).dir = dir
    
    If .Tiles(X, Y).mapnpcnum = mapnpcnum Then
        .Tiles(X, Y).mapnpcnum = 0
    End If
    
    SetMapNPCX MapNum, mapnpcnum, Nx
    SetMapNPCY MapNum, mapnpcnum, ny
    
    .Tiles(Nx, ny).mapnpcnum = mapnpcnum
    
    End With
    
    ComputeNPCSingleMovement = True
End Function

Sub AddNPCToMapRef(ByVal MapNum As Long, ByVal mapnpcnum As Long)
    If MapNum = 0 Or mapnpcnum = 0 Then Exit Sub
    Dim i As Long
    i = GetMapRef(MapNum)
    If i > 0 Then
        With MapReferences(i)
        
        Dim X As Long, Y As Long
        X = GetMapNPCX(MapNum, mapnpcnum)
        Y = GetMapNPCY(MapNum, mapnpcnum)
        
        .Tiles(X, Y).mapnpcnum = mapnpcnum
        
        End With
        
    
    End If
    
End Sub

Sub DeleteNPCFromMapRef(ByVal MapNum As Long, ByVal mapnpcnum As Long)
    If MapNum = 0 Or mapnpcnum = 0 Then Exit Sub
    Dim i As Long
    i = GetMapRef(MapNum)
    If i > 0 Then
        With MapReferences(i)
        
        Dim X As Long, Y As Long
        X = GetMapNPCX(MapNum, mapnpcnum)
        Y = GetMapNPCY(MapNum, mapnpcnum)
        
        If .Tiles(X, Y).mapnpcnum = mapnpcnum Then
            .Tiles(X, Y).mapnpcnum = 0
        End If
        
        
        End With
    
    End If
End Sub

Sub AddResourceIndexToMapRef(ByVal MapNum As Long, ByVal ResourceIndex As Long)
    If MapNum = 0 Or mapnpcnum = 0 Then Exit Sub
    Dim i As Long
    i = GetMapRef(MapNum)
    If i > 0 Then
        With MapReferences(i)
        
        Dim X As Long, Y As Long
        X = ResourceCache(MapNum).ResourceData(ResourceIndex).X
        Y = ResourceCache(MapNum).ResourceData(ResourceIndex).Y
        
        .Tiles(X, Y).ResourceIndex = ResourceIndex
        
        End With
        
    
    End If
    
End Sub

Sub DeleteResourceIndexFromMapRef(ByVal MapNum As Long, ByVal ResourceIndex As Long)
    If MapNum = 0 Or mapnpcnum = 0 Then Exit Sub
    Dim i As Long
    i = GetMapRef(MapNum)
    If i > 0 Then
        With MapReferences(i)
        
        Dim X As Long, Y As Long
        X = ResourceCache(MapNum).ResourceData(ResourceIndex).X
        Y = ResourceCache(MapNum).ResourceData(ResourceIndex).Y
        
        If .Tiles(X, Y).ResourceIndex = ResourceIndex Then
            .Tiles(X, Y).ResourceIndex = -1
        End If
        
        
        End With
    
    End If
End Sub


Function GetMapRefNPCNumByTile(ByVal mapref As Long, ByVal X As Long, ByVal Y As Long) As Long
    'If MapRef < 1 Or MapRef > NMaps Then Exit Function
    'If OutOfBoundries(x, y, MapReferences(MapRef).mapnum) Then Exit Function
    If mapref = 0 Or mapref > NMaps Then Exit Function
    GetMapRefNPCNumByTile = MapReferences(mapref).Tiles(X, Y).mapnpcnum

End Function

Function GetMapRefResourceIndexByTile(ByVal mapref As Long, ByVal X As Long, ByVal Y As Long) As Long
    'If MapRef < 1 Or MapRef > NMaps Then Exit Function
    'If OutOfBoundries(x, y, MapReferences(MapRef).mapnum) Then Exit Function
    GetMapRefResourceIndexByTile = 0
    If mapref = 0 Or mapref > NMaps Then Exit Function
    GetMapRefResourceIndexByTile = MapReferences(mapref).Tiles(X, Y).ResourceIndex

End Function



Function GetMapNPCNumByTile(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Long
    
    Dim i As Long
    For i = 1 To GetMapNpcHighIndex(MapNum)
        If GetMapNPCX(MapNum, i) = X And GetMapNPCY(MapNum, i) = Y Then
            GetMapNPCNumByTile = i
            Exit Function
        End If
    Next
    

End Function

Function GetMapRef(ByVal MapNum As Long) As Long
    If MapNum = 0 Then Exit Function
    GetMapRef = TempMap(MapNum).mapref
End Function

Sub CheckToAddMap(ByVal MapNum As Long)
    If GetMapRef(MapNum) = 0 Then
        AddMapData (MapNum)
    End If
End Sub

Public Function GetMapData(ByRef MapT As MapRec) As Byte()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    With MapT
    
        buffer.WriteConstString "v.2"
        buffer.WriteConstString .Name
        If Len(Trim$(Replace(.TranslatedName, vbNullChar, ""))) = 0 Then .TranslatedName = GetTranslation(.Name)
        buffer.WriteConstString .TranslatedName
        buffer.WriteConstString .Music
        
        buffer.WriteLong .revision
        buffer.WriteByte .moral
        
        buffer.WriteLong .Up
        buffer.WriteLong .Down
        buffer.WriteLong .left
        buffer.WriteLong .right
        
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
                    buffer.WriteLong .Tile(X, Y).Layer(j).X
                    buffer.WriteLong .Tile(X, Y).Layer(j).Y
                    buffer.WriteLong .Tile(X, Y).Layer(j).Tileset
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
            buffer.WriteByte .NPCSProperties(X).Movement
            buffer.WriteByte .NPCSProperties(X).Action
        Next

        buffer.WriteLong .Weather
        
        For X = 1 To Max_States - 1
            buffer.WriteByte .AllowedStates(X)
        Next
    End With
    
    GetMapData = buffer.ToArray
    Set buffer = Nothing
End Function

Public Sub SetMapData(ByRef map As MapRec, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim newVer As Boolean
    buffer.WriteBytes Data
    With map
        If buffer.ReadConstString(3, False) = "v.2" Then newVer = True: buffer.MoveReadHead 3
        .Name = buffer.ReadConstString(NAME_LENGTH)
        If newVer = True Then .TranslatedName = buffer.ReadConstString(NAME_LENGTH)
        If newVer = False Then .TranslatedName = GetTranslation(.Name)
        .Music = buffer.ReadConstString(NAME_LENGTH)
        .revision = buffer.ReadLong
        .moral = buffer.ReadByte
        .Up = buffer.ReadLong
        .Down = buffer.ReadLong
        .left = buffer.ReadLong
        .right = buffer.ReadLong
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
                    .Tile(X, Y).Layer(j).X = buffer.ReadLong
                    .Tile(X, Y).Layer(j).Y = buffer.ReadLong
                    .Tile(X, Y).Layer(j).Tileset = buffer.ReadLong
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
            .NPCSProperties(X).Movement = buffer.ReadByte
            .NPCSProperties(X).Action = buffer.ReadByte
        Next

        .Weather = buffer.ReadLong
        
        For X = 1 To Max_States - 1
            .AllowedStates(X) = buffer.ReadByte
        Next
    
    End With
    Set buffer = Nothing
End Sub

'here
Public Sub SetServerMapData(ByVal MapNum As Long, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    Dim newVer As Boolean

    With map(MapNum)
        If buffer.ReadConstString(3, False) = "v.2" Then newVer = True: buffer.MoveReadHead 3
  
        .Name = buffer.ReadConstString(NAME_LENGTH)
        If newVer = True Then .TranslatedName = buffer.ReadConstString(NAME_LENGTH)
        If newVer = False Then .TranslatedName = GetTranslation(.Name) ': buffer.MoveReadHead NAME_LENGTH
        'this is because the server doesn't need to know the music
         buffer.MoveReadHead NAME_LENGTH
        .revision = buffer.ReadLong
        .moral = buffer.ReadByte
        .Up = buffer.ReadLong
        .Down = buffer.ReadLong
        .left = buffer.ReadLong
        .right = buffer.ReadLong
        .BootMap = buffer.ReadLong
        .BootX = buffer.ReadByte
        .BootY = buffer.ReadByte
        .MaxX = buffer.ReadByte
        .MaxY = buffer.ReadByte
        ReDim .Tile(0 To .MaxX, 0 To .MaxY)
        
        'buffer.Allocate .MaxX * .MaxY * (4 * 3 + 2 * 1) + MAX_MAP_NPCS * (1 * 4 + 2 * 1) + 4 + (Max_States - 1) * 1

        For X = 0 To .MaxX
            For Y = 0 To .MaxY
                Dim j As Byte
                Call buffer.MoveReadHead(84)
                .Tile(X, Y).Type = buffer.ReadByte
                .Tile(X, Y).Data1 = buffer.ReadLong
                .Tile(X, Y).Data2 = buffer.ReadLong
                .Tile(X, Y).Data3 = buffer.ReadLong
                .Tile(X, Y).DirBlock = buffer.ReadByte
            Next
        Next

        For X = 1 To MAX_MAP_NPCS
            .NPC(X) = buffer.ReadLong
            'here
            MapNpc(MapNum).NPC(X).Num = .NPC(X)
            .NPCSProperties(X).Movement = buffer.ReadByte
            .NPCSProperties(X).Action = buffer.ReadByte
        Next

        .Weather = buffer.ReadLong
    
        For X = 1 To Max_States - 1
            .AllowedStates(X) = buffer.ReadByte
        Next
    End With
    Set buffer = Nothing
End Sub


Sub SetMapState(ByVal MapNum As Long, ByVal state As PlayerStateType, ByVal Value As Boolean)
    map(MapNum).AllowedStates(state) = Value
End Sub

Function GetMapState(ByVal MapNum As Long, ByVal state As PlayerStateType) As Boolean
    GetMapState = map(MapNum).AllowedStates(state)
End Function

Function GetMapMoral(ByVal MapNum As Long) As Byte
    GetMapMoral = map(MapNum).moral
End Function

Sub GetOnDeathMap(ByVal index As Long, ByRef MapNum As Long, ByRef X As Long, ByRef Y As Long)

If FixWarpMap_Enabled Then
    MapNum = FixWarpMap
    If MapNum > 0 And MapNum < MAX_MAPS Then
        X = RAND(0, map(MapNum).MaxX)
        Y = RAND(0, map(MapNum).MaxY)
    Else
        MapNum = Class(GetPlayerClass(index)).StartMap
        X = Class(GetPlayerClass(index)).StartMapX
        Y = Class(GetPlayerClass(index)).StartMap
    End If


Else
    With map(GetPlayerMap(index))
        ' to the bootmap if it is set
        If RespawnSite = 0 Then
            If .BootMap > 0 Then
                MapNum = .BootMap
                X = .BootX
                Y = .BootY
            Else
                MapNum = Class(GetPlayerClass(index)).StartMap
                X = Class(GetPlayerClass(index)).StartMapX
                Y = Class(GetPlayerClass(index)).StartMapY
            End If
        ElseIf RespawnSite = 1 Then
                MapNum = Class(GetPlayerClass(index)).StartMap
                X = Class(GetPlayerClass(index)).StartMapX
                Y = Class(GetPlayerClass(index)).StartMapY
        ElseIf RespawnSite = 2 Then
            If Not GetJusticeSpawnSite(GetPlayerPK(index), MapNum, X, Y) Then
                MapNum = Class(GetPlayerClass(index)).StartMap
                X = Class(GetPlayerClass(index)).StartMapX
                Y = Class(GetPlayerClass(index)).StartMapY
            End If
        End If
    End With
End If
End Sub



