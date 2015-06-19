Attribute VB_Name = "modGameEditors"
Option Explicit
' ////////////////
' // Map Editor //
' ////////////////
Public Sub MapEditorInit()
Dim i As Long
Dim smusic() As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' set the width
    frmEditor_Map.Width = 7425
    
    ' we're in the map editor
    InMapEditor = True
    
    ' show the form
    frmEditor_Map.Visible = True
    
    ' set the scrolly bars
    frmEditor_Map.scrlTileSet.Max = NumTileSets
    frmEditor_Map.fraTileSet.Caption = "Tileset: " & 1
    frmEditor_Map.scrlTileSet.value = 1
    
    ' render the tiles
    Call EditorMap_BltTileset
    
    ' set the scrollbars
    frmEditor_Map.scrlPictureY.Max = (frmEditor_Map.picBackSelect.Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.Max = (frmEditor_Map.picBackSelect.Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    MapEditorTileScroll
    
    ' set shops for the shop attribute
    frmEditor_Map.cmbShop.AddItem "None"
    For i = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem i & ": " & Shop(i).TranslatedName
    Next
    
    ' we're not in a shop
    frmEditor_Map.cmbShop.ListIndex = 0
    
    frmEditor_Map.scrlDoor.Max = MAX_DOORS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorProperties()
Dim X As Long
Dim y As Long
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    

    ' add the array to the combo
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "None."
    For i = 1 To UBound(musicCache)
        frmEditor_MapProperties.lstMusic.AddItem musicCache(i)
    Next
    ' finished populating
    
    With frmEditor_MapProperties
        .txtName.Text = Trim$(map.Name)
        .InitAllowedStates
        ' find the music we have set
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            For i = 0 To .lstMusic.ListCount
                If .lstMusic.list(i) = Trim$(map.Music) Then
                    .lstMusic.ListIndex = i
                End If
            Next
        End If
        
        ' rest of it
        .txtUp.Text = CStr(map.Up)
        .txtDown.Text = CStr(map.Down)
        .txtLeft.Text = CStr(map.Left)
        .txtRight.Text = CStr(map.Right)
        .cmbMoral.ListIndex = map.moral
        .txtBootMap.Text = CStr(map.BootMap)
        .txtBootX.Text = CStr(map.BootX)
        .txtBootY.Text = CStr(map.BootY)
        .CmbWeather.ListIndex = map.Weather
        
        'Init movement and actions
        For X = 0 To MAX_ACTIONS
            If X = 0 Then
            .cmbAction.AddItem X & ": No Action"
            Else
            .cmbAction.AddItem X & ": " & Trim$(Actions(X).TranslatedName)
            End If
        Next
        For X = 0 To MAX_MOVEMENTS
            If X = 0 Then
            .cmbMovement.AddItem X & ": No Movement"
            Else
            .cmbMovement.AddItem X & ": " & Trim$(Movements(X).Name)
            End If
        Next
        
        'Fill the first element of actions and movements
        If map.NPCSProperties(1).Action > 0 Then
            .cmbAction.ListIndex = map.NPCSProperties(1).Action
        Else
            .cmbAction.Text = "No Action"
        End If
        If map.NPCSProperties(1).movement > 0 Then
            .cmbMovement.ListIndex = map.NPCSProperties(1).movement
        Else
            .cmbMovement.Text = "No Movement"
        End If

        ' show the map npcs
        .lstNpcs.Clear
        For X = 1 To MAX_MAP_NPCS
            If map.NPC(X) > 0 Then
                .lstNpcs.AddItem X & ": " & Trim$(NPC(map.NPC(X)).TranslatedName)
            Else
                .lstNpcs.AddItem X & ": No NPC"
            End If
        Next
        .lstNpcs.ListIndex = 0
        
        ' show the npc selection combo
        .cmbNpc.Clear
        .cmbNpc.AddItem "No NPC"
        For X = 1 To MAX_NPCS
            .cmbNpc.AddItem X & ": " & Trim$(NPC(X).TranslatedName)
        Next
        
        ' set the combo box properly
        Dim tmpString() As String
        Dim NPCNum As Long
        tmpString = Split(.lstNpcs.list(.lstNpcs.ListIndex))
        NPCNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNpc.ListIndex = map.NPC(NPCNum)
    
        ' show the current map
        .lblMap.Caption = "Current map: " & GetPlayerMap(MyIndex)
        .txtMaxX.Text = map.MaxX
        .txtMaxY.Text = map.MaxY
        

    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorProperties", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False)
Dim x2 As Long, y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not multitile Then ' single
        With map.Tile(X, y)
            ' set layer
            .layer(CurLayer).X = EditorTileX
            .layer(CurLayer).y = EditorTileY
            .layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.value
        End With
    Else ' multitile
        y2 = 0 ' starting tile for y axis
        For y = CurY To CurY + EditorTileHeight - 1
            x2 = 0 ' re-set x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= map.MaxX Then
                    If y >= 0 And y <= map.MaxY Then
                        With map.Tile(X, y)
                            .layer(CurLayer).X = EditorTileX + x2
                            .layer(CurLayer).y = EditorTileY + y2
                            .layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.value
                        End With
                    End If
                End If
                x2 = x2 + 1
            Next
            y2 = y2 + 1
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSetTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal X As Long, ByVal y As Long, Optional ByVal movedMouse As Boolean = True)
Dim i As Long
Dim CurLayer As Long
Dim tmpDir As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.OptLayer(i).value Then
            CurLayer = i
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor_Map.optLayers.value Then
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer
            Else ' multi tile!
                MapEditorSetTile CurX, CurY, CurLayer, True
            End If
        ElseIf frmEditor_Map.optAttribs.value Then
            With map.Tile(CurX, CurY)
                ' blocked tile
                If frmEditor_Map.optBlocked.value Then .Type = TILE_TYPE_BLOCKED
                ' warp tile
                If frmEditor_Map.optWarp.value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                End If
                ' item spawn
                If frmEditor_Map.optItem.value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                End If
                ' npc avoid
                If frmEditor_Map.optNpcAvoid.value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' key
                If frmEditor_Map.optKey.value Then
                    .Type = TILE_TYPE_KEY
                    .Data1 = KeyEditorNum
                    .Data2 = KeyEditorTake
                    .Data3 = 0
                End If
                ' key open
                If frmEditor_Map.optKeyOpen.value Then
                    .Type = TILE_TYPE_KEYOPEN
                    .Data1 = KeyOpenEditorX
                    .Data2 = KeyOpenEditorY
                    .Data3 = 0
                End If
                ' resource
                If frmEditor_Map.optResource.value Then
                    .Type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' door
                If frmEditor_Map.optDoor.value Then
                    .Type = TILE_TYPE_DOOR
                    .Data1 = DoorEditorNum
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' npc spawn
                If frmEditor_Map.optNpcSpawn.value Then
                    .Type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .Data2 = SpawnNpcDir
                    .Data3 = 0
                End If
                ' shop
                If frmEditor_Map.optShop.value Then
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' bank
                If frmEditor_Map.optBank.value Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' heal
                If frmEditor_Map.optHeal.value Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                End If
                ' trap
                If frmEditor_Map.optTrap.value Then
                    .Type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' slide
                If frmEditor_Map.optSlide.value Then
                    .Type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                End If
                If frmEditor_Map.optScript.value Then
                    .Type = TILE_TYPE_SCRIPT
                    .Data1 = MapEditorScriptNum
                    .Data2 = GetScriptData(1)
                    .Data3 = GetScriptData(2)
                End If
                If frmEditor_Map.optIce.value Then
                    .Type = TILE_TYPE_ICE
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End If
            End With
        ElseIf frmEditor_Map.optBlock.value Then
            If movedMouse Then Exit Sub
            ' find what tile it is
            X = X - ((X \ 32) * 32)
            y = y - ((y \ 32) * 32)
            ' see if it hits an arrow
            For i = 1 To 4
                If X >= DirArrowX(i) And X <= DirArrowX(i) + 8 Then
                    If y >= DirArrowY(i) And y <= DirArrowY(i) + 8 Then
                        ' flip the value.
                        setDirBlock map.Tile(CurX, CurY).DirBlock, CByte(i), Not isDirBlocked(map.Tile(CurX, CurY).DirBlock, CByte(i))
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor_Map.optLayers.value Then
            With map.Tile(CurX, CurY)
                ' clear layer
                .layer(CurLayer).X = 0
                .layer(CurLayer).y = 0
                .layer(CurLayer).Tileset = 0
            End With
        ElseIf frmEditor_Map.optAttribs.value Then
            With map.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With

        End If
    End If

    CacheResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetScriptData(ByVal DataNum As Long) As Long
    With frmEditor_Map
    Dim s As String
    Select Case DataNum
    Case 1
        s = .txtData1
    Case 2
        s = .txtData2
    End Select
    
    If IsNumeric(s) Then
        GetScriptData = CLng(s)
    Else
        GetScriptData = 0
    End If
    
    End With
End Function

Public Sub SetAttributeInfo(ByVal Tile_Type As Long, ByVal X As Long, ByVal y As Long)
With map.Tile(X, y)
                ' blocked tile
                
                ' warp tile
                Select Case Tile_Type
                Case TILE_TYPE_BLOCKED
                    .Type = TILE_TYPE_BLOCKED
                Case TILE_TYPE_WARP
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                ' item spawn
                Case TILE_TYPE_ITEM
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                ' npc avoid
                Case TILE_TYPE_NPCAVOID
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                ' key
                Case TILE_TYPE_KEY
                    .Type = TILE_TYPE_KEY
                    .Data1 = KeyEditorNum
                    .Data2 = KeyEditorTake
                    .Data3 = 0
                ' key open
                Case TILE_TYPE_KEYOPEN
                    .Type = TILE_TYPE_KEYOPEN
                    .Data1 = KeyOpenEditorX
                    .Data2 = KeyOpenEditorY
                    .Data3 = 0
                ' resource
                Case TILE_TYPE_RESOURCE
                    .Type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                ' door
                Case TILE_TYPE_DOOR
                    .Type = TILE_TYPE_DOOR
                    .Data1 = DoorEditorNum
                    .Data2 = 0
                    .Data3 = 0
                ' npc spawn
                Case TILE_TYPE_NPCSPAWN
                    .Type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .Data2 = SpawnNpcDir
                    .Data3 = 0
                ' shop
                Case TILE_TYPE_SHOP
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .Data2 = 0
                    .Data3 = 0
                ' bank
                Case TILE_TYPE_BANK
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                ' heal
                Case TILE_TYPE_HEAL
                    .Type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                ' trap
               Case TILE_TYPE_TRAP
                    .Type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                ' slide
                Case TILE_TYPE_SLIDE
                    .Type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                Case TILE_TYPE_SCRIPT
                    .Type = TILE_TYPE_SCRIPT
                    .Data1 = MapEditorScriptNum
                    .Data2 = 0
                    .Data3 = 0
                Case TILE_TYPE_ICE
                    .Type = TILE_TYPE_ICE
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End Select
            End With
End Sub


Public Sub MapEditorChooseTile(Button As Integer, X As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = X \ PIC_X
        EditorTileY = y \ PIC_Y
        
        frmEditor_Map.shpSelected.Top = EditorTileY * PIC_Y
        frmEditor_Map.shpSelected.Left = EditorTileX * PIC_X
        
        frmEditor_Map.shpSelected.Width = PIC_X
        frmEditor_Map.shpSelected.Height = PIC_Y
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorChooseTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        X = (X \ PIC_X) + 1
        y = (y \ PIC_Y) + 1
        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > frmEditor_Map.picBackSelect.Width / PIC_X Then X = frmEditor_Map.picBackSelect.Width / PIC_X
        If y < 0 Then y = 0
        If y > frmEditor_Map.picBackSelect.Height / PIC_Y Then y = frmEditor_Map.picBackSelect.Height / PIC_Y
        ' find out what to set the width + height of map editor to
        If X > EditorTileX Then ' drag right
            EditorTileWidth = X - EditorTileX
        Else ' drag left
            ' TO DO
        End If
        If y > EditorTileY Then ' drag down
            EditorTileHeight = y - EditorTileY
        Else ' drag up
            ' TO DO
        End If
        frmEditor_Map.shpSelected.Width = EditorTileWidth * PIC_X
        frmEditor_Map.shpSelected.Height = EditorTileHeight * PIC_Y
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorDrag", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorTileScroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' horizontal scrolling
    If frmEditor_Map.picBackSelect.Width < frmEditor_Map.picBack.Width Then
        frmEditor_Map.scrlPictureX.enabled = False
    Else
        frmEditor_Map.scrlPictureX.enabled = True
        frmEditor_Map.picBackSelect.Left = (frmEditor_Map.scrlPictureX.value * PIC_X) * -1
    End If
    
    ' vertical scrolling
    If frmEditor_Map.picBackSelect.Height < frmEditor_Map.picBack.Height Then
        frmEditor_Map.scrlPictureY.enabled = False
    Else
        frmEditor_Map.scrlPictureY.enabled = True
        frmEditor_Map.picBackSelect.Top = (frmEditor_Map.scrlPictureY.value * PIC_Y) * -1
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorTileScroll", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSend()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SendMap
    InMapEditor = False
    Unload frmEditor_Map
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSend", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorCancel()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong 1
    SendData Buffer.ToArray()
    InMapEditor = False
    Unload frmEditor_Map
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearLayer()
Dim i As Long
Dim X As Long
Dim y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.OptLayer(i).value Then
            CurLayer = i
            Exit For
        End If
    Next
    
    If CurLayer = 0 Then Exit Sub

    ' ask to clear layer
    If MsgBox("Are you sure you wish to clear this layer?", vbYesNo, Options.Game_Name) = vbYes Then
        For X = 0 To map.MaxX
            For y = 0 To map.MaxY
                map.Tile(X, y).layer(CurLayer).X = 0
                map.Tile(X, y).layer(CurLayer).y = 0
                map.Tile(X, y).layer(CurLayer).Tileset = 0
            Next
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorFillLayer()
Dim i As Long
Dim X As Long
Dim y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optLayers.value = True Then
        ' find which layer we're on
        For i = 1 To MapLayer.Layer_Count - 1
            If frmEditor_Map.OptLayer(i).value Then
                CurLayer = i
                Exit For
            End If
        Next
    
        ' Ground layer
        If MsgBox("Are you sure you wish to fill this layer?", vbYesNo, Options.Game_Name) = vbYes Then
            For X = 0 To map.MaxX
                For y = 0 To map.MaxY
                    map.Tile(X, y).layer(CurLayer).X = EditorTileX
                    map.Tile(X, y).layer(CurLayer).y = EditorTileY
                    map.Tile(X, y).layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.value
                Next
            Next
        End If
    ElseIf frmEditor_Map.optAttribs.value = True Then
        i = GetAttributeChecked
        If MsgBox("Are you sure you wish to fill this layer?", vbYesNo, Options.Game_Name) = vbYes Then
            For X = 0 To map.MaxX
                For y = 0 To map.MaxY
                    Call SetAttributeInfo(i, X, y)
                Next
            Next
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorFillLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub FillEmptyAttributes()
Dim i As Long
Dim X As Long
Dim y As Long

i = GetAttributeChecked
If MsgBox("Are you sure you wish to fill empty spaces with this layer?", vbYesNo, Options.Game_Name) = vbYes Then
   For X = 0 To map.MaxX
        For y = 0 To map.MaxY
            If map.Tile(X, y).Type = TILE_TYPE_WALKABLE Then
                Call SetAttributeInfo(i, X, y)
            End If
        Next
    Next
End If
End Sub

Function GetAttributeChecked() As Long
Dim i As Long
With frmEditor_Map

If .optBlocked Then
    i = TILE_TYPE_BLOCKED
ElseIf .optWarp Then
    i = TILE_TYPE_WARP
ElseIf .optItem Then
    i = TILE_TYPE_ITEM
ElseIf .optNpcAvoid Then
    i = TILE_TYPE_NPCAVOID
ElseIf .optKey Then
    i = TILE_TYPE_KEY
ElseIf .optKeyOpen Then
    i = TILE_TYPE_KEYOPEN
ElseIf .optResource Then
    i = TILE_TYPE_RESOURCE
ElseIf .optDoor Then
    i = TILE_TYPE_DOOR
ElseIf .optNpcSpawn Then
    i = TILE_TYPE_NPCSPAWN
ElseIf .optShop Then
    i = TILE_TYPE_SHOP
ElseIf .optBank Then
    i = TILE_TYPE_BANK
ElseIf .optHeal Then
    i = TILE_TYPE_HEAL
ElseIf .optTrap Then
    i = TILE_TYPE_TRAP
ElseIf .optSlide Then
    i = TILE_TYPE_SLIDE
ElseIf .optScript Then
    i = TILE_TYPE_SCRIPT
ElseIf .optIce Then
    i = TILE_TYPE_ICE
End If

GetAttributeChecked = i
End With

End Function

Public Sub MapEditorClearAttribs()
Dim X As Long
Dim y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, Options.Game_Name) = vbYes Then

        For X = 0 To map.MaxX
            For y = 0 To map.MaxY
                map.Tile(X, y).Type = 0
            Next
        Next

    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearAttribs", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorLeaveMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InMapEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorLeaveMap", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
Dim i As Long
Dim SoundSet As Boolean
Dim index As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Item.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    
    ClearItemCmbType
    InitItemCmbType
    ' add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Item.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    ItemEditorInitRanges

    With Item(EditorIndex)
        frmEditor_Item.txtName.Text = Trim$(.Name)
        If .Pic > frmEditor_Item.scrlPic.Max Then .Pic = 0
        frmEditor_Item.scrlPic.value = .Pic
        frmEditor_Item.cmbType.ListIndex = .Type
        frmEditor_Item.scrlAnim.value = .Animation
        frmEditor_Item.txtDesc.Text = Trim$(.Desc)
        frmEditor_Item.scrlAmmo.value = Item(EditorIndex).ammo
        
        If Item(EditorIndex).ammoreq = 0 Or Item(EditorIndex).ammoreq = 1 Then
            frmEditor_Item.ChkAmmo.value = Item(EditorIndex).ammoreq
        End If

        ' find the sound we have set
        If frmEditor_Item.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Item.cmbSound.ListCount
                If frmEditor_Item.cmbSound.list(i) = Trim$(.Sound) Then
                    frmEditor_Item.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0
        End If

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmEditor_Item.fraEquipment.Visible = True
            frmEditor_Item.scrlDamage.value = .Data2
            frmEditor_Item.cmbTool.ListIndex = .Data3

            If .Speed < 100 Then .Speed = 100
            frmEditor_Item.scrlSpeed.value = .Speed
            
            If Item(EditorIndex).istwohander Then
            frmEditor_Item.ChkTwoh.value = 1
            Else
            frmEditor_Item.ChkTwoh.value = 0
            End If
            
            ' loop for stats
            For i = 1 To Stats.Stat_Count - 1
                frmEditor_Item.scrlStatBonus(i).value = .Add_Stat(i)
            Next
            
            frmEditor_Item.scrlPaperdoll = .Paperdoll
            If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            frmEditor_Item.Frame4.Visible = True
            With Item(EditorIndex).ProjecTile
          frmEditor_Item.scrlProjectileDamage.value = .Damage
          frmEditor_Item.scrlProjectilePic.value = .Pic
          frmEditor_Item.scrlProjectileRange.value = .range
          frmEditor_Item.scrlProjectileSpeed.value = .Speed
          frmEditor_Item.scrlDepth.value = .Depth
    End With
End If
        Else
            frmEditor_Item.fraEquipment.Visible = False
            frmEditor_Item.Frame4.Visible = False
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONSUME Then
            frmEditor_Item.scrlItem.value = .ConsumeItem
            frmEditor_Item.fraVitals.Visible = True
            frmEditor_Item.scrlAddHp.value = .AddHP
            frmEditor_Item.scrlAddMP.value = .AddMP
            frmEditor_Item.scrlAddExp.value = .AddEXP
            frmEditor_Item.scrlCastSpell.value = .CastSpell
            frmEditor_Item.chkInstant.value = .instaCast
        Else
            frmEditor_Item.fraVitals.Visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.Visible = True
            frmEditor_Item.scrlSpell.value = .Data1
        Else
            frmEditor_Item.fraSpell.Visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.Visible = True
            frmEditor_Item.scrlSpell.value = .Data1
        Else
            frmEditor_Item.fraSpell.Visible = False
        End If
        
        ' Make Containers work
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONTAINER) Then
            frmEditor_Item.frameContainer.Visible = True
            frmEditor_Item.scrlContainerIndex.value = 0
            frmEditor_Item.scrlContainer.value = .Container(0).ItemNum
            If (.Container(0).value >= 0 And .Container(0).value <= frmEditor_Item.scrlAmount.Max) Then
                frmEditor_Item.scrlAmount.value = .Container(0).value
                frmEditor_Item.lblAmount.Caption = "Amount: " & .Container(0).value
            End If
        Else
            frmEditor_Item.frameContainer.Visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_BAG) Then
            frmEditor_Item.fraBag.Visible = True
            If (.AddBags <= MAX_RUPEE_BAGS) Then
                frmEditor_Item.scrlBag.value = .AddBags
            End If
        Else
            frmEditor_Item.fraBag.Visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_ADDWEIGHT) Then
            frmEditor_Item.fraAddWeight.Visible = True
            If (.Data1 >= 0 And .Data1 <= 10000) Then
                frmEditor_Item.scrlAddWeight.value = .Data1
            End If
        Else
            frmEditor_Item.fraAddWeight.Visible = False
        End If
        
        ' Basic requirements
        frmEditor_Item.scrlAccessReq.value = .AccessReq
        frmEditor_Item.scrlLevelReq.value = .LevelReq
        
        ' loop for stats
        For i = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatReq(i).value = .Stat_Req(i)
        Next
        
        ' Build cmbClassReq
        frmEditor_Item.cmbClassReq.Clear
        frmEditor_Item.cmbClassReq.AddItem "None"

        For i = 1 To Max_Classes
            frmEditor_Item.cmbClassReq.AddItem Class(i).TranslatedName
        Next

        frmEditor_Item.cmbClassReq.ListIndex = .ClassReq
        ' Info
        frmEditor_Item.scrlPrice.value = .Price
        frmEditor_Item.cmbBind.ListIndex = .BindType
        frmEditor_Item.scrlRarity.value = .Rarity
         
        frmEditor_Item.spinImpactar.value = .Impactar.Spell
        frmEditor_Item.chkImpactarF.value = .Impactar.Auto
        frmEditor_Item.spinExtHp.value = .ExtraHP

         
        EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With

    Call EditorItem_BltItem
    Call EditorItem_BltPaperdoll
    Item_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 1 To MAX_ITEMS
        If Item_Changed(i) Then
            Item(i).Weight = CalculateItemWeight(i)
        End If
    Next

    For i = 1 To MAX_ITEMS
        If Item_Changed(i) Then
            Call SendSaveItem(i)
        End If
    Next
    
    Unload frmEditor_Item
    Editor = 0
    ClearChanged_Item
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Item()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Item", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Animation Editor //
' /////////////////
Public Sub AnimationEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Animation.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Animation.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Animation(EditorIndex)
        frmEditor_Animation.txtName.Text = Trim$(.Name)
        
        ' find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Animation.cmbSound.ListCount
                If frmEditor_Animation.cmbSound.list(i) = Trim$(.Sound) Then
                    frmEditor_Animation.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        End If
        
        For i = 0 To 1
            frmEditor_Animation.scrlSprite(i).value = .sprite(i)
            frmEditor_Animation.scrlFrameCount(i).value = .Frames(i)
            frmEditor_Animation.scrlLoopCount(i).value = .LoopCount(i)
            
            If .looptime(i) > 0 Then
                frmEditor_Animation.scrlLoopTime(i).value = .looptime(i)
            Else
                frmEditor_Animation.scrlLoopTime(i).value = 100
            End If
            
        Next
         
        EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With

    Call EditorAnim_BltAnim
    Animation_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        If Animation_Changed(i) Then
            Call SendSaveAnimation(i)
        End If
    Next
    
    Unload frmEditor_Animation
    Editor = 0
    ClearChanged_Animation
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Animation()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Animation", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ////////////////
' // Npc Editor //
' ////////////////
Public Sub NpcEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_NPC.Visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_NPC.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    Call ClearDropInfoTable
    'Multiple Drop information
    For i = 1 To MAX_NPC_DROPS
            'Testing if item has been defined
            If NPC(EditorIndex).Drops(i).DropItem > 0 Then
                DropsInfo(i).Number = NPC(EditorIndex).Drops(i).DropItem
                DropsInfo(i).Chances = NPC(EditorIndex).Drops(i).DropChance
                DropsInfo(i).value = NPC(EditorIndex).Drops(i).DropItemValue
                'De alguna manera hay que poder distribuir todos los items?, lo mejor es utilizar una tabla
        End If
    Next
    
    With frmEditor_NPC
        .txtName.Text = Trim$(NPC(EditorIndex).Name)
        .txtAttackSay.Text = Trim$(NPC(EditorIndex).AttackSay)
        If NPC(EditorIndex).sprite < 0 Or NPC(EditorIndex).sprite > .scrlSprite.Max Then NPC(EditorIndex).sprite = 0
        .scrlSprite.value = NPC(EditorIndex).sprite
        .txtSpawnSecs.Text = CStr(NPC(EditorIndex).SpawnSecs)
        .cmbBehaviour.ListIndex = NPC(EditorIndex).Behaviour
        .scrlRange.value = NPC(EditorIndex).range
        
        SpellIndex = .scrlSpellNum.value
        .fraSpell.Caption = "Spell - " & SpellIndex
        .scrlSpell.value = NPC(EditorIndex).Spell(SpellIndex)
        
        
        .txtHP.Text = NPC(EditorIndex).HP
        .txtEXP.Text = NPC(EditorIndex).Exp
        .txtLevel.Text = NPC(EditorIndex).Level
        .txtDamage.Text = NPC(EditorIndex).Damage
        .chkQuest.value = NPC(EditorIndex).Quest
        .scrlQuest.value = NPC(EditorIndex).QuestNum
        If NPC(EditorIndex).Speed >= .scrlNpcSpeed.Min And NPC(EditorIndex).Speed <= .scrlNpcSpeed.Max Then
            .scrlNpcSpeed.value = NPC(EditorIndex).Speed
        End If
        .scrlSpellNum.Max = MAX_NPC_SPELLS
        .scrlSpellNum.value = 1
        
        'Init editor with dropped item nº 1
        .txtChance.Text = CStr(DropsInfo(1).Chances)
        .scrlNum.value = DropsInfo(1).Number
        .scrlValue.value = DropsInfo(1).value
        
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.list(i) = Trim$(NPC(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
        
        For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).value = NPC(EditorIndex).stat(i)
        Next
    End With
    
    Call EditorNpc_BltSprite
    NPC_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        If NPC_Changed(i) Then
            Call SendSaveNpc(i)
        End If
    Next
    
    Unload frmEditor_NPC
    Editor = 0
    ClearChanged_NPC
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNpcs
    SendRequestNPCS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_NPC()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory NPC_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_NPC", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ////////////////
' // Resource Editor //
' ////////////////
Public Sub ResourceEditorInit()
Dim i As Long
Dim SoundSet As Boolean

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Resource.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Resource.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    'Clear the table
    ClearRewardsInfoTable
    'Fill the table
    For i = 1 To MAX_RESOURCE_REWARDS
        If Resource(EditorIndex).Rewards(i).Reward > 0 Then
            RewardsInfo(i).Chance = Resource(EditorIndex).Rewards(i).Chance
            RewardsInfo(i).Reward = Resource(EditorIndex).Rewards(i).Reward
        End If
    Next
    
    With frmEditor_Resource
        .scrlRewards.value = .scrlRewards.Min
        .scrlExhaustedPic.Max = NumResources
        .scrlNormalPic.Max = NumResources
        .scrlAnimation.Max = MAX_ANIMATIONS
        
        .txtName.Text = Trim$(Resource(EditorIndex).Name)
        .txtMessage.Text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtMessage2.Text = Trim$(Resource(EditorIndex).EmptyMessage)
        .cmbType.ListIndex = Resource(EditorIndex).ResourceType
        .scrlNormalPic.value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.value = Resource(EditorIndex).ExhaustedImage
        '.scrlReward.Value = Resource(EditorIndex).ItemReward
        .scrlTool.value = Resource(EditorIndex).ToolRequired
        .scrlHealth.value = Resource(EditorIndex).health
        .scrlRespawn.value = Resource(EditorIndex).RespawnTime
        .scrlAnimation.value = Resource(EditorIndex).Animation
        .chkWalkableNormal = BTI(Resource(EditorIndex).WalkableNormal)
        .chkWalkableExhausted = BTI(Resource(EditorIndex).WalkableExhausted)
        .scrlReward.value = RewardsInfo(1).Reward
        .scrlPercent.value = RewardsInfo(1).Chance
        If Resource(EditorIndex).Rewards(1).RewardType > 0 And Resource(EditorIndex).Rewards(1).RewardType <= .cmbRewardType.ListCount Then
            .cmbRewardType.ListIndex = Resource(EditorIndex).Rewards(1).RewardType - 1
        End If
        .chkMessage.value = BTI(Resource(EditorIndex).ItemSuccessMessage)
        
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.list(i) = Trim$(Resource(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
        
    Call EditorResource_BltSprite
    
    Resource_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        If Resource_Changed(i) Then
            Call SendSaveResource(i)
        End If
    Next
    
    Unload frmEditor_Resource
    Editor = 0
    ClearChanged_Resource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Resource()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Resource", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
Dim i As Long
    
    ClearCmbPricetype
    InitCmbPriceType
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Shop.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    
    frmEditor_Shop.txtName.Text = Trim$(Shop(EditorIndex).Name)
    If Shop(EditorIndex).BuyRate > 0 Then
        frmEditor_Shop.scrlBuy.value = Shop(EditorIndex).BuyRate
    Else
        frmEditor_Shop.scrlBuy.value = 100
    End If
    
    frmEditor_Shop.cmbItem.Clear
    frmEditor_Shop.cmbItem.AddItem "None"
    frmEditor_Shop.cmbCostItem.Clear
    frmEditor_Shop.cmbCostItem.AddItem "None"

    For i = 1 To MAX_ITEMS
        frmEditor_Shop.cmbItem.AddItem i & ": " & Trim$(Item(i).TranslatedName)
        frmEditor_Shop.cmbCostItem.AddItem i & ": " & Trim$(Item(i).TranslatedName)
    Next

    frmEditor_Shop.cmbItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem.ListIndex = 0
    
    If Shop(EditorIndex).PriceType >= 0 And Shop(EditorIndex).PriceType < ShopPricesTypeCount Then
        frmEditor_Shop.cmbPriceType.ListIndex = Shop(EditorIndex).PriceType
    End If
    
    UpdateShopTrade
    
    Shop_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Shop.lstTradeItem.Clear

    For i = 1 To MAX_TRADES
        With Shop(EditorIndex).TradeItem(i)
            ' if none, show as none
            If .Item = 0 And .CostItem = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Empty Trade Slot"
            Else
                If Shop(EditorIndex).PriceType = SHItem Then
                    frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).TranslatedName) & " for " & .CostValue & "x " & Trim$(Item(.CostItem).TranslatedName)
                Else
                    frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).TranslatedName) & " for " & .CostValue & " " & GetTranslation(" puntos")
                End If
            End If
        End With
    Next

    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateShopTrade", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        If Shop_Changed(i) Then
            Call SendSaveShop(i)
        End If
    Next
    
    Unload frmEditor_Shop
    Editor = 0
    ClearChanged_Shop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Shop()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Shop", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' //////////////////
' // Spell Editor //
' //////////////////
Public Sub SpellEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Spell.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Spell.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_Spell
        ' set max values
        .scrlAnimCast.Max = MAX_ANIMATIONS
        .scrlAnim.Max = MAX_ANIMATIONS
        .scrlAOE.Max = MAX_BYTE
        .scrlRange.Max = MAX_BYTE
        .scrlMap.Max = MAX_MAPS
        .scrlIcon.Max = NumSpellIcons
        
        ' build class combo
        .cmbClass.Clear
        .cmbClass.AddItem "None"
        For i = 1 To Max_Classes
            .cmbClass.AddItem Trim$(Class(i).TranslatedName)
        Next
        .cmbClass.ListIndex = Spell(EditorIndex).ClassReq

        
        ' set values
        .txtName.Text = Trim$(Spell(EditorIndex).Name)
        .txtDesc.Text = Trim$(Spell(EditorIndex).Desc)
        .cmbType.ListIndex = Spell(EditorIndex).Type
        .cmdPercent.value = BTI(Spell(EditorIndex).UsePercent)
        CheckMPScroll Spell(EditorIndex).UsePercent
        .scrlMP.value = Spell(EditorIndex).MPCost
        .scrlLevel.value = Spell(EditorIndex).LevelReq
        .scrlAccess.value = Spell(EditorIndex).AccessReq
        .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        .scrlCast.value = Spell(EditorIndex).CastTime
        .scrlCool.value = Spell(EditorIndex).CDTime
        .scrlIcon.value = Spell(EditorIndex).Icon
        .scrlMap.value = Spell(EditorIndex).map
        .scrlX.value = Spell(EditorIndex).X
        .scrlY.value = Spell(EditorIndex).y
        .scrlDir.value = Spell(EditorIndex).dir
        .scrlVital.value = Spell(EditorIndex).vital
        .scrlDuration.value = Spell(EditorIndex).Duration
        .scrlInterval.value = Spell(EditorIndex).Interval
        .scrlRange.value = Spell(EditorIndex).range
        If Spell(EditorIndex).IsAoE Then
            .chkAOE.value = 1
        Else
            .chkAOE.value = 0
        End If
        .scrlAOE.value = Spell(EditorIndex).AoE
        .scrlAnimCast.value = Spell(EditorIndex).CastAnim
        .scrlAnim.value = Spell(EditorIndex).SpellAnim
        .scrlStun.value = Spell(EditorIndex).StunDuration
        
        .scrlStat.Min = 0
        .scrlStat.Max = Stats.Stat_Count - 1
        If Spell(EditorIndex).stat < Stats.Stat_Count Then
            .scrlStat.value = Spell(EditorIndex).stat
        End If
        
        .scrlPlayerActions.Min = 1
        .scrlPlayerActions.Max = PlayerActions_Count - 1
        .scrlPlayerActions.value = 1
        If 0 <= Spell(EditorIndex).BlockActions(.scrlPlayerActions.value) <= 1 Then
            .chkPlayerAction.value = BTI(Spell(EditorIndex).BlockActions(.scrlPlayerActions.value))
            .lblPlayerAction.Caption = "Action: " & ActionToStr(.scrlPlayerActions.value)
        Else
            .chkPlayerAction.value = False
            .lblPlayerAction.Caption = "Action: " & ActionToStr(.scrlPlayerActions.value)
        End If
        
        
            
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.list(i) = Trim$(Spell(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
        
        ' build spell damage/defense combo
        ' stat damage
        .cmbStatDamage.Clear

            .cmbStatDamage.AddItem "Strength"
            .cmbStatDamage.AddItem "Endurance"
            .cmbStatDamage.AddItem "Intelligence"
            .cmbStatDamage.AddItem "Agility"
            .cmbStatDamage.AddItem "Willpower"
        .cmbStatDamage.ListIndex = Spell(EditorIndex).StatDamage
        
        ' stat defense
        .cmbStatDefense.Clear

            .cmbStatDefense.AddItem "Strength"
            .cmbStatDefense.AddItem "Endurance"
            .cmbStatDefense.AddItem "Intelligence"
            .cmbStatDefense.AddItem "Agility"
            .cmbStatDefense.AddItem "Willpower"
        .cmbStatDefense.ListIndex = Spell(EditorIndex).StatDefense
        
        .InitChangeStates
    End With
    
    EditorSpell_BltIcon
    
    Spell_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        If Spell_Changed(i) Then
            Call SendSaveSpell(i)
        End If
    Next
    
    Unload frmEditor_Spell
    Editor = 0
    ClearChanged_Spell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Spell()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Spell", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearAttributeDialogue()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Map.fraNpcSpawn.Visible = False
    frmEditor_Map.fraResource.Visible = False
    frmEditor_Map.fraMapItem.Visible = False
    frmEditor_Map.fraMapKey.Visible = False
    frmEditor_Map.fraKeyOpen.Visible = False
    frmEditor_Map.fraMapWarp.Visible = False
    frmEditor_Map.fraShop.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAttributeDialogue", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'/////////
'//DOORS//
'/////////

Public Sub DoorEditorInit()
        If frmEditor_Doors.Visible = False Then Exit Sub
        EditorIndex = frmEditor_Doors.lstIndex.ListIndex + 1
   
        With frmEditor_Doors
   
                .txtName.Text = Doors(EditorIndex).Name
                If Doors(EditorIndex).DoorType = 0 Then
                        .optDoor(0).value = True
                ElseIf Doors(EditorIndex).DoorType = 1 Then
                        .optDoor(1).value = True
                ElseIf Doors(EditorIndex).DoorType = 2 Then
                        .optDoor(2).value = True
                End If
                .scrlKey.value = Doors(EditorIndex).key
                .scrlSwitch.value = Doors(EditorIndex).Switch
                .scrlMap.value = Doors(EditorIndex).WarpMap
                .scrlX.value = Doors(EditorIndex).WarpX
                .scrlY.value = Doors(EditorIndex).WarpY
                If Doors(EditorIndex).UnlockType = 0 Then
                        .OptUnlock(0).value = True
                ElseIf Doors(EditorIndex).UnlockType = 1 Then
                        .OptUnlock(1).value = True
                Else
                        .OptUnlock(2).value = True
                End If
                
                .scrlTime = Doors(EditorIndex).Time
                
                .chkInitialState = BTI(Doors(EditorIndex).InitialState)
           
        End With
        Door_Changed(EditorIndex) = True
End Sub

Public Sub DoorEditorOk()
        Dim i As Long

        For i = 1 To MAX_DOORS
                If Door_Changed(i) Then
                        Call SendSavedoor(i)
                End If
        Next
   
        Unload frmEditor_Doors
        Editor = 0
        ClearChanged_Doors
End Sub

Public Sub DoorEditorCancel()
        Editor = 0
        Unload frmEditor_Doors
        ClearChanged_Doors
        ClearDoors
        SendRequestDoors
End Sub

Public Sub ClearChanged_Doors()
        ZeroMemory Door_Changed(1), MAX_DOORS * 2 ' 2 = boolean length
End Sub

Public Sub ClearDropInfoTable()
Dim i As Byte

For i = 1 To MAX_NPC_DROPS
    DropsInfo(i).Chances = 0
    DropsInfo(i).Number = 0
    DropsInfo(i).value = 0
Next
 
End Sub

Public Sub ClearRewardsInfoTable()
Dim i As Byte

For i = 1 To MAX_RESOURCE_REWARDS
    RewardsInfo(i).Chance = 0
    RewardsInfo(i).Reward = 0
Next
 
End Sub

Public Sub MovementsEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Movements.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Movements.lstIndex.ListIndex + 1
    
    

    With frmEditor_Movements
    
    
        
        .txtMovementName.Text = Trim$(Movements(EditorIndex).Name)
        Select Case Movements(EditorIndex).Type
            Case Onlydirectional
            
                .cmbMovementType.Text = "OnlyDirectional" 'Only Directional
                If Movements(EditorIndex).MovementsTable.nelem > 0 Then
                    .cmbOnlyDirectionDir.Text = DirtoStr(Movements(EditorIndex).MovementsTable.vect(1).Data.direction)
                End If
                
                .frmCustom.Visible = False
                .frmOnlyDirectional.Visible = True
                .cmdAddtoList.Visible = False
    
            
            Case ByDirection
                
                .cmbMovementType.Text = "Custom" 'custom
                .cmbCustomType.Text = "ByDirection" 'bydirection
                .frmCustom.Visible = True
                .frmNumMovements.Visible = True
                .frmDirection.Visible = True
                .frmNumTiles.Visible = False
                .frmOnlyDirectional.Visible = False
                If Movements(EditorIndex).MovementsTable.nelem > 0 Then
                    .scrlMovements.Min = 1
                    .scrlMovements.Max = Movements(EditorIndex).MovementsTable.nelem
                    .cmbDirection.Text = DirtoStr(Movements(EditorIndex).MovementsTable.vect(1).Data.direction)
                End If
                .cmdAddtoList.Visible = True
                .chkRepeat.value = BTI(Movements(EditorIndex).Repeat)
                
            Case Bytile
            
                .cmbMovementType.Text = "Custom" 'custom
                .cmbCustomType.Text = "ByTiles" 'bytile
                .frmCustom.Visible = True
                .frmNumTiles.Visible = True
                .frmOnlyDirectional.Visible = False
                .frmNumMovements.Visible = True
                .frmDirection.Visible = True
                .frmNumTiles.Visible = True
                If Movements(EditorIndex).MovementsTable.nelem > 0 Then
                    .scrlMovements.Min = 1
                    .scrlMovements.Max = Movements(EditorIndex).MovementsTable.nelem
                    .cmbDirection.Text = DirtoStr(Movements(EditorIndex).MovementsTable.vect(1).Data.direction)
                    .lblNumTiles.Caption = "Tiles: " & Movements(EditorIndex).MovementsTable.vect(1).Data.NumberOfTiles
                End If
                .cmdAddtoList.Visible = True
                .chkRepeat.value = BTI(Movements(EditorIndex).Repeat)
            
            Case Random
            
                .cmbMovementType.Text = "Custom" 'Custom
                .cmbCustomType.Text = "Random" 'Random
                .frmCustom.Visible = True
                .frmNumMovements.Visible = False
                .frmDirection.Visible = False
                .frmNumTiles.Visible = False
                .cmdAddtoList.Visible = False
                        
            Case Else
            
                .frmOnlyDirectional.Visible = False
                .frmCustom.Visible = False
                .frmNumMovements.Visible = False
                .frmDirection.Visible = False
                .frmNumTiles.Visible = False
                .cmdAddtoList.Visible = False
                
                
            
        End Select

    End With
    Movement_Changed(EditorIndex) = True
    Exit Sub
errorhandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MovementsEditorOk()
        Dim i As Byte

        For i = 1 To MAX_MOVEMENTS
                If Movement_Changed(i) Then
                        Call LMDeleteNulls(Movements(i).MovementsTable)
                        Call LMOptimize(Movements(i).MovementsTable)
                        Call SendSavemovement(i)
                End If
        Next
   
        Unload frmEditor_Movements
        Editor = 0
        ClearChanged_Movements
End Sub

Public Sub MovementsEditorCancel()
        Editor = 0
        Unload frmEditor_Movements
        ClearChanged_Movements
        ClearMovements
        SendRequestMovements
End Sub

Public Sub ClearChanged_Movements()
        ZeroMemory Movement_Changed(1), MAX_MOVEMENTS * 2 ' 2 = boolean length
End Sub

Public Sub ActionsEditorInit()
Dim i As Long

    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Actions.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Actions.lstIndex.ListIndex + 1
    
    

    With frmEditor_Actions
        
        .scrlMap.Max = MAX_MAPS
        .scrlX.Max = MAX_BYTE 'Player must know map max X/Y, otherwise server will correct that info
        .scrlY.Max = MAX_BYTE
        .scrlVital.Max = Vitals.Vital_Count - 1
        .txtName.Text = Trim$(Actions(EditorIndex).Name)
        .cmbActionType.ListIndex = Actions(EditorIndex).Type
        'Call ClearActionTypeFrames
        
        If .txtName.Text = vbNullString Then
            Actions(EditorIndex).Data1 = 0
            Actions(EditorIndex).Data2 = 0
            Actions(EditorIndex).Data3 = 0
            Actions(EditorIndex).Data4 = 0
            .frmSubVital.Visible = False
            .frmWarp.Visible = False
            Exit Sub
        End If
        
        Select Case Actions(EditorIndex).Type
        Case 0 'Sub-Vital
            
            .frmSubVital.Visible = True
            If Actions(EditorIndex).Data1 > 0 And Actions(EditorIndex).Data1 < Vitals.Vital_Count Then
                .scrlVital.value = Actions(EditorIndex).Data1
                If Actions(EditorIndex).Data2 = 0 Then
                    
                    .scrlVitalNum.value = Actions(EditorIndex).Data3
                    .optVitalNum = True
                    
                ElseIf Actions(EditorIndex).Data2 = 1 Then
                    
                    .txtVitalAbstract.Text = CStr(Actions(EditorIndex).Data3)
                    .optVitalAbstract.value = True
                    
                End If
                .chkLoseExp.value = Actions(EditorIndex).Data4
            End If
            
            
            
        Case 1 'Warp
        
            If Actions(EditorIndex).Data1 > 0 Then
            .frmWarp.Visible = True
            .scrlMap.value = Actions(EditorIndex).Data1
            .scrlX.value = Actions(EditorIndex).Data2
            .scrlY.value = Actions(EditorIndex).Data3
            End If
            
        Case Else
               
        End Select
        
        .cmbMoment.ListIndex = Actions(EditorIndex).Moment

    End With
    Action_Changed(EditorIndex) = True
    Exit Sub
errorhandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ActionsEditorOk()
        Dim i As Byte

        For i = 1 To MAX_ACTIONS
                If Action_Changed(i) Then
                        Call SendSaveAction(i)
                End If
        Next
   
        Unload frmEditor_Actions
        Editor = 0
        ClearChanged_Actions
End Sub

Public Sub ActionsEditorCancel()
        Editor = 0
        Unload frmEditor_Actions
        ClearChanged_Actions
        ClearActions
        SendRequestActions
End Sub

Public Sub ClearChanged_Actions()
        ZeroMemory Door_Changed(1), MAX_ACTIONS * 2 ' 2 = boolean length
End Sub

Public Sub PetsEditorInit()
Dim i As Long

    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Pets.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Pets.lstIndex.ListIndex + 1
    
    If Pet(EditorIndex).Name = "" Then Exit Sub
        
    
    With frmEditor_Pets
    
    .scrlNpcNum.Max = MAX_NPCS
    .scrlMaxLevel.Max = MAX_LEVELS
    .scrlPoints.Max = MAX_PET_POINTS_PERLVL
    .scrlExp.Max = 15
    
        .txtName.Text = Trim$(Pet(EditorIndex).Name)
        .scrlNpcNum.value = Pet(EditorIndex).NPCNum
        .scrlMaxLevel.value = Pet(EditorIndex).MaxLevel
        .scrlPoints.value = Pet(EditorIndex).PointsProgression
        .scrlExp.value = Pet(EditorIndex).ExpProgression
        .scrlTame.value = Pet(EditorIndex).TamePoints
    End With
    
    Pet_Changed(EditorIndex) = True
    Exit Sub
errorhandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PetsEditorOk()
        Dim i As Byte

        For i = 1 To MAX_PETS
                If Pet_Changed(i) Then
                        Call SendSavePet(i)
                End If
        Next
   
        Unload frmEditor_Pets
        Editor = 0
        ClearChanged_Pets
End Sub

Public Sub PetEditorCancel()
        Editor = 0
        Unload frmEditor_Pets
        ClearChanged_Pets
        ClearPets
        SendRequestPets
End Sub

Public Sub ClearChanged_Pets()
        ZeroMemory Pet_Changed(1), MAX_PETS * 2 ' 2 = boolean length
End Sub

Public Sub CustomSpritesEditorInit()
Dim i As Long

    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_CustomSprites.Visible = False Then Exit Sub
    EditorIndex = frmEditor_CustomSprites.lstIndex.ListIndex + 1
    
    If CustomSprites(EditorIndex).Name = "" Then Exit Sub
        
    
    
    With frmEditor_CustomSprites
    
        If GetCustomSpriteNLayers(CustomSprites(EditorIndex)) = 0 Then
            Call AddEmptyLayer(CustomSprites(EditorIndex))
        End If
        
        .txtName.Text = Trim$(CustomSprites(EditorIndex).Name)
        
        .scrlLayers.Min = 1
        .scrlLayers.Max = GetCustomSpriteNLayers(CustomSprites(EditorIndex))
        .scrlLayers.value = 1
        
        .chkPlayerSprite.value = BTI(IsLayerUsingPlayerSprite(CustomSprites(EditorIndex).Layers(1)))
        .chkCenter.value = BTI(IsLayerUsingCenter(CustomSprites(EditorIndex).Layers(1)))
        
        .scrlSprite.Min = 0
        .scrlSprite.Max = NumCharacters
        .scrlSprite.value = GetLayerSprite(CustomSprites(EditorIndex).Layers(1))
        
        .scrlDir.Min = 0
        .scrlDir.Max = MAX_DIRECTIONS - 1
        .scrlDir.value = 0
        
        .scrlAnims.Min = 0
        .scrlAnims.Max = MAX_SPRITE_ANIMS - 1
        .scrlAnims.value = 0
        
        .txtX.Text = CStr(GetLayerCenterX(CustomSprites(EditorIndex).Layers(1), .scrlDir.value))
        .txtY.Text = CStr(GetLayerCenterY(CustomSprites(EditorIndex).Layers(1), .scrlDir.value))
        
        '.chkEnableAnim = BTI(IsAnimEnabled(GetSpriteLayerFixed(CustomSprites(EditorIndex).Layers(1)), 0))
    End With
    
    
    CustomSprite_Changed(EditorIndex) = True
    Exit Sub
errorhandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CustomSpritesEditorOk()
        Dim i As Byte

        For i = 1 To MAX_CUSTOM_SPRITES
                If CustomSprite_Changed(i) Then
                        Call SendSaveCustomSprite(i)
                End If
        Next
   
        Unload frmEditor_CustomSprites
        Editor = 0
        ClearChanged_CustomSprites
End Sub

Public Sub CustomSpritesEditorCancel()
        Editor = 0
        Unload frmEditor_CustomSprites
        ClearChanged_CustomSprites
        ClearCustomSprites
        SendRequestCustomSprites
End Sub

Public Sub ClearChanged_CustomSprites()
        ZeroMemory CustomSprite_Changed(1), MAX_CUSTOM_SPRITES * 2 ' 2 = boolean length
End Sub





