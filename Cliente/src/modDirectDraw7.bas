Attribute VB_Name = "modDirectDraw7"
Option Explicit
' **********************
' ** Renders graphics **
' **********************

Public DDS_TestImageSurf As DirectDrawSurface7
Public DDSD_TestImage As DDSURFACEDESC2

' DirectDraw7 Object
Public DD As DirectDraw7

' Clipper object
Public DD_Clip As DirectDrawClipper

' primary surface
Public DDS_Primary As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

' back buffer
Public DDS_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

' Used for pre-rendering
Public DDS_Map As DirectDrawSurface7
Public DDSD_Map As DDSURFACEDESC2

' gfx buffers
Public DDS_Item() As DirectDrawSurface7 ' arrays
Public DDS_Character() As DirectDrawSurface7
Public DDS_Paperdoll() As DirectDrawSurface7
Public DDS_Tileset() As DirectDrawSurface7
Public DDS_Resource() As DirectDrawSurface7
Public DDS_Animation() As DirectDrawSurface7
Public DDS_SpellIcon() As DirectDrawSurface7
Public DDS_Face() As DirectDrawSurface7
Public DDS_Projectile() As DirectDrawSurface7
Public DDS_Door As DirectDrawSurface7 ' singes
Public DDS_Blood As DirectDrawSurface7
Public DDS_Misc As DirectDrawSurface7
Public DDS_Direction As DirectDrawSurface7
Public DDS_Target As DirectDrawSurface7
Public DDS_Bars As DirectDrawSurface7
Public DDS_Rain As DirectDrawSurface7
Public DDS_Snow As DirectDrawSurface7
Public DDS_Sandstorm As DirectDrawSurface7
Public DDS_Rupee As DirectDrawSurface7
Public DDS_Hearts As DirectDrawSurface7
Public DDS_Health As DirectDrawSurface7
Public DDS_MagicBar As DirectDrawSurface7
Public DDS_MiniMap As DirectDrawSurface7

' descriptions
Public DDSD_Temp As DDSURFACEDESC2 ' arrays
Public DDSD_Item() As DDSURFACEDESC2
Public DDSD_Character() As DDSURFACEDESC2
Public DDSD_Paperdoll() As DDSURFACEDESC2
Public DDSD_Tileset() As DDSURFACEDESC2
Public DDSD_Resource() As DDSURFACEDESC2
Public DDSD_Animation() As DDSURFACEDESC2
Public DDSD_SpellIcon() As DDSURFACEDESC2
Public DDSD_Face() As DDSURFACEDESC2
Public DDSD_Projectile() As DDSURFACEDESC2
Public DDSD_Door As DDSURFACEDESC2 ' singles
Public DDSD_Blood As DDSURFACEDESC2
Public DDSD_Misc As DDSURFACEDESC2
Public DDSD_Direction As DDSURFACEDESC2
Public DDSD_Target As DDSURFACEDESC2
Public DDSD_Bars As DDSURFACEDESC2
Public DDSD_Rain As DDSURFACEDESC2
Public DDSD_Snow As DDSURFACEDESC2
Public DDSD_Sandstorm As DDSURFACEDESC2
Public DDSD_Rupee As DDSURFACEDESC2
Public DDSD_Hearts As DDSURFACEDESC2
Public DDSD_Health As DDSURFACEDESC2
Public DDSD_MagicBar As DDSURFACEDESC2
Public DDSD_MiniMap As DDSURFACEDESC2

' timers
Public Const SurfaceTimerMax As Long = 10000
Public CharacterTimer() As Long
Public PaperdollTimer() As Long
Public ItemTimer() As Long
Public ResourceTimer() As Long
Public AnimationTimer() As Long
Public SpellIconTimer() As Long
Public FaceTimer() As Long

' Number of graphic files
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public NumItems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long
Public NumProjectiles As Long

Public DDS_ChatBubble As DirectDrawSurface7
Public DDSD_ChatBubble As DDSURFACEDESC2

Dim tmrWaitBlt As Long

Dim BlockRect As RECT, WarpRect As RECT, ItemRect As RECT, ShopRect As RECT, NpcOtherRect As RECT, PlayerRect As RECT, PlayerPkRect As RECT, NpcAttackerRect As RECT, NpcShopRect As RECT, BlankRect As RECT

' ********************
' ** Initialization **
' ********************
Public Function InitDirectDraw() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Clear DD7
    Call DestroyDirectDraw
    
    ' Init Direct Draw
    Set DD = DX7.DirectDrawCreate(vbNullString)
    
    ' Windowed
    DD.SetCooperativeLevel frmMain.hwnd, DDSCL_NORMAL

    ' Init type and set the primary surface
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
        .lBackBufferCount = 1
    End With
    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMain.picScreen.hwnd
    
    ' Have the blits to the screen clipped to the picture box
    DDS_Primary.SetClipper DD_Clip
    
    ' Initialise the surfaces
    InitSurfaces
    
With BlankRect
.Top = 4
.Bottom = .Top + 4
.Left = 0
.Right = .Left + 4
End With

' Player - Norm
With PlayerRect
.Top = 0
.Bottom = .Top + 4
.Left = 4
.Right = .Left + 4
End With

' Player - PK
With PlayerPkRect
.Top = 0
.Bottom = .Top + 4
.Left = 8
.Right = .Left + 4
End With

' NPC - Others
With NpcOtherRect
.Top = 0
.Bottom = .Top + 4
.Left = 12
.Right = .Left + 4
End With

' NPC - Shopkeeper
With NpcShopRect
.Top = 0
.Bottom = .Top + 4
.Left = 16
.Right = .Left + 4
End With

' NPC - Attack when attacked
With NpcAttackerRect
.Top = 0
.Bottom = .Top + 4
.Left = 20
.Right = .Left + 4
End With

' Attributes - Block
With BlockRect
.Top = 4
.Bottom = .Top + 4
.Left = 4
.Right = .Left + 4
End With

' Attributes - Warp
With WarpRect
.Top = 4
.Bottom = .Top + 4
.Left = 8
.Right = .Left + 4
End With

' Attributes - Item
With ItemRect
.Top = 4
.Bottom = .Top + 4
.Left = 12
.Right = .Left + 4
End With

' Attributes - Shop
With ShopRect
.Top = 4
.Bottom = .Top + 4
.Left = 16
.Right = .Left + 4
End With
    
    ' We're done
    InitDirectDraw = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InitDirectDraw", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub InitSurfaces()
Dim rec As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' DirectDraw Surface memory management setting
    DDSD_Temp.lFlags = DDSD_CAPS
    DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    
    ' clear out everything for re-init
    Set DDS_BackBuffer = Nothing

    ' Initialize back buffer
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (MAX_MAPX + 3) * PIC_X
        .lHeight = (MAX_MAPY + 3) * PIC_Y
    End With
    Set DDS_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    ' load persistent surfaces
    If FileExist(App.Path & "\data\graphics\door.bmp", True) Then Call InitDDSurf("door", DDSD_Door, DDS_Door)
    If FileExist(App.Path & "\data\graphics\direction.bmp", True) Then Call InitDDSurf("direction", DDSD_Direction, DDS_Direction)
    If FileExist(App.Path & "\data\graphics\target.bmp", True) Then Call InitDDSurf("target", DDSD_Target, DDS_Target)
    If FileExist(App.Path & "\data\graphics\misc.bmp", True) Then Call InitDDSurf("misc", DDSD_Misc, DDS_Misc)
    If FileExist(App.Path & "\data\graphics\blood.bmp", True) Then Call InitDDSurf("blood", DDSD_Blood, DDS_Blood)
    If FileExist(App.Path & "\data\graphics\bars.bmp", True) Then Call InitDDSurf("bars", DDSD_Bars, DDS_Bars)
    If FileExist(App.Path & "\data\graphics\chatbubble.bmp", True) Then Call InitDDSurf("chatbubble", DDSD_ChatBubble, DDS_ChatBubble)
    If FileExist(App.Path & "\data\graphics\rain.bmp", True) Then Call InitDDSurf("rain", DDSD_Rain, DDS_Rain)
    If FileExist(App.Path & "\data\graphics\snow.bmp", True) Then Call InitDDSurf("snow", DDSD_Snow, DDS_Snow)
    If FileExist(App.Path & "\data\graphics\sandstorm.bmp", True) Then Call InitDDSurf("sandstorm", DDSD_Sandstorm, DDS_Sandstorm)
    If FileExist(App.Path & "\data\graphics\rupee.bmp", True) Then Call InitDDSurf("Rupee", DDSD_Rupee, DDS_Rupee)
    If FileExist(App.Path & "\data\graphics\hearts.bmp", True) Then Call InitDDSurf("Hearts", DDSD_Hearts, DDS_Hearts)
    If FileExist(App.Path & "\data\graphics\health.bmp", True) Then Call InitDDSurf("Health", DDSD_Health, DDS_Health)
    If FileExist(App.Path & "\data\graphics\hearts.bmp", True) Then Call InitDDSurf("MagicBar", DDSD_MagicBar, DDS_MagicBar)
    If FileExist(App.Path & "\data\graphics\minimap.bmp", True) Then Call InitDDSurf("minimap", DDSD_MiniMap, DDS_MiniMap)
     
    ' count the blood sprites
    BloodCount = DDSD_Blood.lWidth / 32
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' This sub gets the mask color from the surface loaded from a bitmap image
Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal Y As Long)
Dim TmpR As RECT
Dim TmpDDSD As DDSURFACEDESC2
Dim TmpColorKey As DDCOLORKEY

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With TmpR
        .Left = X
        .Top = Y
        .Right = X
        .Bottom = Y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .Low = TheSurface.GetLockedPixel(X, Y)
        .High = .Low
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
    TheSurface.Unlock TmpR
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetMaskColorFromPixel", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Initializing a surface, using a bitmap
Public Sub InitDDSurf(Filename As String, ByRef SurfDesc As DDSURFACEDESC2, ByRef Surf As DirectDrawSurface7)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Set path
    Filename = App.Path & GFX_PATH & Filename & GFX_EXT

    ' Destroy surface if it exist
    If Not Surf Is Nothing Then
        Set Surf = Nothing
        Call ZeroMemory(ByVal VarPtr(SurfDesc), LenB(SurfDesc))
    End If

    ' set flags
    SurfDesc.lFlags = DDSD_CAPS
    SurfDesc.ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
    
    ' init object
    Set Surf = DD.CreateSurfaceFromFile(Filename, SurfDesc)
    
    ' Set mask
    Call SetMaskColorFromPixel(Surf, 0, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitDDSurf", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function CheckSurfaces() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if we need to restore surfaces
    If Not DD.TestCooperativeLevel = DD_OK Then
        CheckSurfaces = False
    Else
        CheckSurfaces = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "CheckSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function NeedToRestoreSurfaces() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not DD.TestCooperativeLevel = DD_OK Then
        NeedToRestoreSurfaces = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "NeedToRestoreSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub ReInitDD()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call InitDirectDraw
    
    LoadTilesets
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ReInitDD", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyDirectDraw()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Unload DirectDraw
    Set DDS_Misc = Nothing
    
    For i = 1 To NumTileSets
        Set DDS_Tileset(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i))
    Next

    For i = 1 To NumItems
        Set DDS_Item(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i))
    Next

    For i = 1 To NumCharacters
        Set DDS_Character(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Character(i)), LenB(DDSD_Character(i))
    Next
    
    For i = 1 To NumPaperdolls
        Set DDS_Paperdoll(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Paperdoll(i)), LenB(DDSD_Paperdoll(i))
    Next
    
    For i = 1 To NumResources
        Set DDS_Resource(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Resource(i)), LenB(DDSD_Resource(i))
    Next
    
    For i = 1 To NumAnimations
        Set DDS_Animation(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Animation(i)), LenB(DDSD_Animation(i))
    Next
    
    For i = 1 To NumSpellIcons
        Set DDS_SpellIcon(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_SpellIcon(i)), LenB(DDSD_SpellIcon(i))
    Next
    
    For i = 1 To NumFaces
        Set DDS_Face(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Face(i)), LenB(DDSD_Face(i))
    Next
    
    For i = 1 To NumProjectiles
    Set DDS_Projectile(i) = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Projectile(i)), LenB(DDSD_Projectile(i))
    Next
    
    Set DDS_Blood = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Blood), LenB(DDSD_Blood)
    
    Set DDS_Door = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Door), LenB(DDSD_Door)
    
    Set DDS_Direction = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Direction), LenB(DDSD_Direction)
    
    Set DDS_Target = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Target), LenB(DDSD_Target)

    Set DDS_ChatBubble = Nothing
    ZeroMemory ByVal VarPtr(DDSD_ChatBubble), LenB(DDSD_ChatBubble)

    Set DDS_Rupee = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Rupee), LenB(DDSD_Rupee)
    
    Set DDS_Hearts = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Hearts), LenB(DDSD_Hearts)
    
    Set DDS_Health = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Health), LenB(DDSD_Health)
    
    Set DDS_Hearts = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Hearts), LenB(DDSD_MagicBar)
    
    Set DDS_MiniMap = Nothing
    ZeroMemory ByVal VarPtr(DDSD_MiniMap), LenB(DDSD_MiniMap)

    Set DDS_BackBuffer = Nothing
    Set DDS_Primary = Nothing
    Set DD_Clip = Nothing
    Set DD = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroyDirectDraw", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **************
' ** Blitting **
' **************
Public Sub Engine_BltFast(ByVal dx As Long, ByVal dy As Long, ByRef ddS As DirectDrawSurface7, srcRECT As RECT, trans As CONST_DDBLTFASTFLAGS)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    
    If Not ddS Is Nothing Then
        Call DDS_BackBuffer.BltFast(dx, dy, ddS, srcRECT, trans)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Engine_BltFast", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function Engine_BltToDC(ByRef Surface As DirectDrawSurface7, sRECT As DxVBLib.RECT, dRECT As DxVBLib.RECT, ByRef picBox As VB.PictureBox, Optional Clear As Boolean = True) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Clear Then
        picBox.Cls
    End If

    Call Surface.BltToDC(picBox.hDC, sRECT, dRECT)
    picBox.Refresh
    Engine_BltToDC = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Engine_BltToDC", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub BltDirection(ByVal X As Long, ByVal Y As Long)
Dim rec As DxVBLib.RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' render grid
    rec.Top = 24
    rec.Left = 0
    rec.Right = rec.Left + 32
    rec.Bottom = rec.Top + 32
    Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Direction, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' render dir blobs
    For i = 1 To 4
        rec.Left = (i - 1) * 8
        rec.Right = rec.Left + 8
        ' find out whether render blocked or not
        If Not isDirBlocked(map.Tile(X, Y).DirBlock, CByte(i)) Then
            rec.Top = 8
        Else
            rec.Top = 16
        End If
        rec.Bottom = rec.Top + 8
        'render!
        Call Engine_BltFast(ConvertMapX(X * PIC_X) + DirArrowX(i), ConvertMapY(Y * PIC_Y) + DirArrowY(i), DDS_Direction, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltDirection", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltTarget(ByVal X As Long, ByVal Y As Long)
Dim sRECT As DxVBLib.RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DDS_Target Is Nothing Then Exit Sub
    
    Width = DDSD_Target.lWidth / 2
    Height = DDSD_Target.lHeight

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = 0
        .Right = Width
    End With
    
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRECT
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(X, Y, DDS_Target, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTarget", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltHover(ByVal tType As Long, ByVal target As Long, ByVal X As Long, ByVal Y As Long)
Dim sRECT As DxVBLib.RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DDS_Target Is Nothing Then Exit Sub
    
    Width = DDSD_Target.lWidth / 2
    Height = DDSD_Target.lHeight

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = Width
        .Right = .Left + Width
    End With
    
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRECT
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    
    Call Engine_BltFast(X, Y, DDS_Target, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHover", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapTile(ByVal X As Long, ByVal Y As Long)
Dim rec As DxVBLib.RECT
Dim i As Long
Dim door_index As Long
Dim HavetoBlt As Boolean

'possibly change the tileset here with a command, maybe dodongo caverns was supposed to be a different tileset?
'I have no idea what happened..
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    HavetoBlt = True

    With map.Tile(X, Y)
        For i = MapLayer.Ground To MapLayer.Mask2
            
        'If i = MapLayer.Mask Then
        '    If Len(frmMain.txtChat.text) > 0 Then
        '        .layer(i).Tileset = CLng(frmMain.txtChat.text)
        '    End If
        'End If
        
        'Debug.Print X & "/" & Y & " " & .layer(i).Tileset
        'DoEvents
        
        ' skip opened doors
        If .Type = TILE_TYPE_DOOR Then   '.Type = TILE_TYPE_KEY Or
        
            If (i = MapLayer.Mask And TempTile(X, Y).DoorOpen = YES) Then
                HavetoBlt = False
            ElseIf i = MapLayer.Mask Then
                'door_index = Map.Tile(X, Y).Data1
                'If door_index > 0 Then
                    'If Player(MyIndex).PlayerDoors(door_index).state = 1 Then
                        'HavetoBlt = False
                    'End If
                'End If
            End If
        ElseIf .Type = TILE_TYPE_KEY Then
        
            If (i = MapLayer.Mask And TempTile(X, Y).DoorOpen = YES) Then
                HavetoBlt = False
            End If
            
        End If
         
            ' skip tile?
            If HavetoBlt And (.layer(i).Tileset > 0 And .layer(i).Tileset <= NumTileSets) And (.layer(i).X > 0 Or .layer(i).Y > 0) Then
                ' sort out rec
                rec.Top = .layer(i).Y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .layer(i).X * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            
        Next
        
        If GetTickCount Mod 500 < 250 Then
            ' skip tile?
            If (.layer(MapLayer.MaskAnim).Tileset > 0 And .layer(MapLayer.MaskAnim).Tileset <= NumTileSets) And (.layer(MapLayer.MaskAnim).X > 0 Or .layer(MapLayer.MaskAnim).Y > 0) Then
                ' sort out rec
                rec.Top = .layer(MapLayer.MaskAnim).Y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .layer(MapLayer.MaskAnim).X * PIC_X
                rec.Right = rec.Left + PIC_X
                
                If (TempTile(X, Y).DoorOpen = NO) Then
                ' render
                Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.layer(MapLayer.MaskAnim).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If
        End If
        
    End With
    
    ' Error handler
    Exit Sub
    
errorhandler:
    HandleError "BltMapTile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapFringeTile(ByVal X As Long, ByVal Y As Long)
Dim rec As DxVBLib.RECT
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With map.Tile(X, Y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
            ' skip tile if tileset isn't set
            If (.layer(i).Tileset > 0 And .layer(i).Tileset <= NumTileSets) And (.layer(i).X > 0 Or .layer(i).Y > 0) Then
                ' sort out rec
                rec.Top = .layer(i).Y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .layer(i).X * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next
        
        If GetTickCount Mod 500 < 250 Then
            ' skip tile?
            If (.layer(MapLayer.FringeAnim).Tileset > 0 And .layer(MapLayer.FringeAnim).Tileset <= NumTileSets) And (.layer(MapLayer.FringeAnim).X > 0 Or .layer(MapLayer.FringeAnim).Y > 0) Then
                ' sort out rec
                rec.Top = .layer(MapLayer.FringeAnim).Y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .layer(MapLayer.FringeAnim).X * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.layer(MapLayer.FringeAnim).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltMapFringeTile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltDoor(ByVal X As Long, ByVal Y As Long)
Dim rec As DxVBLib.RECT
Dim x2 As Long, y2 As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' sort out animation
    With TempTile(X, Y)
        If .DoorAnimate = 1 Then ' opening
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame < 4 Then
                    .DoorFrame = .DoorFrame + 1
                Else
                    .DoorAnimate = 2 ' set to closing
                End If
                .DoorTimer = GetTickCount
            End If
        ElseIf .DoorAnimate = 2 Then ' closing
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame > 1 Then
                    .DoorFrame = .DoorFrame - 1
                Else
                    .DoorAnimate = 0 ' end animation
                End If
                .DoorTimer = GetTickCount
            End If
        End If
        
        If .DoorFrame = 0 Then .DoorFrame = 1
    End With

    With rec
        .Top = 0
        .Bottom = DDSD_Door.lHeight
        .Left = ((TempTile(X, Y).DoorFrame - 1) * (DDSD_Door.lWidth / 4))
        .Right = .Left + (DDSD_Door.lWidth / 4)
    End With

    x2 = (X * PIC_X)
    y2 = (Y * PIC_Y) - (DDSD_Door.lHeight / 2) + 4
    Call DDS_BackBuffer.BltFast(ConvertMapX(x2), ConvertMapY(y2), DDS_Door, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltDoor", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltBlood(ByVal index As Long)
Dim rec As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Blood(index)
        ' check if we should be seeing it
        If .Timer + 20000 < GetTickCount Then Exit Sub
        
        rec.Top = 0
        rec.Bottom = PIC_Y
        rec.Left = (.sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        
        Engine_BltFast ConvertMapX(.X * PIC_X), ConvertMapY(.Y * PIC_Y), DDS_Blood, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBlood", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltAnimation(ByVal index As Long, ByVal layer As Long)
Dim sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
Dim i As Long
Dim Width As Long, Height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim X As Long, Y As Long
Dim lockindex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If AnimInstance(index).Animation = 0 Then
        ClearAnimInstance index
        Exit Sub
    End If
    
    sprite = Animation(AnimInstance(index).Animation).sprite(layer)
    
    If sprite < 1 Or sprite > NumAnimations Then Exit Sub
    
    FrameCount = Animation(AnimInstance(index).Animation).Frames(layer)
    
    AnimationTimer(sprite) = GetTickCount + SurfaceTimerMax
    
    If DDS_Animation(sprite) Is Nothing Then
        Call InitDDSurf("animations\" & sprite, DDSD_Animation(sprite), DDS_Animation(sprite))
    End If
    
    ' total width divided by frame count
    Width = DDSD_Animation(sprite).lWidth / FrameCount
    Height = DDSD_Animation(sprite).lHeight
    
    sRECT.Top = 0
    sRECT.Bottom = Height
    sRECT.Left = (AnimInstance(index).FrameIndex(layer) - 1) * Width
    sRECT.Right = sRECT.Left + Width
    
    ' change x or y if locked
    If AnimInstance(index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).XOffset
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).YOffset
                Else
                    ClearAnimInstance index
                End If
            Else
                ClearAnimInstance index
            End If
        ElseIf AnimInstance(index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then
                ' check if alive
                If MapNpc(lockindex).vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).XOffset
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).YOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        X = (AnimInstance(index).X * 32) + 16 - (Width / 2)
        Y = (AnimInstance(index).Y * 32) + 16 - (Height / 2)
    End If
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    ' Clip to screen
    If Y < 0 Then

        With sRECT
            .Top = .Top - Y
        End With

        Y = 0
    End If

    If X < 0 Then

        With sRECT
            .Left = .Left - X
        End With

        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    
    Call Engine_BltFast(X, Y, DDS_Animation(sprite), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltAnimation", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltItem(ByVal ItemNum As Long)
Dim PicNum As Long
Dim rec As DxVBLib.RECT
Dim Maxframes As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if it's not us then don't render
    If MapItem(ItemNum).PlayerName <> vbNullString Then
        If MapItem(ItemNum).PlayerName <> (GetPlayerName(MyIndex)) Then Exit Sub
    End If
    
    ' get the picture
    PicNum = Item(MapItem(ItemNum).num).Pic

    If PicNum < 1 Or PicNum > NumItems Then Exit Sub
    ItemTimer(PicNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(PicNum) Is Nothing Then
        Call InitDDSurf("items\" & PicNum, DDSD_Item(PicNum), DDS_Item(PicNum))
    End If

    If DDSD_Item(PicNum).lWidth > 64 Then ' has more than 1 frame
        With rec
            .Top = 0
            .Bottom = 32
            .Left = (MapItem(ItemNum).Frame * 32)
            .Right = .Left + 32
        End With
    Else
        With rec
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    End If

    Call Engine_BltFast(ConvertMapX(MapItem(ItemNum).X * PIC_X), ConvertMapY(MapItem(ItemNum).Y * PIC_Y), DDS_Item(PicNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ScreenshotMap()
Dim X As Long, Y As Long, i As Long, rec As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' clear the surface
    Set DDS_Map = Nothing
    
    ' Initialize it
    With DDSD_Map
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (map.MaxX + 1) * 32
        .lHeight = (map.MaxY + 1) * 32
    End With
    Set DDS_Map = DD.CreateSurface(DDSD_Map)
    
    ' render the tiles
    For X = 0 To map.MaxX
        For Y = 0 To map.MaxY
            With map.Tile(X, Y)
                For i = MapLayer.Ground To MapLayer.Mask2
                    ' skip tile?
                    If (.layer(i).Tileset > 0 And .layer(i).Tileset <= NumTileSets) And (.layer(i).X > 0 Or .layer(i).Y > 0) Then
                        ' sort out rec
                        rec.Top = .layer(i).Y * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        rec.Left = .layer(i).X * PIC_X
                        rec.Right = rec.Left + PIC_X
                        ' render
                        DDS_Map.BltFast X * PIC_X, Y * PIC_Y, DDS_Tileset(.layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                Next
            End With
        Next
    Next
    
    ' render the resources
    For Y = 0 To map.MaxY
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).Y = Y Then
                            Call BltMapResource(i, True)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' render the tiles
    For X = 0 To map.MaxX
        For Y = 0 To map.MaxY
            With map.Tile(X, Y)
                For i = MapLayer.Fringe To MapLayer.Fringe2
                    ' skip tile?
                    If (.layer(i).Tileset > 0 And .layer(i).Tileset <= NumTileSets) And (.layer(i).X > 0 Or .layer(i).Y > 0) Then
                        ' sort out rec
                        rec.Top = .layer(i).Y * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        rec.Left = .layer(i).X * PIC_X
                        rec.Right = rec.Left + PIC_X
                        ' render
                        DDS_Map.BltFast X * PIC_X, Y * PIC_Y, DDS_Tileset(.layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                Next
            End With
        Next
    Next
    
    ' dump and save
    frmMain.picSSMap.Width = DDSD_Map.lWidth
    frmMain.picSSMap.Height = DDSD_Map.lHeight
    rec.Top = 0
    rec.Left = 0
    rec.Bottom = DDSD_Map.lHeight
    rec.Right = DDSD_Map.lWidth
    Engine_BltToDC DDS_Map, rec, rec, frmMain.picSSMap
    SavePicture frmMain.picSSMap.Image, App.Path & "\map" & GetPlayerMap(MyIndex) & ".jpg"
    
    
    ' let them know we did it
    AddText GetTranslation("Foto del mapa #") & GetPlayerMap(MyIndex) & " " & GetTranslation(" guardada."), BrightGreen, False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScreenshotMap", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapResource(ByVal Resource_Num As Long, Optional ByVal screenShot As Boolean = False)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim rec As DxVBLib.RECT
Dim X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource_Num > Resource_Index Then Exit Sub
    ' make sure it's not out of map
    If MapResource(Resource_Num).X > map.MaxX Then Exit Sub
    If MapResource(Resource_Num).Y > map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = map.Tile(MapResource(Resource_Num).X, MapResource(Resource_Num).Y).Data1
    
    If Resource_master = 0 Then Exit Sub

    'BUG[If Resource(Resource_master).ResourceImage = 0 Then Exit Sub]BUG'
    ' Get the Resource state
    Resource_state = MapResource(Resource_Num).ResourceState
    
    If Resource_master > MAX_RESOURCES Then Exit Sub

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    
    ' cut down everything if we're editing
    If InMapEditor Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' Load early
    If DDS_Resource(Resource_sprite) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource_sprite, DDSD_Resource(Resource_sprite), DDS_Resource(Resource_sprite))
    End If

    ' src rect
    With rec
        .Top = 0
        .Bottom = DDSD_Resource(Resource_sprite).lHeight
        .Left = 0
        .Right = DDSD_Resource(Resource_sprite).lWidth
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_Num).X * PIC_X) - (DDSD_Resource(Resource_sprite).lWidth / 2) + 16
    Y = (MapResource(Resource_Num).Y * PIC_Y) - DDSD_Resource(Resource_sprite).lHeight + 32
    
    ' render it
    If Not screenShot Then
        Call BltResource(Resource_sprite, X, Y, rec)
    Else
        Call ScreenshotResource(Resource_sprite, X, Y, rec)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltMapResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltResource(ByVal Resource As Long, ByVal dx As Long, dy As Long, rec As DxVBLib.RECT)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long
Dim destRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    ResourceTimer(Resource) = GetTickCount + SurfaceTimerMax

    If DDS_Resource(Resource) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource, DDSD_Resource(Resource), DDS_Resource(Resource))
    End If

    X = ConvertMapX(dx)
    Y = ConvertMapY(dy)
    
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If

    ' End clipping
    Call Engine_BltFast(X, Y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScreenshotResource(ByVal Resource As Long, ByVal X As Long, Y As Long, rec As DxVBLib.RECT)
Dim Width As Long
Dim Height As Long
Dim destRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    ResourceTimer(Resource) = GetTickCount + SurfaceTimerMax

    If DDS_Resource(Resource) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource, DDSD_Resource(Resource), DDS_Resource(Resource))
    End If
    
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_Map.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_Map.lHeight)
    End If

    If X + Width > DDSD_Map.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_Map.lWidth)
    End If

    ' End clipping
    'Call Engine_BltFast(x, y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    DDS_Map.BltFast X, Y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScreenshotResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltBars()
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim sRECT As RECT
Dim barWidth As Long
Dim i As Long, NPCNum As Long, partyIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' dynamic bar calculations
    sWidth = DDSD_Bars.lWidth
    sHeight = DDSD_Bars.lHeight / 6
    
    ' render health bars
    For i = 1 To MAX_MAP_NPCS
        NPCNum = MapNpc(i).num
        ' exists?
        If NPCNum > 0 Then
            ' alive?
            'If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < NPC(NPCNum).HP Then
            If MapNpc(i).vital(Vitals.HP) > 0 And MapNpc(i).vital(Vitals.HP) < GetNpcMaxVital(MapNpc(i).num, Vitals.HP, MapNpc(i).petData.Owner) Then
                ' lock to npc
                tmpX = MapNpc(i).X * PIC_X + MapNpc(i).XOffset + 16 - (sWidth / 2)
                tmpY = MapNpc(i).Y * PIC_Y + MapNpc(i).YOffset + 35
                
                ' calculate the width to fill
                barWidth = ((MapNpc(i).vital(Vitals.HP) / sWidth) / (GetNpcMaxVital(MapNpc(i).num, Vitals.HP, MapNpc(i).petData.Owner) / sWidth)) * sWidth
                
                ' draw bar background
                With sRECT
                    .Top = sHeight * 1 ' HP bar background
                    .Left = 0
                    .Right = .Left + sWidth
                    .Bottom = .Top + sHeight
                End With
                Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                
                ' draw the bar proper
                With sRECT
                    .Top = 0 ' HP bar
                    .Left = 0
                    .Right = .Left + barWidth
                    .Bottom = .Top + sHeight
                End With
                Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).YOffset + 35 + sHeight + 1
            
            ' calculate the width to fill
            barWidth = (GetTickCount - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * sWidth
            
            ' draw bar background
            With sRECT
                .Top = sHeight * 3 ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .Top + sHeight
            End With
            Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            
            ' draw the bar proper
            With sRECT
                .Top = sHeight * 2 ' cooldown bar
                .Left = 0
                .Right = barWidth
                .Bottom = .Top + sHeight
            End With
            Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).YOffset + 35
       
        ' calculate the width to fill
        barWidth = ((GetPlayerVital(MyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / sWidth)) * sWidth
       
        ' draw bar background
        With sRECT
            .Top = sHeight * 1 ' HP bar background
            .Left = 0
            .Right = .Left + sWidth
            .Bottom = .Top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
       
        ' draw the bar proper
        With sRECT
            .Top = 0 ' HP bar
            .Left = 0
            .Right = .Left + barWidth
            .Bottom = .Top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End If
    
    ' draw party health bars
    If Party.Leader > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(i)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + Player(partyIndex).XOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + Player(partyIndex).YOffset + 35
                    
                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(partyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / sWidth)) * sWidth
                    
                    ' draw bar background
                    With sRECT
                        .Top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    
                    ' draw the bar proper
                    With sRECT
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .Top + sHeight
                    End With
                    Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            End If
        Next
    End If
    
        If GetRideStamina(MyIndex) < MAX_STAMINA Then
            'tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
            'tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).YOffset + 35 + sHeight + 1
    
                tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
                tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).YOffset + 35 + sHeight * 2 + 2
                
                ' calculate the width to fill
                barWidth = ((GetRideStamina(MyIndex) / sWidth) / ((MAX_STAMINA) / sWidth)) * sWidth
                
                ' draw bar background
                With sRECT
                    .Top = sHeight * 1 ' HP bar background
                    .Left = 0
                    .Right = .Left + sWidth
                    .Bottom = .Top + sHeight
                End With
                Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                
                ' draw the bar proper
                With sRECT
                    If IsSpecialBonusActive(MyIndex) Then
                        .Top = sHeight * 5
                    Else
                        .Top = sHeight * 4 ' HP bar
                    End If
                    .Left = 0
                    .Right = .Left + barWidth
                    .Bottom = .Top + sHeight
                End With
                Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
          End If

    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBars", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltHotbar()
Dim sRECT As RECT, dRECT As RECT, i As Long, num As String, N As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.picHotbar.Cls
    For i = 1 To MAX_HOTBAR
        With dRECT
            .Top = HotbarTop
            .Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
            .Bottom = .Top + 32
            .Right = .Left + 32
        End With
        
        With sRECT
            .Top = 0
            .Left = 32
            .Bottom = 32
            .Right = 64
        End With
        
        Select Case Hotbar(i).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(i).Slot).Name) > 0 Then
                    If Item(Hotbar(i).Slot).Pic > 0 Then
                        If DDS_Item(Item(Hotbar(i).Slot).Pic) Is Nothing Then
                            Call InitDDSurf("Items\" & Item(Hotbar(i).Slot).Pic, DDSD_Item(Item(Hotbar(i).Slot).Pic), DDS_Item(Item(Hotbar(i).Slot).Pic))
                        End If
                        Engine_BltToDC DDS_Item(Item(Hotbar(i).Slot).Pic), sRECT, dRECT, frmMain.picHotbar, False
                    End If
                End If
            Case 2 ' spell
                With sRECT
                    .Top = 0
                    .Left = 0
                    .Bottom = 32
                    .Right = 32
                End With
                If Len(Spell(Hotbar(i).Slot).Name) > 0 Then
                    If Spell(Hotbar(i).Slot).Icon > 0 Then
                        If DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon) Is Nothing Then
                            Call InitDDSurf("Spellicons\" & Spell(Hotbar(i).Slot).Icon, DDSD_SpellIcon(Spell(Hotbar(i).Slot).Icon), DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon))
                        End If
                        ' check for cooldown
                        For N = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(N) = Hotbar(i).Slot Then
                                ' has spell
                                If Not SpellCD(N) = 0 Then
                                    sRECT.Left = 32
                                    sRECT.Right = 64
                                End If
                            End If
                        Next
                        Engine_BltToDC DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon), sRECT, dRECT, frmMain.picHotbar, False
                    End If
                End If
        End Select
        
        ' render the letters
        num = "F" & Str(i)
        DrawText frmMain.picHotbar.hDC, dRECT.Left + 2, dRECT.Top + 16, num, QBColor(White)
    Next
    frmMain.picHotbar.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHotbar", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltPlayer(ByVal index As Long)
Dim Anim As Byte, i As Long, X As Long, Y As Long
Dim sprite As Long, spritetop As Long
Dim rec As DxVBLib.RECT
Dim attackspeed As Long
    
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    sprite = GetPlayerSprite(index)

    If sprite < 1 Or sprite > NumCharacters Then Exit Sub
    
    CharacterTimer(sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(sprite) Is Nothing Then
        Call InitDDSurf("characters\" & sprite, DDSD_Character(sprite), DDS_Character(sprite))
    End If

    ' speed from weapon
    If GetPlayerEquipment(index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(index, Weapon)).Speed
    Else
        attackspeed = 1000
    End If

    ' Reset frame
    If Player(index).step = 3 Then
        Anim = 0
    ElseIf Player(index).step = 1 Then
        Anim = 2
    End If
    
    ' Check for attacking animation
    If Player(index).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If Player(index).Attacking = 1 Then
            Anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case GetPlayerDir(index)
            Case DIR_UP
                If (Player(index).YOffset > 8) Then Anim = Player(index).step
            Case DIR_DOWN
                If (Player(index).YOffset < -8) Then Anim = Player(index).step
            Case DIR_LEFT
                If (Player(index).XOffset > 8) Then Anim = Player(index).step
            Case DIR_RIGHT
                If (Player(index).XOffset < -8) Then Anim = Player(index).step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With Player(index)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case GetPlayerDir(index)
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = spritetop * (DDSD_Character(sprite).lHeight / 4)
        .Bottom = .Top + (DDSD_Character(sprite).lHeight / 4)
        .Left = Anim * (DDSD_Character(sprite).lWidth / 4)
        .Right = .Left + (DDSD_Character(sprite).lWidth / 4)
    End With

    ' Calculate the X
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset - ((DDSD_Character(sprite).lWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (DDSD_Character(sprite).lHeight) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - ((DDSD_Character(sprite).lHeight / 4) - 32)
    Else
        ' Proceed as normal
        Y = GetPlayerY(index) * PIC_Y + Player(index).YOffset
    End If
    'If index = 2 Then Debug.Print X & " - " & Y
    ' render the actual sprite
    Call BltSprite(sprite, X, Y, rec)
    
    If Not Player(index).MovementSprite Then
    ' check for paperdolling
        For i = 1 To UBound(PaperdollOrder)
            If GetPlayerEquipment(index, PaperdollOrder(i)) > 0 Then
                If Item(GetPlayerEquipment(index, PaperdollOrder(i))).Paperdoll > 0 Then
                    Call BltPaperdoll(X, Y, Item(GetPlayerEquipment(index, PaperdollOrder(i))).Paperdoll, Anim, spritetop)
                End If
            End If
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPlayer", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltNpc(ByVal mapnpcnum As Long)
Dim Anim As Byte, i As Long, X As Long, Y As Long, sprite As Long, spritetop As Long
Dim rec As DxVBLib.RECT
Dim attackspeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MapNpc(mapnpcnum).num = 0 Then Exit Sub ' no npc set
    
    sprite = NPC(MapNpc(mapnpcnum).num).sprite

    If sprite < 1 Or sprite > NumCharacters Then Exit Sub
    
    CharacterTimer(sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(sprite) Is Nothing Then
        Call InitDDSurf("characters\" & sprite, DDSD_Character(sprite), DDS_Character(sprite))
    End If

    attackspeed = 1000

    ' Reset frame
    Anim = 0
    ' Check for attacking animation
    If MapNpc(mapnpcnum).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If MapNpc(mapnpcnum).Attacking = 1 Then
            Anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(mapnpcnum).dir
            Case DIR_UP
                If (MapNpc(mapnpcnum).YOffset > 8) Then Anim = MapNpc(mapnpcnum).step
            Case DIR_DOWN
                If (MapNpc(mapnpcnum).YOffset < -8) Then Anim = MapNpc(mapnpcnum).step
            Case DIR_LEFT
                If (MapNpc(mapnpcnum).XOffset > 8) Then Anim = MapNpc(mapnpcnum).step
            Case DIR_RIGHT
                If (MapNpc(mapnpcnum).XOffset < -8) Then Anim = MapNpc(mapnpcnum).step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(mapnpcnum)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(mapnpcnum).dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = (DDSD_Character(sprite).lHeight / 4) * spritetop
        .Bottom = .Top + DDSD_Character(sprite).lHeight / 4
        .Left = Anim * (DDSD_Character(sprite).lWidth / 4)
        .Right = .Left + (DDSD_Character(sprite).lWidth / 4)
    End With

    ' Calculate the X
    X = MapNpc(mapnpcnum).X * PIC_X + MapNpc(mapnpcnum).XOffset - ((DDSD_Character(sprite).lWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (DDSD_Character(sprite).lHeight / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(mapnpcnum).Y * PIC_Y + MapNpc(mapnpcnum).YOffset - ((DDSD_Character(sprite).lHeight / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNpc(mapnpcnum).Y * PIC_Y + MapNpc(mapnpcnum).YOffset
    End If

    Call BltSprite(sprite, X, Y, rec)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltNpc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltPaperdoll(ByVal x2 As Long, ByVal y2 As Long, ByVal sprite As Long, ByVal Anim As Long, ByVal spritetop As Long)
Dim rec As DxVBLib.RECT
Dim X As Long, Y As Long
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If sprite < 1 Or sprite > NumPaperdolls Then Exit Sub
    
    If DDS_Paperdoll(sprite) Is Nothing Then
        Call InitDDSurf("Paperdolls\" & sprite, DDSD_Paperdoll(sprite), DDS_Paperdoll(sprite))
    End If
    
    With rec
        .Top = spritetop * (DDSD_Paperdoll(sprite).lHeight / 4)
        .Bottom = .Top + (DDSD_Paperdoll(sprite).lHeight / 4)
        .Left = Anim * (DDSD_Paperdoll(sprite).lWidth / 4)
        .Right = .Left + (DDSD_Paperdoll(sprite).lWidth / 4)
    End With
    
    ' clipping
    X = ConvertMapX(x2)
    Y = ConvertMapY(y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' Clip to screen
    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(X, Y, DDS_Paperdoll(sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPaperdoll", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltSprite(ByVal sprite As Long, ByVal x2 As Long, y2 As Long, rec As DxVBLib.RECT)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If sprite < 1 Or sprite > NumCharacters Then Exit Sub
    X = ConvertMapX(x2)
    Y = ConvertMapY(y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' clipping
    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(X, Y, DDS_Character(sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltAnimatedInvItems()
Dim i As Long
Dim ItemNum As Long, itempic As Long
Dim X As Long, Y As Long
Dim Maxframes As Byte
Dim amount As Long
Dim rec As RECT, rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' check for map animation changes#
    For i = 1 To MAX_MAP_ITEMS

        If MapItem(i).num > 0 Then
            itempic = Item(MapItem(i).num).Pic

            If itempic < 1 Or itempic > NumItems Then Exit Sub
            Maxframes = (DDSD_Item(itempic).lWidth / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

            If MapItem(i).Frame < Maxframes - 1 Then
                MapItem(i).Frame = MapItem(i).Frame + 1
            Else
                MapItem(i).Frame = 1
            End If
        End If

    Next

    For i = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, i)

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            itempic = Item(ItemNum).Pic

            If itempic > 0 And itempic <= NumItems Then
                If DDSD_Item(itempic).lWidth > 64 Then
                    Maxframes = (DDSD_Item(itempic).lWidth / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

                    If InvItemFrame(i) < Maxframes - 1 Then
                        InvItemFrame(i) = InvItemFrame(i) + 1
                    Else
                        InvItemFrame(i) = 1
                    End If

                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = (DDSD_Item(itempic).lWidth / 2) + (InvItemFrame(i) * 32) ' middle to get the start of inv gfx, then +32 for each frame
                        .Right = .Left + 32
                    End With

                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' Load item if not loaded, and reset timer
                    ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_Item(itempic) Is Nothing Then
                        Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                    End If

                    ' We'll now re-blt the item, and place the currency value over it again :P
                    Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picInventory, False

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = rec_pos.Top + 22
                        X = rec_pos.Left - 4
                        amount = CStr(GetPlayerInvItemValue(MyIndex, i))
                        ' Draw currency but with k, m, b etc. using a convertion function
                        DrawText frmMain.picInventory.hDC, X, Y, ConvertCurrency(amount), QBColor(Yellow)

                        ' Check if it's gold, and update the label
                        If GetPlayerInvItemNum(MyIndex, i) = 1 Then '1 = gold :P
                            BltGold (amount)
                        End If
                    End If
                End If
            End If
        End If

    Next

    frmMain.picInventory.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltAnimatedInvItems", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltFace()
Dim rec As RECT, rec_pos As RECT, faceNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumFaces = 0 Then Exit Sub
    
    frmMain.picFace.Cls
    
    faceNum = Player(MyIndex).sprite
    
    If faceNum <= 0 Or faceNum > NumFaces Then Exit Sub

    With rec
        .Top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    With rec_pos
        .Top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    ' Load face if not loaded, and reset timer
    FaceTimer(faceNum) = GetTickCount + SurfaceTimerMax

    If DDS_Face(faceNum) Is Nothing Then
        Call InitDDSurf("Faces\" & faceNum, DDSD_Face(faceNum), DDS_Face(faceNum))
    End If

    Engine_BltToDC DDS_Face(faceNum), rec, rec_pos, frmMain.picFace, False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltFace", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltEquipment()
Dim i As Long, ItemNum As Long, itempic As Long
Dim rec As RECT, rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumItems = 0 Then Exit Sub
    
    frmMain.picCharacter.Cls

    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(MyIndex, i)

        If ItemNum > 0 Then
            itempic = Item(ItemNum).Pic

            With rec
                .Top = 0
                .Bottom = 32
                .Left = 32
                .Right = 64
            End With

            With rec_pos
                .Top = EqTop
                .Bottom = .Top + PIC_Y
                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            ' Load item if not loaded, and reset timer
            ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

            If DDS_Item(itempic) Is Nothing Then
                Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
            End If

            Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picCharacter, False
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltEquipment", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltInventory()
Dim i As Long, X As Long, Y As Long, ItemNum As Long, itempic As Long
Dim amount As Long
Dim rec As RECT, rec_pos As RECT
Dim colour As Long
Dim tmpItem As Long, amountModifier As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' reset gold label
    frmMain.lblGold.Caption = "0 " & GetTranslation("Rupias")
    frmMain.lblGold.ForeColor = QBColor(White)
    
    frmMain.picInventory.Cls

    For i = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, i)

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            itempic = Item(ItemNum).Pic
            
            amountModifier = 0
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(X).num)
                    If TradeYourOffer(X).num = i Then
                        ' check if currency
                        If Not isItemStackable(tmpItem) Then
                            ' normal item, exit out
                            GoTo NextLoop
                        Else
                            ' if amount = all currency, remove from inventory
                            If TradeYourOffer(X).value = GetPlayerInvItemValue(MyIndex, i) Then
                                GoTo NextLoop
                            Else
                                ' not all, change modifier to show change in currency count
                                amountModifier = TradeYourOffer(X).value
                            End If
                        End If
                    End If
                Next
            End If

            If itempic > 0 And itempic <= NumItems Then
                If DDSD_Item(itempic).lWidth <= 64 Then ' more than 1 frame is handled by anim sub

                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With

                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' Load item if not loaded, and reset timer
                    ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_Item(itempic) Is Nothing Then
                        Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                    End If

                    Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picInventory, False

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = rec_pos.Top + 22
                        X = rec_pos.Left - 4
                        
                        amount = GetPlayerInvItemValue(MyIndex, i) - amountModifier
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If amount < 1000000 Then
                            colour = QBColor(White)
                        ElseIf amount > 1000000 And amount < 10000000 Then
                            colour = QBColor(Yellow)
                        ElseIf amount > 10000000 Then
                            colour = QBColor(BrightGreen)
                        End If
                        
                        
                        DrawText frmMain.picInventory.hDC, X, Y, Format$(ConvertCurrency(Str(amount)), "#,###,###,###"), colour

                        ' Check if it's gold, and update the label
                        If GetPlayerInvItemNum(MyIndex, i) = 1 Then '1 = gold :P
                            BltGold (amount)
                        End If
                    End If
                End If
            End If
        End If
NextLoop:
    Next
    
    frmMain.picInventory.Refresh
    'update animated items
    BltAnimatedInvItems
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventory", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltTrade()
Dim i As Long, X As Long, Y As Long, ItemNum As Long, itempic As Long
Dim amount As Long
Dim rec As RECT, rec_pos As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    frmMain.picYourTrade.Cls
    frmMain.picTheirTrade.Cls
    
    For i = 1 To MAX_INV
        ' blt your own offer
        ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            itempic = Item(ItemNum).Pic

            If itempic > 0 And itempic <= NumItems Then
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With

                With rec_pos
                    .Top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load item if not loaded, and reset timer
                ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                If DDS_Item(itempic) Is Nothing Then
                    Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                End If

                Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picYourTrade, False

                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).value > 1 Then
                    Y = rec_pos.Top + 22
                    X = rec_pos.Left - 4
                    
                    amount = TradeYourOffer(i).value
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If amount < 1000000 Then
                        colour = QBColor(White)
                    ElseIf amount > 1000000 And amount < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf amount > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    
                    DrawText frmMain.picYourTrade.hDC, X, Y, ConvertCurrency(Str(amount)), colour
                End If
            End If
        End If
            
        ' blt their offer
        ItemNum = TradeTheirOffer(i).num

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            itempic = Item(ItemNum).Pic

            If itempic > 0 And itempic <= NumItems Then
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With

                With rec_pos
                    .Top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load item if not loaded, and reset timer
                ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                If DDS_Item(itempic) Is Nothing Then
                    Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                End If

                Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picTheirTrade, False

                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).value > 1 Then
                    Y = rec_pos.Top + 22
                    X = rec_pos.Left - 4
                    
                    amount = TradeTheirOffer(i).value
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If amount < 1000000 Then
                        colour = QBColor(White)
                    ElseIf amount > 1000000 And amount < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf amount > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    
                    DrawText frmMain.picTheirTrade.hDC, X, Y, ConvertCurrency(Str(amount)), colour
                End If
            End If
        End If
    Next
    
    frmMain.picYourTrade.Refresh
    frmMain.picTheirTrade.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTrade", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltPlayerSpells()
Dim i As Long, X As Long, Y As Long, spellnum As Long, spellicon As Long
Dim amount As String
Dim rec As RECT, rec_pos As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    frmMain.picSpells.Cls

    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i)

        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellicon = Spell(spellnum).Icon

            If spellicon > 0 And spellicon <= NumSpellIcons Then
            
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 0
                    .Right = 32
                End With
                
                If Not SpellCD(i) = 0 Then
                    rec.Left = 32
                    rec.Right = 64
                End If

                With rec_pos
                    .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load spellicon if not loaded, and reset timer
                SpellIconTimer(spellicon) = GetTickCount + SurfaceTimerMax

                If DDS_SpellIcon(spellicon) Is Nothing Then
                    Call InitDDSurf("SpellIcons\" & spellicon, DDSD_SpellIcon(spellicon), DDS_SpellIcon(spellicon))
                End If

                Engine_BltToDC DDS_SpellIcon(spellicon), rec, rec_pos, frmMain.picSpells, False
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPlayerSpells", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltShop()
Dim i As Long, X As Long, Y As Long, ItemNum As Long, itempic As Long
Dim amount As String
Dim rec As RECT, rec_pos As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    frmMain.picShopItems.Cls

    For i = 1 To MAX_TRADES
        ItemNum = Shop(InShop).TradeItem(i).Item 'GetPlayerInvItemNum(MyIndex, i)
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            itempic = Item(ItemNum).Pic
            If itempic > 0 And itempic <= NumItems Then
            
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With
                
                With rec_pos
                    .Top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                    .Right = .Left + PIC_X
                End With
                
                ' Load item if not loaded, and reset timer
                ItemTimer(itempic) = GetTickCount + SurfaceTimerMax
                
                If DDS_Item(itempic) Is Nothing Then
                    Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                End If
                
                Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picShopItems, False
                
                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                    Y = rec_pos.Top + 22
                    X = rec_pos.Left - 4
                    amount = CStr(Shop(InShop).TradeItem(i).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(amount) < 1000000 Then
                        colour = QBColor(White)
                    ElseIf CLng(amount) > 1000000 And CLng(amount) < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf CLng(amount) > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    
                    DrawText frmMain.picShopItems.hDC, X, Y, ConvertCurrency(amount), colour
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltShop", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltInventoryItem(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT, rec_pos As RECT
Dim ItemNum As Long, itempic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)

    If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
        itempic = Item(ItemNum).Pic
        
        If itempic = 0 Then Exit Sub

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = DDSD_Item(itempic).lWidth / 2
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 2
            .Bottom = .Top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        ' Load item if not loaded, and reset timer
        ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

        If DDS_Item(itempic) Is Nothing Then
            Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
        End If

        Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picTempInv, False

        With frmMain.picTempInv
            .Top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventoryItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltDraggedSpell(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT, rec_pos As RECT
Dim spellnum As Long, spellpic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = PlayerSpells(DragSpell)

    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        spellpic = Spell(spellnum).Icon
        
        If spellpic = 0 Then Exit Sub

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 2
            .Bottom = .Top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        ' Load item if not loaded, and reset timer
        SpellIconTimer(spellpic) = GetTickCount + SurfaceTimerMax

        If DDS_SpellIcon(spellpic) Is Nothing Then
            Call InitDDSurf("Spellicons\" & spellpic, DDSD_SpellIcon(spellpic), DDS_SpellIcon(spellpic))
        End If

        Engine_BltToDC DDS_SpellIcon(spellpic), rec, rec_pos, frmMain.picTempSpell, False

        With frmMain.picTempSpell
            .Top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventoryItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltItemDesc(ByVal ItemNum As Long)
Dim rec As RECT, rec_pos As RECT
Dim itempic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmMain.picItemDescPic.Cls
    
    If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
        itempic = Item(ItemNum).Pic

        If itempic = 0 Then Exit Sub
        
        ' Load item if not loaded, and reset timer
        ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

        If DDS_Item(itempic) Is Nothing Then
            Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
        End If

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = DDSD_Item(itempic).lWidth / 2
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picItemDescPic, False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltItemDesc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltSpellDesc(ByVal spellnum As Long)
Dim rec As RECT, rec_pos As RECT
Dim spellpic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmMain.picSpellDescPic.Cls

    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        spellpic = Spell(spellnum).Icon

        If spellpic <= 0 Or spellpic > NumSpellIcons Then Exit Sub
        
        ' Load item if not loaded, and reset timer
        SpellIconTimer(spellpic) = GetTickCount + SurfaceTimerMax

        If DDS_SpellIcon(spellpic) Is Nothing Then
            Call InitDDSurf("SpellIcons\" & spellpic, DDSD_SpellIcon(spellpic), DDS_SpellIcon(spellpic))
        End If

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        Engine_BltToDC DDS_SpellIcon(spellpic), rec, rec_pos, frmMain.picSpellDescPic, False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltSpellDesc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_BltTileset()
Dim Height As Long
Dim Width As Long
Dim Tileset As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.value
    
    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub
    
    ' make sure it's loaded
    If DDS_Tileset(Tileset) Is Nothing Then
        Call InitDDSurf("tilesets\" & Tileset, DDSD_Tileset(Tileset), DDS_Tileset(Tileset))
    End If
    
    Height = DDSD_Tileset(Tileset).lHeight
    Width = DDSD_Tileset(Tileset).lWidth
    
    dRECT.Top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width
    
    frmEditor_Map.picBackSelect.Height = Height
    frmEditor_Map.picBackSelect.Width = Width
    
    Call Engine_BltToDC(DDS_Tileset(Tileset), sRECT, dRECT, frmEditor_Map.picBackSelect)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltTileset", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltTileOutline()
Dim rec As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optBlock.value Then Exit Sub

    With rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    Call Engine_BltFast(ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), DDS_Misc, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTileOutline", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NewCharacterBltSprite()
Dim sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If frmMenu.optMale.value = True Then
        sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite(newCharSprite)
    Else
        sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite(newCharSprite)
    End If
    
    If sprite < 1 Or sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    CharacterTimer(sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(sprite) Is Nothing Then
        Call InitDDSurf("Characters\" & sprite, DDSD_Character(sprite), DDS_Character(sprite))
    End If
    
    Width = DDSD_Character(sprite).lWidth / 4
    Height = DDSD_Character(sprite).lHeight / 4
    
    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height
    
    sRECT.Top = 0
    sRECT.Bottom = sRECT.Top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    
    dRECT.Top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width
    
    Call Engine_BltToDC(DDS_Character(sprite), sRECT, dRECT, frmMenu.picSprite)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NewCharacterBltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_BltMapItem()
Dim ItemNum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = Item(frmEditor_Map.scrlMapItem.value).Pic

    If ItemNum < 1 Or ItemNum > NumItems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    ItemTimer(ItemNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(ItemNum) Is Nothing Then
        Call InitDDSurf("Items\" & ItemNum, DDSD_Item(ItemNum), DDS_Item(ItemNum))
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    Call Engine_BltToDC(DDS_Item(ItemNum), sRECT, dRECT, frmEditor_Map.picMapItem)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltMapItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_BltKey()
Dim ItemNum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = Item(frmEditor_Map.scrlMapKey.value).Pic

    If ItemNum < 1 Or ItemNum > NumItems Then
        frmEditor_Map.picMapKey.Cls
        Exit Sub
    End If

    ItemTimer(ItemNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(ItemNum) Is Nothing Then
        Call InitDDSurf("Items\" & ItemNum, DDSD_Item(ItemNum), DDS_Item(ItemNum))
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    Call Engine_BltToDC(DDS_Item(ItemNum), sRECT, dRECT, frmEditor_Map.picMapKey)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltKey", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_BltItem()
Dim ItemNum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = frmEditor_Item.scrlPic.value

    If ItemNum < 1 Or ItemNum > NumItems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If

    ItemTimer(ItemNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(ItemNum) Is Nothing Then
        Call InitDDSurf("Items\" & ItemNum, DDSD_Item(ItemNum), DDS_Item(ItemNum))
    End If

    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    dRECT = sRECT
    Call Engine_BltToDC(DDS_Item(ItemNum), sRECT, dRECT, frmEditor_Item.picItem)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_BltItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_BltPaperdoll()
Dim sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Item.picPaperdoll.Cls
    
    sprite = frmEditor_Item.scrlPaperdoll.value

    If sprite < 1 Or sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    PaperdollTimer(sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Paperdoll(sprite) Is Nothing Then
        Call InitDDSurf("paperdolls\" & sprite, DDSD_Paperdoll(sprite), DDS_Paperdoll(sprite))
    End If

    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = DDSD_Paperdoll(sprite).lHeight
    sRECT.Left = 0
    sRECT.Right = DDSD_Paperdoll(sprite).lWidth
    ' same for destination as source
    dRECT = sRECT
    
    Call Engine_BltToDC(DDS_Paperdoll(sprite), sRECT, dRECT, frmEditor_Item.picPaperdoll)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_BltPaperdoll", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_BltIcon()
Dim iconnum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    iconnum = frmEditor_Spell.scrlIcon.value
    
    If iconnum < 1 Or iconnum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    SpellIconTimer(iconnum) = GetTickCount + SurfaceTimerMax
    
    If DDS_SpellIcon(iconnum) Is Nothing Then
        Call InitDDSurf("SpellIcons\" & iconnum, DDSD_SpellIcon(iconnum), DDS_SpellIcon(iconnum))
    End If
    
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    
    Call Engine_BltToDC(DDS_SpellIcon(iconnum), sRECT, dRECT, frmEditor_Spell.picSprite)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_BltIcon", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorAnim_BltAnim()
Dim Animationnum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
Dim i As Long
Dim Width As Long, Height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim ShouldRender As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).value
        
        If Animationnum < 1 Or Animationnum > NumAnimations Then
            frmEditor_Animation.picSprite(i).Cls
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = GetTickCount
                ShouldRender = True
            End If
        
            If ShouldRender Then
                frmEditor_Animation.picSprite(i).Cls
            
                AnimationTimer(Animationnum) = GetTickCount + SurfaceTimerMax
                
                If DDS_Animation(Animationnum) Is Nothing Then
                    Call InitDDSurf("animations\" & Animationnum, DDSD_Animation(Animationnum), DDS_Animation(Animationnum))
                End If
                
                If frmEditor_Animation.scrlFrameCount(i).value > 0 Then
                    ' total width divided by frame count
                    Width = DDSD_Animation(Animationnum).lWidth / frmEditor_Animation.scrlFrameCount(i).value
                    Height = DDSD_Animation(Animationnum).lHeight
                    
                    sRECT.Top = 0
                    sRECT.Bottom = Height
                    sRECT.Left = (AnimEditorFrame(i) - 1) * Width
                    sRECT.Right = sRECT.Left + Width
                    
                    dRECT.Top = 0
                    dRECT.Bottom = Height
                    dRECT.Left = 0
                    dRECT.Right = Width
                    
                    Call Engine_BltToDC(DDS_Animation(Animationnum), sRECT, dRECT, frmEditor_Animation.picSprite(i))
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorAnim_BltAnim", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_BltSprite()
Dim sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    sprite = frmEditor_NPC.scrlSprite.value

    If sprite < 1 Or sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    CharacterTimer(sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(sprite) Is Nothing Then
        Call InitDDSurf("characters\" & sprite, DDSD_Character(sprite), DDS_Character(sprite))
    End If

    sRECT.Top = 0
    sRECT.Bottom = SIZE_Y
    sRECT.Left = PIC_X * 3 ' facing down
    sRECT.Right = sRECT.Left + SIZE_X
    dRECT.Top = 0
    dRECT.Bottom = SIZE_Y
    dRECT.Left = 0
    dRECT.Right = SIZE_X
    Call Engine_BltToDC(DDS_Character(sprite), sRECT, dRECT, frmEditor_NPC.picSprite)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorResource_BltSprite()
Dim sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' normal sprite
    sprite = frmEditor_Resource.scrlNormalPic.value

    If sprite < 1 Or sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        ResourceTimer(sprite) = GetTickCount + SurfaceTimerMax
        If DDS_Resource(sprite) Is Nothing Then
            Call InitDDSurf("Resources\" & sprite, DDSD_Resource(sprite), DDS_Resource(sprite))
        End If
        sRECT.Top = 0
        sRECT.Bottom = DDSD_Resource(sprite).lHeight
        sRECT.Left = 0
        sRECT.Right = DDSD_Resource(sprite).lWidth
        dRECT.Top = 0
        dRECT.Bottom = DDSD_Resource(sprite).lHeight
        dRECT.Left = 0
        dRECT.Right = DDSD_Resource(sprite).lWidth
        Call Engine_BltToDC(DDS_Resource(sprite), sRECT, dRECT, frmEditor_Resource.picNormalPic)
    End If

    ' exhausted sprite
    sprite = frmEditor_Resource.scrlExhaustedPic.value

    If sprite < 1 Or sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        ResourceTimer(sprite) = GetTickCount + SurfaceTimerMax
        If DDS_Resource(sprite) Is Nothing Then
            Call InitDDSurf("Resources\" & sprite, DDSD_Resource(sprite), DDS_Resource(sprite))
        End If
        sRECT.Top = 0
        sRECT.Bottom = DDSD_Resource(sprite).lHeight
        sRECT.Left = 0
        sRECT.Right = DDSD_Resource(sprite).lWidth
        dRECT.Top = 0
        dRECT.Bottom = DDSD_Resource(sprite).lHeight
        dRECT.Left = 0
        dRECT.Right = DDSD_Resource(sprite).lWidth
        Call Engine_BltToDC(DDS_Resource(sprite), sRECT, dRECT, frmEditor_Resource.picExhaustedPic)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorResource_BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Render_Graphics()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim amount As Long
Dim rec As DxVBLib.RECT
Dim rec_pos As DxVBLib.RECT

    ' check if player is loading the map
    If TempPlayer(MyIndex).IsLoading = True Then Exit Sub
    
    If InGame = False Then Exit Sub

    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check if automation is screwed
    If Not CheckSurfaces Then
        ' exit out and let them know we need to re-init
        ReInitSurfaces = True
        Exit Sub
    Else
        ' if we need to fix the surfaces then do so
        If ReInitSurfaces Then
            ReInitSurfaces = False
            ReInitDD
        End If
    End If
    
    ' don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    If GettingMap Then Exit Sub

    ' update the viewpoint
    UpdateCamera
    
    
    ' update animation editor
    If Editor = EDITOR_ANIMATION Then
        EditorAnim_BltAnim
    End If
    
    ' fill it with black
    DDS_BackBuffer.BltColorFill rec_pos, 0
    
    ' blit lower tiles
    If NumTileSets > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    Call BltMapTile(X, Y)
                End If
            Next
        Next
    End If
    
   
    
    ' render the decals
    For i = 1 To MAX_BYTE
        Call BltBlood(i)
    Next

    ' Blit out the items
    If NumItems > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                Call BltItem(i)
            End If
        Next
    End If
    
    ' draw animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                BltAnimation i, 0
            End If
        Next
    End If
    
    ' blt projec tiles for each player
    For i = 1 To Player_HighIndex
        For X = 1 To MAX_PLAYER_PROJECTILES
            If Player(i).ProjecTile(X).Pic > 0 Then
                BltProjectile i, X
            End If
        Next
    Next

    

    'MASK RESOURCE
    'Y based Render
    ' double loop in order to prevent dis-tilage
    ' Resources
    Call bltMaskResource

    
    For Y = TileView.Top To TileView.Bottom
        If NumCharacters > 0 Then
            ' Players
            For i = 1 To Player_HighIndex
            If Not TempPlayer(i).IsLoading = True Then
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).Y = Y And (GetPlayerVisible(i) = False Or i = MyIndex) Then
                        Call ComputeBltPlayer(i)
                    End If
                End If
            End If
            Next
        
            ' Npcs
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Y = Y Then
                    Call BltNpc(i)
                End If
            Next
        End If

    Next
    
    
    
    
    'FRINGE RESOURCE
    'Y based Render
    ' double loop in order to prevent dis-tilage
    ' Resources
    If TempPlayer(MyIndex).IsLoading = False Then Call bltFringeResource
    
    ' animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                BltAnimation i, 1
            End If
        Next
    End If

    ' blit out upper tiles
    If NumTileSets > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    Call BltMapFringeTile(X, Y)
                End If
            Next
        Next
    End If
    
    ' blit out a square at mouse cursor
    If InMapEditor Then
        If frmEditor_Map.optBlock.value = True Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.Top To TileView.Bottom
                    If IsValidMapPoint(X, Y) Then
                        Call BltDirection(X, Y)
                    End If
                Next
            Next
        End If
        Call BltTileOutline
    End If
    ' minimap
    If Options.MiniMap = 1 Then BltMiniMap
    DoEvents
    
    ' Blt the target icon
    If myTarget > 0 Then
    

    If myTargetType = TARGET_TYPE_PLAYER And GetPlayerVisible(i) = False Then
                 BltTarget (Player(myTarget).X * 32) + Player(myTarget).XOffset, (Player(myTarget).Y * 32) + Player(myTarget).YOffset
         ElseIf myTargetType = TARGET_TYPE_NPC Then
                 BltTarget (MapNpc(myTarget).X * 32) + MapNpc(myTarget).XOffset, (MapNpc(myTarget).Y * 32) + MapNpc(myTarget).YOffset
         End If
    End If
    
    ' blt the hover icon
For i = 1 To Player_HighIndex
         If IsPlaying(i) And GetPlayerVisible(i) = False Then
                 If InGame = False Then Exit Sub
                 If Player(i).map = Player(MyIndex).map Then
                    
                         If CurX = Player(i).X And CurY = Player(i).Y Then
                                 If myTargetType = TARGET_TYPE_PLAYER And myTarget = i Then
                                         ' dont render lol
                                 Else
                                         BltHover TARGET_TYPE_PLAYER, i, (Player(i).X * 32) + Player(i).XOffset, (Player(i).Y * 32) + Player(i).YOffset
                                 End If
                         End If
                 End If
         End If
Next
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If CurX = MapNpc(i).X And CurY = MapNpc(i).Y Then
                If myTargetType = TARGET_TYPE_NPC And myTarget = i Then
                    ' dont render lol
                Else
                    BltHover TARGET_TYPE_NPC, i, (MapNpc(i).X * 32) + MapNpc(i).XOffset, (MapNpc(i).Y * 32) + MapNpc(i).YOffset
                End If
            End If
        End If
    Next

' if map has weather
' also if map has Sandstorm, ignore Rainon
If map.Weather > 0 And InMapEditor = False And (Rainon = True Or map.Weather = 3) Then

     ' rain drops
     Dim G As Long, RainRect As DxVBLib.RECT, RainMax As Long
    
     ' check if we've loaded the raindrop
     If DDS_Rain Is Nothing Then
         Call InitDDSurf("rain.bmp", DDSD_Rain, DDS_Rain)
     ElseIf DDS_Snow Is Nothing Then
         Call InitDDSurf("snow.bmp", DDSD_Snow, DDS_Snow)
     ElseIf DDS_Sandstorm Is Nothing Then
         Call InitDDSurf("Sandstorm.bmp", DDSD_Sandstorm, DDS_Sandstorm)
     End If

     ' setup raindrop rec
     With rec
         .Top = 0
         .Bottom = 2
         .Left = 0
         .Right = 14
     End With
    
     ' less drops for snow
     If map.Weather = 1 Then RainMax = 25
     If map.Weather = 2 Then RainMax = 15
     If map.Weather = 3 Then RainMax = 100
    
     ' loop through all raindrops
     For G = 1 To RainMax
         If RainDrop(G).InMotion = 1 Then
             ' render the rain
             If map.Weather = 1 Then ' rain
                Call Engine_BltFast(Camera.Left + RainDrop(G).X, Camera.Top + RainDrop(G).Y, DDS_Rain, RainRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
             ElseIf map.Weather = 2 Then ' snow
                Call Engine_BltFast(Camera.Left + RainDrop(G).X, Camera.Top + RainDrop(G).Y, DDS_Snow, RainRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
             ElseIf map.Weather = 3 Then ' sandstorm
                Call Engine_BltFast(Camera.Left + RainDrop(G).X, Camera.Top + RainDrop(G).Y, DDS_Sandstorm, RainRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
             End If
         End If
     Next
End If
    ' Lock the backbuffer so we can draw text and names
    TexthDC = DDS_BackBuffer.GetDC
    
    ' draw chat to screen / transparency chat
    If Options.ChatToScreen = 2 Then
        DrawChat
    End If
    
    ' draw FPS
    If BFPS Then
        Call DrawText(TexthDC, Camera.Right - (Len("FPS: " & GameFPS) * 9), Camera.Top + 30, ("FPS: " & GameFPS), QBColor(Yellow))
    End If

    ' draw cursor, player X and Y locations
    If BLoc Then
        Call DrawText(TexthDC, Camera.Left, Camera.Top + 60, ("cur x: " & CurX & " y: " & CurY), QBColor(Yellow))
        Call DrawText(TexthDC, Camera.Left, Camera.Top + 75, ("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), QBColor(Yellow))
        Call DrawText(TexthDC, Camera.Left, Camera.Top + 90, (" (map #" & GetPlayerMap(MyIndex) & ")"), QBColor(Yellow))
    End If

    
    ' draw player names & level
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) And (Not GetPlayerVisible(i) = 1 Or i = MyIndex) Then
            If Options.Names = 1 Then
                Call DrawPlayerName(i)
            End If
            If Options.Level = 1 Then
                Call DrawPlayerLevel(i)
            End If
        End If
    Next
    
    ' draw npc names & level
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If Options.Names = 1 Then
                Call DrawNpcName(i)
            End If
            If Options.Level = 1 Then
                Call DrawNpcLevel(i)
            End If
        End If
    Next
    
    For i = 1 To Action_HighIndex
        Call BltActionMsg(i)
    Next i

    ' Blit out map attributes
    If InMapEditor Then
        Call BltMapAttributes
    End If

    ' Draw map name
    Call DrawText(TexthDC, DrawMapNameX, DrawMapNameY, (map.TranslatedName), DrawMapNameColor)
    
    If Not Options.MappingMode = 1 Then
    
        'Draw rupees
        If HasMaxGold Then
                Call DrawText(TexthDC, Camera.Left + 700, Camera.Top + 580, (frmMain.lblGold.Caption), QBColor(BrightGreen))
            Else
                Call DrawText(TexthDC, Camera.Left + 700, Camera.Top + 580, (frmMain.lblGold.Caption), QBColor(White))
        End If
        
        'show secure mode
        If Options.SafeMode = 1 Then
            'Call DrawText(TexthDC, Camera.Right - 305, Camera.Top + 35, Trim$("Modo Seguro Activado"), QBColor(BrightGreen))
        Else
            Call DrawText(TexthDC, Camera.Right - 305, Camera.Top + 35, GetTranslation("Modo Seguro Desactivado"), QBColor(BrightRed))
        End If
    End If
    
    ' Release DC
    DDS_BackBuffer.ReleaseDC TexthDC
    
    ' draw the messages at the very top!
    For i = 1 To MAX_BYTE
        If chatBubble(i).active Then
            DrawChatBubble i
        End If
    Next
    
        'blt health, magic and experience bar
    If Not Options.MappingMode = 1 Then
        'blt rupee
        Call BltRupee(i)
        'blt empty hearts
        Call BltHearts(i)
        'blt magic bar
        Call BltMagicBar(i)
        'blt red hearts
        Call BltHealth(i)
        'render the bars
        BltBars
    End If

    ' Get rec
    With rec
        .Top = Camera.Top
        .Bottom = .Top + ScreenY
        .Left = Camera.Left
        .Right = .Left + ScreenX
    End With
    
    ' rec_pos
    With rec_pos
        .Bottom = ((MAX_MAPY + 1) * PIC_Y)
        .Right = ((MAX_MAPX + 1) * PIC_X)
    End With
    
    ' Flip and render
    DX7.GetWindowRect frmMain.picScreen.hwnd, rec_pos
    DDS_Primary.Blt rec_pos, DDS_BackBuffer, rec, DDBLT_WAIT
    
    ' Error handler
    Exit Sub
    
errorhandler:
    HandleError "Render_Graphics", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateCamera()
Dim offsetX As Long
Dim offsetY As Long
Dim StartX As Long
Dim StartY As Long
Dim EndX As Long
Dim EndY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    offsetX = Player(MyIndex).XOffset + PIC_X
    offsetY = Player(MyIndex).YOffset + PIC_Y

    StartX = GetPlayerX(MyIndex) - StartXValue
    StartY = GetPlayerY(MyIndex) - StartYValue
    If StartX < 0 Then
        offsetX = 0
        If StartX = -1 Then
            If Player(MyIndex).XOffset > 0 Then
                offsetX = Player(MyIndex).XOffset
            End If
        End If
        StartX = 0
    End If
    If StartY < 0 Then
        offsetY = 0
        If StartY = -1 Then
            If Player(MyIndex).YOffset > 0 Then
                offsetY = Player(MyIndex).YOffset
            End If
        End If
        StartY = 0
    End If
    
    EndX = StartX + EndXValue
    EndY = StartY + EndYValue
    If EndX > map.MaxX Then
        offsetX = 32
        If EndX = map.MaxX + 1 Then
            If Player(MyIndex).XOffset < 0 Then
                offsetX = Player(MyIndex).XOffset + PIC_X
            End If
        End If
        EndX = map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If
    If EndY > map.MaxY Then
        offsetY = 32
        If EndY = map.MaxY + 1 Then
            If Player(MyIndex).YOffset < 0 Then
                offsetY = Player(MyIndex).YOffset + PIC_Y
            End If
        End If
        EndY = map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = offsetY
        .Bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With
    
    UpdateDrawMapName

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateCamera", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapX = X - (TileView.Left * PIC_X)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapX", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapY = Y - (TileView.Top * PIC_Y)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapY", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InViewPort = False

    If X < TileView.Left Then Exit Function
    If Y < TileView.Top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.Bottom Then Exit Function
    InViewPort = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InViewPort", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > map.MaxX Then Exit Function
    If Y > map.MaxY Then Exit Function
    IsValidMapPoint = True
        
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsValidMapPoint", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub LoadTilesets()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim tilesetInUse() As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim tilesetInUse(0 To NumTileSets)
    
    For X = 0 To map.MaxX
        For Y = 0 To map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                ' check exists
                If map.Tile(X, Y).layer(i).Tileset > 0 And map.Tile(X, Y).layer(i).Tileset <= NumTileSets Then
                    tilesetInUse(map.Tile(X, Y).layer(i).Tileset) = True
                End If
            Next
        Next
    Next
    
    For i = 1 To NumTileSets
        If tilesetInUse(i) Then
            ' load tileset
            If DDS_Tileset(i) Is Nothing Then
                Call InitDDSurf("tilesets\" & i, DDSD_Tileset(i), DDS_Tileset(i))
            End If
        Else
            ' unload tileset
            Call ZeroMemory(ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i)))
            Set DDS_Tileset(i) = Nothing
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTilesets", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltBank()
Dim i As Long, X As Long, Y As Long, ItemNum As Long
Dim amount As String
Dim sRECT As RECT, dRECT As RECT
Dim sprite As Long, colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMain.picBank.Visible = True Then
        frmMain.picBank.Cls
                
        For i = 1 To MAX_BANK
            ItemNum = GetBankItemNum(i)
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            
                sprite = Item(ItemNum).Pic
                
                If sprite <= 0 Or sprite > NumItems Then Exit Sub
                
                If DDS_Item(sprite) Is Nothing Then
                    Call InitDDSurf("Items\" & sprite, DDSD_Item(sprite), DDS_Item(sprite))
                End If
            
                With sRECT
                    .Top = 0
                    .Bottom = .Top + PIC_Y
                    .Left = DDSD_Item(sprite).lWidth / 2
                    .Right = .Left + PIC_X
                End With
                
                With dRECT
                    .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                    .Right = .Left + PIC_X
                End With
                
                Engine_BltToDC DDS_Item(sprite), sRECT, dRECT, frmMain.picBank, False

                ' If item is a stack - draw the amount you have
                If GetBankItemValue(0, i) > 1 Then
                    Y = dRECT.Top + 22
                    X = dRECT.Left - 4
                
                    amount = CStr(GetBankItemValue(0, i))
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(amount) < 1000000 Then
                        colour = QBColor(White)
                    ElseIf CLng(amount) > 1000000 And CLng(amount) < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf CLng(amount) > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    DrawText frmMain.picBank.hDC, X, Y, ConvertCurrency(amount), colour
                End If
            End If
        Next
    
        frmMain.picBank.Refresh
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBank", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltBankItem(ByVal X As Long, ByVal Y As Long)
Dim sRECT As RECT, dRECT As RECT
Dim ItemNum As Long
Dim sprite As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = GetBankItemNum(DragBankSlotNum)
    sprite = Item(GetBankItemNum(DragBankSlotNum)).Pic
    
    If DDS_Item(sprite) Is Nothing Then
        Call InitDDSurf("Items\" & sprite, DDSD_Item(sprite), DDS_Item(sprite))
    End If
    
    If ItemNum > 0 Then
        If ItemNum <= MAX_ITEMS Then
            With sRECT
                .Top = 0
                .Bottom = .Top + PIC_Y
                .Left = DDSD_Item(sprite).lWidth / 2
                .Right = .Left + PIC_X
            End With
        End If
    End If
    
    With dRECT
        .Top = 2
        .Bottom = .Top + PIC_Y
        .Left = 2
        .Right = .Left + PIC_X
    End With

    Engine_BltToDC DDS_Item(sprite), sRECT, dRECT, frmMain.picTempBank
    
    With frmMain.picTempBank
        .Top = Y
        .Left = X
        .Visible = True
        .ZOrder (0)
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBankItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' player Projectiles
Public Sub BltProjectile(ByVal index As Long, ByVal PlayerProjectile As Long)
Dim X As Long, Y As Long, PicNum As Long, i As Long
Dim rec As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check for subscript error
    If index < 1 Or PlayerProjectile < 1 Or PlayerProjectile > MAX_PLAYER_PROJECTILES Then Exit Sub
    
    ' check to see if it's time to move the Projectile
    If GetTickCount > Player(index).ProjecTile(PlayerProjectile).TravelTime Then
        With Player(index).ProjecTile(PlayerProjectile)
            ' set next travel time and the current position and then set the actual direction based on RMXP arrow tiles.
            Select Case .direction
                ' down
                Case 0
                    .Y = .Y + 1
                    ' check if they reached maxrange
                    If .Y = (GetPlayerY(index) + .range) + 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
                ' up
                Case 1
                    .Y = .Y - 1
                    ' check if they reached maxrange
                    If .Y = (GetPlayerY(index) - .range) - 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
                ' right
                Case 2
                    .X = .X + 1
                    ' check if they reached max range
                    If .X = (GetPlayerX(index) + .range) + 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
                ' left
                Case 3
                    .X = .X - 1
                    ' check if they reached maxrange
                    If .X = (GetPlayerX(index) - .range) - 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
            End Select
            .TravelTime = GetTickCount + .Speed
        End With
    End If
    
    ' set the x, y & pic values for future reference
    X = Player(index).ProjecTile(PlayerProjectile).X
    Y = Player(index).ProjecTile(PlayerProjectile).Y
    PicNum = Player(index).ProjecTile(PlayerProjectile).Pic
    
    ' check if left map
    If X > map.MaxX Or Y > map.MaxY Or X < 0 Or Y < 0 Then
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
    ' check if we hit a block
    If map.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
    ' check if we hit an item
    If map.Tile(X, Y).Type = TILE_TYPE_ITEM Then
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
    ' check if we hit a resource
    If map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
    ' check for player hit
    For i = 1 To Player_HighIndex
        If X = GetPlayerX(i) And Y = GetPlayerY(i) Then
            ' they're hit, remove it
            If Not X = Player(MyIndex).X Or Not Y = GetPlayerY(MyIndex) Then
                ClearProjectile index, PlayerProjectile
                Exit Sub
            End If
        End If
    Next
    
    ' check for npc hit
    For i = 1 To MAX_MAP_NPCS
        If X = MapNpc(i).X And Y = MapNpc(i).Y Then
            ' they're hit, remove it
            ClearProjectile index, PlayerProjectile
            Exit Sub
        End If
    Next
    
    ' if projectile is not loaded, load it, female dog.
    If DDS_Projectile(PicNum) Is Nothing Then
        Call InitDDSurf("projectiles\" & PicNum, DDSD_Projectile(PicNum), DDS_Projectile(PicNum))
    End If
    
    ' get positioning in the texture
    With rec
        .Top = 0
        .Bottom = SIZE_Y
        .Left = Player(index).ProjecTile(PlayerProjectile).direction * SIZE_X
        .Right = .Left + SIZE_X
    End With

    ' blt the projectile
    Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Projectile(PicNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltProjectile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawChatBubble(ByVal index As Long)
Dim theArray() As String, X As Long, Y As Long, i As Long, MaxWidth As Long, xwidth As Long, yheight As Long, colour As Long, x3 As Long, y3 As Long
    
Dim MMx As Long
Dim MMy As Long
    
Dim TOPLEFTrect As RECT
Dim TOPCENTERrect As RECT
Dim TOPRIGHTrect As RECT
Dim MIDDLELEFTrect As RECT
Dim MIDDLECENTERrect As RECT
Dim MIDDLERIGHTrect As RECT
Dim BOTTOMLEFTrect As RECT
Dim BOTTOMCENTERrect As RECT
Dim BOTTOMRIGHTrect As RECT
Dim TIPrect As RECT

' DESIGNATE CHATBUBBLE SECTIONS FROM CHATBUBBLE IMAGE
With TOPRIGHTrect
    .Top = 0
    .Bottom = .Top + 4
    .Left = 0
    .Right = .Left + 4
End With

With TOPCENTERrect
    .Top = 0
    .Bottom = .Top + 4
    .Left = 4
    .Right = .Left + 4
End With

With TOPLEFTrect
    .Top = 0
    .Bottom = .Top + 4
    .Left = 8
    .Right = .Left + 4
End With

With MIDDLERIGHTrect
    .Top = 4
    .Bottom = .Top + 4
    .Left = 0
    .Right = .Left + 4
End With

With MIDDLECENTERrect
    .Top = 4
    .Bottom = .Top + 4
    .Left = 4
    .Right = .Left + 4
End With

With MIDDLELEFTrect
    .Top = 4
    .Bottom = .Top + 4
    .Left = 8
    .Right = .Left + 4
End With

With BOTTOMRIGHTrect
    .Top = 8
    .Bottom = .Top + 4
    .Left = 0
    .Right = .Left + 4
End With

With BOTTOMCENTERrect
    .Top = 8
    .Bottom = .Top + 4
    .Left = 4
    .Right = .Left + 4
End With

With BOTTOMLEFTrect
    .Top = 8
    .Bottom = .Top + 4
    .Left = 8
    .Right = .Left + 4
End With

With TIPrect
    .Top = 12
    .Bottom = .Top + 4
    .Left = 0
    .Right = .Left + 4
End With

    Call DDS_BackBuffer.SetForeColor(RGB(255, 255, 255))

    With chatBubble(index)
        If .targetType = TARGET_TYPE_PLAYER Then
            ' it's a player
            If GetPlayerMap(.target) = GetPlayerMap(MyIndex) Then
                ' change the colour depending on access
                colour = QBColor(Yellow)
                
                ' it's on our map - get co-ords
                X = ConvertMapX((Player(.target).X * 32) + Player(.target).XOffset) + 16
                Y = ConvertMapY((Player(.target).Y * 32) + Player(.target).YOffset) - 36
                
                ' word wrap the text
                WordWrap_Array .msg, ChatBubbleWidth, theArray
                
                ' find max width
                For i = 1 To UBound(theArray)
                    If getWidth(TexthDC, theArray(i)) > MaxWidth Then MaxWidth = getWidth(TexthDC, theArray(i))
                Next
                
                ' calculate the new position xwidth relative to DDS_ChatBubble and yheight relative to DDS_ChatBubble
                xwidth = 10 + MaxWidth ' the first five is just air.
                yheight = 10 + (UBound(theArray) * 3) ' the first three are just air.
                
                ' Compensate the yheight drift
                Y = Y - yheight
                
                ' render bubble
                
                ' top left
                ' RenderTexture Tex_GUI(37), xwidth - 9, yheight - 5, 0, 0, 9, 5, 9, 5
                Call Engine_BltFast(X + (xwidth + 4), Y - (yheight - 4), DDS_ChatBubble, TOPLEFTrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
                ' top center
                ' RenderTexture Tex_GUI(37), xwidth + MaxWidth, yheight - 5, 119, 0, 9, 5, 9, 5
                For x3 = X - (xwidth - 8) To X + (xwidth)
                    Call Engine_BltFast(x3, Y - (yheight - 4), DDS_ChatBubble, TOPCENTERrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Next x3
                
                ' top right
                ' RenderTexture Tex_GUI(37), xwidth, yheight - 5, 9, 0, MaxWidth, 5, 5, 5
                Call Engine_BltFast(X - (xwidth - 4), Y - (yheight - 4), DDS_ChatBubble, TOPRIGHTrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
                ' middle left
                ' RenderTexture Tex_GUI(37), xwidth - 9, y, 0, 19, 9, 6, 9, 6
                For y3 = Y - (yheight - 8) To Y + (yheight)
                    Call Engine_BltFast(X + (xwidth + 4), y3, DDS_ChatBubble, MIDDLELEFTrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Next y3
                
                ' middle center
                ' RenderTexture Tex_GUI(37), xwidth + MaxWidth, y, 119, 19, 9, 6, 9, 6
                For y3 = Y - (yheight - 8) To Y + (yheight)
                    For x3 = X - (xwidth - 8) To X + (xwidth)
                        Call Engine_BltFast(x3, y3, DDS_ChatBubble, MIDDLECENTERrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Next x3
                Next y3
                
                ' middle right
                ' RenderTexture Tex_GUI(37), xwidth, y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
                For y3 = Y - (yheight - 8) To Y + (yheight)
                    Call Engine_BltFast(X - (xwidth - 4), y3, DDS_ChatBubble, MIDDLERIGHTrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Next y3

                ' bottom left
                ' RenderTexture Tex_GUI(37), xwidth + (MaxWidth \ 2) + 6, y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
                Call Engine_BltFast(X + (xwidth + 4), Y + (yheight + 4), DDS_ChatBubble, BOTTOMLEFTrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
                ' bottom center
                ' RenderTexture Tex_GUI(37), xwidth - 9, yheight, 0, 6, 9, (UBound(theArray) * 12), 9, 1
                For x3 = X - (xwidth - 8) To X + (xwidth)
                    Call Engine_BltFast(x3, Y + (yheight + 4), DDS_ChatBubble, BOTTOMCENTERrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Next x3
                
                ' bottom right
                ' RenderTexture Tex_GUI(37), xwidth + MaxWidth, yheight, 119, 6, 9, (UBound(theArray) * 12), 9, 1
                Call Engine_BltFast(X - (xwidth - 4), Y + (yheight + 4), DDS_ChatBubble, BOTTOMRIGHTrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                                
                ' little pointy bit
                ' RenderTexture Tex_GUI(37), x - 5, y, 58, 19, 11, 11, 11, 11
                Call Engine_BltFast(X, Y + (yheight + 8), DDS_ChatBubble, TIPrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                                
                ' Lock the backbuffer so we can draw text and names
                TexthDC = DDS_BackBuffer.GetDC
    
                
                ' render each line centralised
                Y = Y - (yheight - 8)
                For i = 1 To UBound(theArray)
                    DrawTextNoShadow TexthDC, X - (getWidth(TexthDC, theArray(i)) - 10), Y, theArray(i), QBColor(White) ' .colour
                    Y = Y + 12
                Next
            
                ' Release DC
                DDS_BackBuffer.ReleaseDC TexthDC
            
            End If
        End If
                
        ' check if it's timed out - close it if so
        If .Timer + 5000 < GetTickCount Then
            .active = False
        End If
    End With
End Sub

Sub BltGold(ByVal amount As Long)

Dim colour As Long
If amount = GetPlayerMaxMoney(MyIndex) Then
    colour = QBColor(BrightGreen)
Else
    colour = QBColor(White)
End If
frmMain.lblGold.Caption = Format$(amount, "#,###,###,###") & " " & GetTranslation("Rupias")
frmMain.lblGold.ForeColor = colour
End Sub


Public Sub BltPlayerCustomSpriteLayer(ByVal index As Long, ByRef layer As SpriteLayer)
Dim Anim As Byte, i As Long, X As Long, Y As Long
Dim spritetop As Long
Dim rec As DxVBLib.RECT
Dim attackspeed As Long
Dim sprite As Long
    
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    'get the sprite
    If IsLayerUsingPlayerSprite(layer) Then
        sprite = GetPlayerSprite(index)
    Else
        sprite = GetLayerSprite(layer)
    End If

    If sprite < 1 Or sprite > NumCharacters Then Exit Sub
    
    CharacterTimer(sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(sprite) Is Nothing Then
        Call InitDDSurf("characters\" & sprite, DDSD_Character(sprite), DDS_Character(sprite))
    End If

    ' speed from weapon
    If GetPlayerEquipment(index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(index, Weapon)).Speed
    Else
        attackspeed = 1000
    End If
    
    
    'no fixed anim set, find it then

    ' Reset frame
    If Player(index).step = 3 Then
        Anim = 0
    ElseIf Player(index).step = 1 Then
        Anim = 2
    End If

    ' Check for attacking animation
    If Player(index).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If Player(index).Attacking = 1 Then
            Anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case GetPlayerDir(index)
            Case DIR_UP
                If (Player(index).YOffset > 8) Then Anim = Player(index).step
            Case DIR_DOWN
                If (Player(index).YOffset < -8) Then Anim = Player(index).step
            Case DIR_LEFT
                If (Player(index).XOffset > 8) Then Anim = Player(index).step
            Case DIR_RIGHT
                If (Player(index).XOffset < -8) Then Anim = Player(index).step
        End Select
    End If
    
    

    ' Check to see if we want to stop making him attack
    With Player(index)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    
    Dim dir As Byte
    
    dir = GetPlayerDir(index)
    
    Dim XCenter As Long
    Dim YCenter As Long
    XCenter = GetLayerCenterX(layer, dir)
    YCenter = GetLayerCenterY(layer, dir)
    
    ' Set the left
    Select Case dir
        Case DIR_UP
            spritetop = 3
            YCenter = YCenter - (DDSD_Character(sprite).lHeight / 4) * spritetop
        Case DIR_RIGHT
            spritetop = 2
            YCenter = YCenter - (DDSD_Character(sprite).lHeight / 4) * spritetop
        Case DIR_DOWN
            spritetop = 0
            YCenter = YCenter - (DDSD_Character(sprite).lHeight / 4) * spritetop
        Case DIR_LEFT
            spritetop = 1
            YCenter = YCenter - (DDSD_Character(sprite).lHeight / 4) * spritetop
    End Select
    
    'Anim processing
    If IsLayerUsingCenter(layer) Then
        Anim = GetAnimFromCurrentAnim(GetSpriteLayerFixed(layer), dir, Anim)
    End If
    'If Not IsAnimEnabled(GetSpriteLayerFixed(layer), anim) Then
        'anim = GetClosestAnimFromOne(GetSpriteLayerFixed(layer), anim)
    'End If

    With rec
        .Top = spritetop * (DDSD_Character(sprite).lHeight / 4)
        .Bottom = .Top + (DDSD_Character(sprite).lHeight / 4)
        .Left = Anim * (DDSD_Character(sprite).lWidth / 4)
        .Right = .Left + (DDSD_Character(sprite).lWidth / 4)
    End With
    
    ' Calculate the X
    If Not IsLayerUsingCenter(layer) Then
        X = GetPlayerX(index) * PIC_X + Player(index).XOffset - ((DDSD_Character(sprite).lWidth / 4 - 32) / 2)
        
        ' Is the player's height more than 32..?
        If (DDSD_Character(sprite).lHeight) > 32 Then
            ' Create a 32 pixel offset for larger sprites
            Y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - ((DDSD_Character(sprite).lHeight / 4) - 32)
        Else
            ' Proceed as normal
            Y = GetPlayerY(index) * PIC_Y + Player(index).YOffset
        End If
    Else
         XCenter = XCenter - 16
         X = GetPlayerX(index) * PIC_X + Player(index).XOffset - XCenter
         
         ' Is the player's height more than 32..?
        If (DDSD_Character(sprite).lHeight) > 32 Then
            ' Create a 32 pixel offset for larger sprites
            Y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - YCenter
        Else
            ' Proceed as normal
            Y = GetPlayerY(index) * PIC_Y + Player(index).YOffset
        End If
    End If
    


    

    ' render the actual sprite
    Call BltSprite(sprite, X, Y, rec)
    
    If GetPlayerSprite(index) = sprite Then
    ' check for paperdolling
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(index, PaperdollOrder(i))).Paperdoll > 0 Then
                Call BltPaperdoll(X, Y, Item(GetPlayerEquipment(index, PaperdollOrder(i))).Paperdoll, Anim, spritetop)
            End If
        End If
    Next
    
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPlayer", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltPlayerCustomSprite(ByVal index As Long, ByVal CustomSprite As Byte)
    
    Dim i As Byte
    
    For i = 1 To GetCustomSpriteNLayers(CustomSprites(CustomSprite))
        Call BltPlayerCustomSpriteLayer(index, GetCustomSpriteLayer(CustomSprites(CustomSprite), i))
    Next
    
    
End Sub

Public Sub ComputeBltPlayer(ByVal index As Long)

    Dim CustomSprite As Byte
    'CustomSprite = Player(Index).CustomSprite
    CustomSprite = GetPlayerCustomSprite(index)

    If CustomSprite = 0 Or CustomSprite > MAX_CUSTOM_SPRITES Then
        Call BltPlayer(index)
    Else
        Call BltPlayerCustomSprite(index, CustomSprite)
    End If
End Sub


Sub bltFringeResource()
Dim HT As Long
Dim X As Long
Dim Y As Long
Dim i As Long

If NumResources > 0 Then
    If Resources_Init Then
        If Resource_Index > 0 Then
            For HT = 1 To Player_HighIndex
                If GetPlayerMap(HT) = GetPlayerMap(MyIndex) Then
                    i = GetResourceIndex(GetPlayerX(HT), GetPlayerY(HT))
                    If i > -1 Then
                        If MapResource(i).ResourceState = 0 Then
                            BltMapResource i
                        End If
                    End If
                    i = GetResourceIndex(GetPlayerX(HT), GetPlayerY(HT) + 1)
                    If i > -1 Then
                        If MapResource(i).ResourceState = 0 Then
                            BltMapResource i
                        End If
                    End If
                End If
            Next HT
        End If
    End If
End If

End Sub
Sub bltMaskResource()
Dim i As Long

If NumResources > 0 Then
    If Resources_Init Then
        
        If Resource_Index > 0 Then
            Dim MaxX As Long, MaxY As Long
            If map.MaxX > UBound(MapResources, 1) Then
                MaxX = UBound(MapResources, 1)
            Else
                MaxX = TileView.Right
            End If
            
            If map.MaxY > UBound(MapResources, 2) Then
                MaxY = UBound(MapResources, 2)
            Else
                MaxY = TileView.Bottom
            End If
            
            Dim X As Long, Y As Long
            
                
            For X = TileView.Left To MaxX
                For Y = TileView.Top To MaxY
                Dim r As Long
                If X >= 0 And Y >= 0 Then
                    r = MapResources(X, Y)
                    If r > 0 Then
                        BltMapResource r
                    End If
                End If
            Next
        Next
        End If
    End If
End If

End Sub
Public Sub BltRupee(ByVal index As Long)
Dim rec As DxVBLib.RECT
  ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

        rec.Top = 0
        rec.Bottom = 0
        rec.Left = 0
        rec.Right = 0
        Engine_BltFast Camera.Left + 665, Camera.Top + 572, DDS_Rupee, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
   
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltRupee", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub BltHearts(ByVal index As Long)
Dim rec As DxVBLib.RECT
  ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

        rec.Top = 0
        rec.Bottom = 0
        rec.Left = 0
        rec.Right = 0
        Engine_BltFast Camera.Left + 7, Camera.Top + 5, DDS_Hearts, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHearts", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub BltHealth(ByVal index As Long)
Dim rec As DxVBLib.RECT
  ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

        rec.Top = 0
        rec.Bottom = 30
        rec.Left = 0
        rec.Right = 288 * (CSng((GetPlayerVital(MyIndex, Vitals.HP)))) / CSng(GetPlayerMaxVital(MyIndex, HP))
        Engine_BltFast Camera.Left + 7, Camera.Top + 5, DDS_Health, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
   
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHealth", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub BltMagicBar(ByVal index As Long)
Dim rec As DxVBLib.RECT
  ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

        rec.Top = 0
        rec.Bottom = 0
        rec.Left = 0
        rec.Right = 0
        Engine_BltFast Camera.Left + 5, Camera.Top + 35, DDS_MagicBar, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHearts", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Sub BltMiniMap()
Dim i As Long
Dim X As Integer, Y As Integer
Dim CameraX As Long, CameraY As Long

'potentially re-write this to only use a certain area around the player, and not show the entire map.

Dim MapX As Long, MapY As Long

'If Not GetRealTickCount > tmrWaitBlt + 2 Then Exit Sub
'tmrWaitBlt = GetRealTickCount

' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler


' Map size
MapX = map.MaxX
MapY = map.MaxY

' ****************
' ** Rectangles **
' ****************
' Blank


' Set attributes in map
For X = 0 To MapX
For Y = 0 To MapY

' Camera loc
CameraX = Camera.Left + 25 + (X * 4)
CameraY = Camera.Top + 65 + (Y * 4)


' Blank tile
'Engine_BltFast CameraX, CameraY, DDS_MiniMap, BlankRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

Select Case map.Tile(X, Y).Type
Case TILE_TYPE_BLOCKED
Engine_BltFast CameraX, CameraY, DDS_MiniMap, BlockRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
Case TILE_TYPE_WARP
Engine_BltFast CameraX, CameraY, DDS_MiniMap, WarpRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
Case TILE_TYPE_ITEM
Engine_BltFast CameraX, CameraY, DDS_MiniMap, ItemRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
Case TILE_TYPE_SHOP
Engine_BltFast CameraX, CameraY, DDS_MiniMap, ShopRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
Case Else 'only need to draw blank when it's actually blank!
If Options.MiniMapBltElse Then Engine_BltFast CameraX, CameraY, DDS_MiniMap, BlankRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End Select
Next Y
Next X

' Set players in mini map
For i = 1 To Player_HighIndex
If IsPlaying(i) Then
' Player loc
'X = Player(i).X
'Y = Player(i).Y

' Camera loc
CameraX = Camera.Left + 25 + (Player(i).X * 4)
CameraY = Camera.Top + 65 + (Player(i).Y * 4)


Select Case Player(i).PK
Case NO
Call Engine_BltFast(CameraX, CameraY, DDS_MiniMap, PlayerRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Case YES
Call Engine_BltFast(CameraX, CameraY, DDS_MiniMap, PlayerPkRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Select
End If
Next i
' Set npcs in mini map
For i = 1 To Npc_HighIndex
If MapNpc(i).num > 0 Then
' Npc loc
'X = MapNpc(i).X
'Y = MapNpc(i).Y

' Camera loc
CameraX = Camera.Left + 25 + (MapNpc(i).X * 4)
CameraY = Camera.Top + 65 + (MapNpc(i).Y * 4)

Select Case NPC(MapNpc(i).num).Behaviour
    Case NPC_BEHAVIOUR_ATTACKONSIGHT
        Call Engine_BltFast(CameraX, CameraY, DDS_MiniMap, NpcAttackerRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
        Call Engine_BltFast(CameraX, CameraY, DDS_MiniMap, NpcAttackerRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Case NPC_BEHAVIOUR_SHOPKEEPER
        Call Engine_BltFast(CameraX, CameraY, DDS_MiniMap, NpcShopRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Case NPC_BEHAVIOUR_FRIENDLY
        Call Engine_BltFast(CameraX, CameraY, DDS_MiniMap, NpcOtherRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Case Else
        Call Engine_BltFast(CameraX, CameraY, DDS_MiniMap, NpcOtherRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End Select
End If
Next i
'tmrWaitBlt = GetRealTickCount
'DoEvents
' Error handler
Exit Sub
errorhandler:
HandleError "BltMiniMap", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
