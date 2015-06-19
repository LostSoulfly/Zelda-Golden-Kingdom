Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long


Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim FileName As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\data files\logs\errors.txt"
    Open FileName For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleError", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If LCase$(dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not RAW Then
        If LenB(dir(App.Path & "\" & FileName)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(dir(FileName)) > 0 Then
            FileExist = True
        End If
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' gets a string from a text file
Public Function GetVar(File As String, header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, header As String, Var As String, value As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call WritePrivateProfileString$(header, Var, value, File)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveOptions()
Dim FileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\Data Files\config.ini"
    
    Call PutVar(FileName, "Options", "Game_Name", Trim$(Options.Game_Name))
    Call PutVar(FileName, "Options", "Username", Trim$(Options.Username))
    Call PutVar(FileName, "Options", "Password", Trim$(Options.Password))
    Call PutVar(FileName, "Options", "SavePass", Str(Options.SavePass))
    Call PutVar(FileName, "Options", "IP", Options.ip)
    Call PutVar(FileName, "Options", "Port", Str(Options.Port))
    Call PutVar(FileName, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(FileName, "Options", "Music", Str(Options.Music))
    Call PutVar(FileName, "Options", "Sound", Str(Options.Sound))
    Call PutVar(FileName, "Options", "Debug", Str(Options.Debug))
    Call PutVar(FileName, "Options", "Names", Str(Options.Names))
    Call PutVar(FileName, "Options", "Level", Str(Options.Level))
    Call PutVar(FileName, "Options", "Chat", Str(Options.Chat))
    Call PutVar(FileName, "Options", "SafeMode", Str(Options.SafeMode))
    Call PutVar(FileName, "Options", "DefaultVolume", Str(Options.DefaultVolume))
    Call PutVar(FileName, "Options", "MiniMap", Str(Options.MiniMap))
    Call PutVar(FileName, "Options", "MappingMode", Str(Options.MappingMode))
    Call PutVar(FileName, "Options", "ChatToScreen", Str(Options.ChatToScreen))
    
    Dim i As Byte
    For i = 1 To ChatType.ChatType_Count - 1
        Call PutVar(FileName, "ChatOptions", Str(i), Str(BTI(Options.ActivatedChats(i))))
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
Dim FileName As String
Dim i As Byte
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\Data Files\config.ini"
    
    If Not FileExist(FileName, True) Then
        Options.Game_Name = "Eclipse Origins"
        Options.Password = vbNullString
        Options.SavePass = 0
        Options.Username = vbNullString
        Options.ip = "127.0.0.1"
        Options.Port = 7001
        Options.MenuMusic = vbNullString
        Options.Music = 1
        Options.Sound = 1
        Options.Debug = 0
        Options.Names = 1
        Options.Level = 0
        Options.Chat = 1
        Options.SafeMode = NO
        Options.DefaultVolume = 50
        Options.MiniMap = 1
        Options.MappingMode = 0
        Options.ChatToScreen = 1
        For i = 1 To ChatType.ChatType_Count - 1
           Options.ActivatedChats(i) = True
        Next
        SaveOptions
    Else
        Options.Game_Name = GetVar(FileName, "Options", "Game_Name")
        Options.Username = GetVar(FileName, "Options", "Username")
        Options.Password = GetVar(FileName, "Options", "Password")
        Options.SavePass = Val(GetVar(FileName, "Options", "SavePass"))
        Options.ip = GetVar(FileName, "Options", "IP")
        Options.Port = Val(GetVar(FileName, "Options", "Port"))
        Options.MenuMusic = GetVar(FileName, "Options", "MenuMusic")
        Options.Music = GetVar(FileName, "Options", "Music")
        Options.Sound = GetVar(FileName, "Options", "Sound")
        Options.Debug = GetVar(FileName, "Options", "Debug")
        Options.Names = GetVar(FileName, "Options", "Names")
        Options.Level = GetVar(FileName, "Options", "Level")
        Options.Chat = GetVar(FileName, "Options", "Chat")
        Options.SafeMode = GetVar(FileName, "Options", "SafeMode")
        Options.DefaultVolume = GetVar(FileName, "Options", "DefaultVolume")
        Options.MiniMap = GetVar(FileName, "Options", "MiniMap")
        Options.MappingMode = GetVar(FileName, "Options", "MappingMode")
        Options.ChatToScreen = GetVar(FileName, "Options", "ChatToScreen")
        For i = 1 To ChatType.ChatType_Count - 1
           Options.ActivatedChats(i) = STB(GetVar(FileName, "ChatOptions", Str(i)))
        Next
    End If
    
    ' show in GUI
    If Options.Music = 0 Then
        frmMain.optMOff.value = True
    Else
        frmMain.optMOn.value = True
    End If
    
    If Options.Sound = 0 Then
        frmMain.optSOff.value = True
    Else
        frmMain.optSOn.value = True
    End If
    
    If Options.Names = 0 Then
        frmMain.optNOff.value = True
    Else
        frmMain.optNOn.value = True
    End If
    
    If Options.Level = 0 Then
        frmMain.optLvlOff.value = True
    Else
        frmMain.optLvlOn.value = True
    End If
    
    If Options.SafeMode = 0 Then
        frmMain.optSafeOff = True
    Else
        frmMain.optSafeOn = True
    End If
    
    If Options.MiniMap = 0 Then
        frmMain.optMiniMapOff.value = True
    Else
        frmMain.optMiniMapOn.value = True
    End If
    
        'chat options
    If Options.ChatToScreen = 1 Then
        frmMain.txtChat.Visible = True
        frmMain.ChatOpts(0).Top = 444
        frmMain.ChatOpts(1).Top = 444
        frmMain.ChatOpts(2).Top = 444
        frmMain.ChatOpts(3).Top = 444
    ElseIf Options.ChatToScreen = 2 Then
        frmMain.txtChat.Visible = False
        frmMain.ChatOpts(0).Top = 594
        frmMain.ChatOpts(1).Top = 594
        frmMain.ChatOpts(2).Top = 594
        frmMain.ChatOpts(3).Top = 594
    Else
        frmMain.txtChat.Visible = False
        frmMain.ChatOpts(0).Visible = False
        frmMain.ChatOpts(1).Visible = False
        frmMain.ChatOpts(2).Visible = False
        frmMain.ChatOpts(3).Visible = False
    End If
    
    frmMain.scrlVolume.value = Options.DefaultVolume
    DefaultVolume = Options.DefaultVolume
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal mapnum As Long)
Dim FileName As String
Dim f As Long
Dim X As Long
Dim y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & MAP_PATH & "map" & mapnum & MAP_EXT

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , map.Name
    Put #f, , map.Music
    Put #f, , map.Revision
    Put #f, , map.moral
    Put #f, , map.Up
    Put #f, , map.Down
    Put #f, , map.Left
    Put #f, , map.Right
    Put #f, , map.BootMap
    Put #f, , map.BootX
    Put #f, , map.BootY
    Put #f, , map.MaxX
    Put #f, , map.MaxY
    Put #f, , map.Weather

    For X = 0 To map.MaxX
        For y = 0 To map.MaxY
            Put #f, , map.Tile(X, y)
        Next

        DoEvents
    Next

    For X = 1 To MAX_MAP_NPCS
        Put #f, , map.NPC(X)
    Next
    
    For X = 1 To MAX_MAP_NPCS
        Put #f, , map.NPCSProperties(X).movement
        Put #f, , map.NPCSProperties(X).Action
    Next

    For X = 1 To Max_States - 1
        Put #f, , map.AllowedStates(X)
    Next
    Close #f
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal mapnum As Long)
Dim FileName As String
Dim f As Long
Dim X As Long
Dim y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & MAP_PATH & "map" & mapnum & MAP_EXT
    ClearMap
    f = FreeFile
    Open FileName For Binary As #f
    Get #f, , map.Name
    Get #f, , map.Music
    Get #f, , map.Revision
    Get #f, , map.moral
    Get #f, , map.Up
    Get #f, , map.Down
    Get #f, , map.Left
    Get #f, , map.Right
    Get #f, , map.BootMap
    Get #f, , map.BootX
    Get #f, , map.BootY
    Get #f, , map.MaxX
    Get #f, , map.MaxY
    Get #f, , map.Weather
    
    ' have to set the tile()
    ReDim map.Tile(0 To map.MaxX, 0 To map.MaxY)

    For X = 0 To map.MaxX
        For y = 0 To map.MaxY
            Get #f, , map.Tile(X, y)
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Get #f, , map.NPC(X)
    Next
    
    For X = 1 To MAX_MAP_NPCS
        Get #f, , map.NPCSProperties(X).movement
        Get #f, , map.NPCSProperties(X).Action
    Next
    
    For X = 1 To Max_States - 1
        Get #f, , map.AllowedStates(X)
    
    Next

    Close #f
    ClearTempTile
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckTilesets()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "\tilesets\" & i & GFX_EXT)
        NumTileSets = NumTileSets + 1
        i = i + 1
    Wend
    
    If NumTileSets = 0 Then Exit Sub
    
    ReDim DDS_Tileset(1 To NumTileSets)
    ReDim DDSD_Tileset(1 To NumTileSets)
    ReDim TilesetTimer(1 To NumTileSets)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckTilesets", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckCharacters()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "characters\" & i & GFX_EXT)
        NumCharacters = NumCharacters + 1
        i = i + 1
    Wend
    
    If NumCharacters = 0 Then Exit Sub

    ReDim DDS_Character(1 To NumCharacters)
    ReDim DDSD_Character(1 To NumCharacters)
    ReDim CharacterTimer(1 To NumCharacters)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckCharacters", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPaperdolls()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "paperdolls\" & i & GFX_EXT)
        NumPaperdolls = NumPaperdolls + 1
        i = i + 1
    Wend
    
    If NumPaperdolls = 0 Then Exit Sub

    ReDim DDS_Paperdoll(1 To NumPaperdolls)
    ReDim DDSD_Paperdoll(1 To NumPaperdolls)
    ReDim PaperdollTimer(1 To NumPaperdolls)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPaperdolls", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimations()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "animations\" & i & GFX_EXT)
        NumAnimations = NumAnimations + 1
        i = i + 1
    Wend
    
    If NumAnimations = 0 Then Exit Sub

    ReDim DDS_Animation(1 To NumAnimations)
    ReDim DDSD_Animation(1 To NumAnimations)
    ReDim AnimationTimer(1 To NumAnimations)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Items\" & i & GFX_EXT)
        NumItems = NumItems + 1
        i = i + 1
    Wend
    
    If NumItems = 0 Then Exit Sub

    ReDim DDS_Item(1 To NumItems)
    ReDim DDSD_Item(1 To NumItems)
    ReDim ItemTimer(1 To NumItems)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckResources()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Resources\" & i & GFX_EXT)
        NumResources = NumResources + 1
        i = i + 1
    Wend
    
    If NumResources = 0 Then Exit Sub

    ReDim DDS_Resource(0 To NumResources)
    ReDim DDSD_Resource(0 To NumResources)
    ReDim ResourceTimer(1 To NumResources)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckSpellIcons()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "SpellIcons\" & i & GFX_EXT)
        NumSpellIcons = NumSpellIcons + 1
        i = i + 1
    Wend
    
    If NumSpellIcons = 0 Then Exit Sub

    ReDim DDS_SpellIcon(1 To NumSpellIcons)
    ReDim DDSD_SpellIcon(1 To NumSpellIcons)
    ReDim SpellIconTimer(1 To NumSpellIcons)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckSpellIcons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFaces()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Faces\" & i & GFX_EXT)
        NumFaces = NumFaces + 1
        i = i + 1
    Wend
    
    If NumFaces = 0 Then Exit Sub

    ReDim DDS_Face(1 To NumFaces)
    ReDim DDSD_Face(1 To NumFaces)
    ReDim FaceTimer(1 To NumFaces)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckFaces", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearPlayer(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimInstance(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimation(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimations()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNPC(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    NPC(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNpcs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpell(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpells()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShop(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShops()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResource(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResources()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(map), LenB(map))
    map.Name = vbNullString
    map.MaxX = MAX_MAPX
    map.MaxY = MAX_MAPY
    ReDim map.Tile(0 To map.MaxX, 0 To map.MaxY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpcs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********************
' ** Player functions **
' **********************
Function GetPlayerName(ByVal Index As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Name = Name
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(Index).Class
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Class = ClassNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).sprite
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal sprite As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).sprite = sprite
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Level = Level
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).Exp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Exp = Exp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Access = Access
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).PK = PK
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).vital(vital)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal vital As Vitals, ByVal value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).vital(vital) = value

    If GetPlayerVital(Index, vital) > GetPlayerMaxVital(Index, vital) Then
        Player(Index).vital(vital) = GetPlayerMaxVital(Index, vital)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerMaxVital = Player(Index).MaxVital(vital)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMaxVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerStat(ByVal Index As Long, stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).stat(stat)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerStat(ByVal Index As Long, stat As Stats, ByVal value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    If value <= 0 Then value = 1
    If value > MAX_BYTE Then value = MAX_BYTE
    Player(Index).stat(stat) = value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).points
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal points As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).points = points
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).map
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal mapnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).map = mapnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).X
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).X = X
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).y
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).y = y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).dir
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal dir As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).dir = dir
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invslot).num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long, ByVal ItemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).num = ItemNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invslot).value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = InvNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


' projectiles
Public Sub CheckProjectiles()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Projectiles\" & i & GFX_EXT)
        NumProjectiles = NumProjectiles + 1
        i = i + 1
    Wend
    
    If NumProjectiles = 0 Then Exit Sub

    ReDim DDS_Projectile(1 To NumProjectiles)
    ReDim DDSD_Projectile(1 To NumProjectiles)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearProjectile(ByVal Index As Long, ByVal PlayerProjectile As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Player(Index).ProjecTile(PlayerProjectile)
        .direction = 0
        .Pic = 0
        .TravelTime = 0
        .X = 0
        .y = 0
        .range = 0
        .Damage = 0
        .Speed = 0
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearProjectile", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function Spin()
Dim d As String
    d = GetPlayerDir(MyIndex)

            If Player(MyIndex).dir = DIR_DOWN Then
                Call SetPlayerDir(MyIndex, DIR_LEFT)
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
            ElseIf Player(MyIndex).dir = DIR_LEFT Then
                Call SetPlayerDir(MyIndex, DIR_UP)
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
            ElseIf Player(MyIndex).dir = DIR_UP Then
                Call SetPlayerDir(MyIndex, DIR_RIGHT)
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
            ElseIf Player(MyIndex).dir = DIR_RIGHT Then
                Call SetPlayerDir(MyIndex, DIR_DOWN)
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
            End If
End Function

Sub ClearDoor(ByVal Index As Long)
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        Call ZeroMemory(ByVal VarPtr(Doors(Index)), LenB(Doors(Index)))
        Doors(Index).Name = vbNullString
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "ClearDoor", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Sub ClearDoors()
Dim i As Long

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        For i = 1 To MAX_DOORS
                Call ClearDoor(i)
        Next

        ' Error handler
        Exit Sub
errorhandler:
        HandleError "ClearDoors", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
Sub ClearMovement(ByVal Index As Long)
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        Call ZeroMemory(ByVal VarPtr(Movements(Index)), LenB(Movements(Index)))
        Movements(Index).Name = vbNullString
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "ClearMovement", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
Sub ClearMovements()
Dim i As Long

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        For i = 1 To MAX_MOVEMENTS
                Call ClearMovement(i)
        Next

        ' Error handler
        Exit Sub
errorhandler:
        HandleError "ClearDoors", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Sub ClearAction(ByVal Index As Long)
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        Call ZeroMemory(ByVal VarPtr(Actions(Index)), LenB(Actions(Index)))
        Actions(Index).Name = vbNullString
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "ClearAction", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
Sub ClearActions()
Dim i As Long

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        For i = 1 To MAX_ACTIONS
                Call ClearAction(i)
        Next

        ' Error handler
        Exit Sub
errorhandler:
        HandleError "ClearDoors", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Function GetPlayerVisible(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If Index > MAX_PLAYERS Then Exit Function
GetPlayerVisible = Player(Index).Visible

' Error handler
Exit Function
errorhandler:
HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Function
End Function

Sub SetPlayerVisible(ByVal Index As Long, ByVal Visible As Long)
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If Index > MAX_PLAYERS Then Exit Sub
Player(Index).Visible = Visible

' Error handler
Exit Sub
errorhandler:
HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Sub ClearPet(ByVal Index As Long)
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        Call ZeroMemory(ByVal VarPtr(Pet(Index)), LenB(Pet(Index)))
        Pet(Index).Name = vbNullString
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "ClearPet", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
Sub ClearPets()
Dim i As Long

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        For i = 1 To MAX_PETS
                Call ClearPet(i)
        Next

        ' Error handler
        Exit Sub
errorhandler:
        HandleError "ClearDoors", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Function GetPlayerPetStat(ByVal Index As Long, stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If Player(Index).ActualPet < 1 Or Player(Index).ActualPet > MAX_PLAYER_PETS Then Exit Function
    GetPlayerPetStat = Player(Index).Pet(Player(Index).ActualPet).StatsAdd(stat)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPetStat(ByVal Index As Long, stat As Stats, ByVal value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    If value <= 0 Then value = 1
    If value > MAX_BYTE Then value = MAX_BYTE
    If Player(Index).ActualPet < 1 Or Player(Index).ActualPet > MAX_PLAYER_PETS Then Exit Sub
    Player(Index).Pet(Player(Index).ActualPet).StatsAdd(stat) = value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPetPOINTS(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    
    If Player(Index).ActualPet < 1 Or Player(Index).ActualPet > MAX_PLAYER_PETS Then Exit Function
    
    GetPlayerPetPOINTS = Player(Index).Pet(Player(Index).ActualPet).points
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPetPOINTS(ByVal Index As Long, ByVal points As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    
    If Player(Index).ActualPet < 1 Or Player(Index).ActualPet > MAX_PLAYER_PETS Then Exit Sub
    
    
    Player(Index).Pet(Player(Index).ActualPet).points = points
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetNpcMaxVital(ByVal NPCNum As Long, ByVal vital As Vitals, Optional ByVal PetOwner As Long = 0) As Long
    Dim X As Long
    Dim TempPetOwner As Long
    
    ' Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If
    
    TempPetOwner = 0
    'Prevent pet system
    If PetOwner > 0 Then
        If PetOwner = MyIndex Then
            If GetPlayerPetMapNPCNum(PetOwner) > 0 Then
                If NPCNum = Pet(Player(PetOwner).Pet(Player(PetOwner).ActualPet).NumPet).NPCNum Then
                    TempPetOwner = PetOwner
                End If
            End If
        End If
    End If
    
    'Only stethic
    If PetOwner <> TempPetOwner Then
        PetOwner = 0
    End If
    
    Select Case PetOwner
    
    Case Is > 0
        'Has owner
        Select Case vital
            Case HP
                GetNpcMaxVital = NPC(NPCNum).HP + ((Player(PetOwner).Pet(Player(PetOwner).ActualPet).Level / 2) + (GetNpcStat(NPCNum, Endurance, PetOwner) / 2) * 10)
            Case mp
                GetNpcMaxVital = 30 + ((Player(PetOwner).Pet(Player(PetOwner).ActualPet).Level / 2) + (GetNpcStat(NPCNum, Intelligence, PetOwner) / 2)) * 10
            End Select
    Case Else
            Select Case vital
            Case HP
                GetNpcMaxVital = NPC(NPCNum).HP
            Case mp
                GetNpcMaxVital = 30 + (NPC(NPCNum).stat(Intelligence) * 10) + 2
            End Select
    End Select

End Function

Public Function GetNpcStat(ByVal NPCNum As Long, ByVal stat As Stats, Optional ByVal PetOwner As Long = 0) As Long
Dim value As Long


If NPCNum < 1 Or NPCNum > MAX_NPCS Then Exit Function

value = NPC(NPCNum).stat(stat)

If PetOwner > 0 And PetOwner < MAX_PLAYERS Then
    If PetOwner = MyIndex Then
        If GetPlayerPetMapNPCNum(PetOwner) > 0 Then
            If NPCNum = Pet(Player(PetOwner).Pet(Player(PetOwner).ActualPet).NumPet).NPCNum Then
                value = value + Player(PetOwner).Pet(Player(PetOwner).ActualPet).StatsAdd(stat)
            End If
        End If
    End If
End If

GetNpcStat = value

End Function


Function AccountExist(ByVal Name As String) As Boolean
    Dim FileName As String
    FileName = "data\accounts\" & Trim(Name) & ".bin"

    If FileExist(FileName) Then
        AccountExist = True
    End If

End Function

Sub SavePlayer(ByRef Data() As Byte, ByRef Login As String)
    Dim FileName As String
    Dim f As Long
    
    If Trim$(Login) = ".bin" Then Exit Sub

    FileName = App.Path & "\data files\accounts\" & Login
    f = FreeFile
    
    Open FileName For Binary As #f
    Put #f, , Data
    Close #f

End Sub

Sub SaveBank(ByRef Bank As ServerBankRec, ByVal Login As String)
    Dim FileName As String
    Dim f As Long
    

    FileName = App.Path & "\data files\banks\" & Trim$(Login) & ".bin"
    f = FreeFile
    
    Open FileName For Binary As #f
    Put #f, , Bank
    Close #f
    
End Sub

Sub SaveGuild(ByRef Guild As ServerGuildRec, ByVal N As Long)
    Dim FileName As String
    Dim f As Long
    

    FileName = App.Path & "\data files\guilds\Guild" & CStr(N) & ".dat"
    f = FreeFile
    
    Open FileName For Binary As #f
    Put #f, , Guild
    Close #f
    
    Dim GuildName As String
    GuildName = RTrim(Guild.Guild_Name)
    
    FileName = App.Path & "\data files\guildnames\" & GuildName & ".dat"
    f = FreeFile
    
    Open FileName For Binary As #f
    Put #f, , ""
    Close #f
    
End Sub



Function FindChar(ByVal Name As String) As Boolean
    Dim f As Long
    Dim s As String
    f = FreeFile
    Open App.Path & "\data files\accounts\charlist.txt" For Input As #f

    Do While Not EOF(f)
        Input #f, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #f
            Exit Function
        End If

    Loop

    Close #f
End Function

Sub ClearCustomSprite(ByVal Index As Long)
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        Call ZeroMemory(ByVal VarPtr(CustomSprites(Index)), LenB(CustomSprites(Index)))
        CustomSprites(Index).Name = vbNullString
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "ClearCustomSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
Sub ClearCustomSprites()
Dim i As Long

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        For i = 1 To MAX_CUSTOM_SPRITES
                Call ClearCustomSprite(i)
        Next

        ' Error handler
        Exit Sub
errorhandler:
        HandleError "ClearDoors", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub

Public Function SetPlayerCustomSprite(ByVal Index As Long, ByVal CustomSprite As Byte)
    If CustomSprite > MAX_CUSTOM_SPRITES Then Exit Function
    Player(Index).CustomSprite = CustomSprite
End Function

Public Function GetPlayerCustomSprite(ByVal Index As Long) As Byte
    If Player(Index).CustomSprite > MAX_CUSTOM_SPRITES Then Exit Function
    GetPlayerCustomSprite = Player(Index).CustomSprite
End Function

