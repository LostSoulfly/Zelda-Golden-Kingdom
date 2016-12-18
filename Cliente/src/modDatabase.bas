Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long


Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim Filename As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Filename = App.Path & "\data\logs\errors.txt"
    Open Filename For Append As #1
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

Public Function FileExist(ByVal Filename As String, Optional RAW As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not RAW Then
        If LenB(dir(App.Path & "\" & Filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(dir(Filename)) > 0 Then
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
Dim Filename As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Filename = App.Path & "\data\config.ini"
    
    Call PutVar(Filename, "Options", "Game_Name", Trim$(Options.Game_Name))
    Call PutVar(Filename, "Options", "Username", Trim$(Options.Username))
    Call PutVar(Filename, "Options", "Password", Trim$(Options.Password))
    Call PutVar(Filename, "Options", "SavePass", Str(Options.SavePass))
    Call PutVar(Filename, "Options", "IP", Options.ip)
    Call PutVar(Filename, "Options", "Port", Str(Options.port))
    Call PutVar(Filename, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(Filename, "Options", "Music", Str(Options.Music))
    Call PutVar(Filename, "Options", "Sound", Str(Options.Sound))
    Call PutVar(Filename, "Options", "Debug", Str(Options.Debug))
    Call PutVar(Filename, "Options", "Names", Str(Options.Names))
    Call PutVar(Filename, "Options", "Level", Str(Options.Level))
    Call PutVar(Filename, "Options", "WASD", Str(Options.WASD))
    Call PutVar(Filename, "Options", "MiniMapBltElse", Str(Options.MiniMapBltElse))
    Call PutVar(Filename, "Options", "Chat", Str(Options.Chat))
    Call PutVar(Filename, "Options", "SafeMode", Str(Options.SafeMode))
    Call PutVar(Filename, "Options", "DefaultVolume", Str(Options.DefaultVolume))
    Call PutVar(Filename, "Options", "MiniMap", Str(Options.MiniMap))
    Call PutVar(Filename, "Options", "MappingMode", Str(Options.MappingMode))
    Call PutVar(Filename, "Options", "ChatToScreen", Str(Options.ChatToScreen))
    
    Dim i As Byte
    For i = 1 To ChatType.ChatType_Count - 1
        Call PutVar(Filename, "ChatOptions", Str(i), Str(BTI(Options.ActivatedChats(i))))
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
Dim Filename As String
Dim i As Byte
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Filename = App.Path & "\data\config.ini"
    
    If Not FileExist(Filename, True) Then
        Options.Game_Name = "Legend of Zelda: The Golden Kingdom"
        Options.Password = vbNullString
        Options.SavePass = 0
        Options.Username = vbNullString
        Options.ip = "trollparty.org"
        Options.port = 4000
        Options.MenuMusic = vbNullString
        Options.Music = 1
        Options.Sound = 1
        Options.Debug = 1
        Options.Names = 1
        Options.Level = 1
        Options.WASD = 1
        Options.Chat = 1
        Options.SafeMode = 0
        Options.DefaultVolume = 50
        Options.MiniMap = 1
        Options.MiniMapBltElse = 1
        Options.MappingMode = 0
        Options.ChatToScreen = 1
        Options.SavePass = 1
        For i = 1 To ChatType.ChatType_Count - 1
           Options.ActivatedChats(i) = True
        Next
        SaveOptions
    Else

        Options.Game_Name = GetVar(Filename, "Options", "Game_Name")
        If Len(Trim$(Replace(Options.Username, vbNullChar, ""))) = 0 Then Options.Username = GetVar(Filename, "Options", "Username")
        If Len(Trim$(Replace(Options.Password, vbNullChar, ""))) = 0 Then Options.Password = GetVar(Filename, "Options", "Password")
        Options.SavePass = Val(GetVar(Filename, "Options", "SavePass"))
        
        'Call PutVar(Filename, "Options", "RequireLauncher", "0")
        
        If Val(GetVar(Filename, "Options", "RequireLauncher")) = 1 Then
            If InStr(1, Command, "-launcher 1") <= 0 Then
                MsgBox "Please launch the game with the Launcher." & vbNewLine & _
                "This keeps your client on the latest version!", vbCritical, "Launcher Required"
                DestroyGame
                End
            Else
                frmMain.Caption = Options.Game_Name & " - Official Server"
            End If
        End If
        
        If Len(Trim$(Replace(Options.ip, vbNullChar, ""))) = 0 Then Options.ip = GetVar(Filename, "Options", "IP")
        If Len(Trim$(Replace(Options.port, vbNullChar, ""))) <= 1 Then Options.port = Val(GetVar(Filename, "Options", "Port"))
        Options.MenuMusic = GetVar(Filename, "Options", "MenuMusic")
        Options.Music = GetVar(Filename, "Options", "Music")
        Options.Sound = GetVar(Filename, "Options", "Sound")
        Options.Debug = GetVar(Filename, "Options", "Debug")
        Options.Names = GetVar(Filename, "Options", "Names")
        Options.Level = GetVar(Filename, "Options", "Level")
        Options.WASD = GetVar(Filename, "Options", "WASD")
        Options.Chat = GetVar(Filename, "Options", "Chat")
        Options.SafeMode = GetVar(Filename, "Options", "SafeMode")
        Options.DefaultVolume = GetVar(Filename, "Options", "DefaultVolume")
        Options.MiniMap = GetVar(Filename, "Options", "MiniMap")
        Options.MiniMapBltElse = GetVar(Filename, "Options", "MiniMapBltElse")
        Options.MappingMode = GetVar(Filename, "Options", "MappingMode")
        Options.ChatToScreen = GetVar(Filename, "Options", "ChatToScreen")
        For i = 1 To ChatType.ChatType_Count - 1
           Options.ActivatedChats(i) = STB(GetVar(Filename, "ChatOptions", Str(i)))
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
    frmMain.cmbChat.ListIndex = Options.ChatToScreen
    'If Options.ChatToScreen = 1 Then
    '    frmMain.txtChat.Visible = True
    'ElseIf Options.ChatToScreen = 2 Then
    '    frmMain.txtChat.Visible = False
    'Else
    '    frmMain.txtChat.Visible = False
    'End If
    
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
Dim Filename As String
Dim f As Long
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Filename = App.Path & MAP_PATH & "map" & mapnum & MAP_EXT

    f = FreeFile
    Open Filename For Binary As #f
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
        For Y = 0 To map.MaxY
            Put #f, , map.Tile(X, Y)
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
Dim Filename As String
Dim f As Long
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Filename = App.Path & MAP_PATH & "map" & mapnum & MAP_EXT
    ClearMap
    f = FreeFile
    Open Filename For Binary As #f
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
        For Y = 0 To map.MaxY
            Get #f, , map.Tile(X, Y)
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
    
    'map.Name = Trim$(map.Name)
    'map.Music = Trim$(map.Music)
    'map.TranslatedName = map.Name

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

Sub ClearPlayer(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
    Player(index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItem(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(index)), LenB(Item(index)))
    Item(index).Name = vbNullString
    Item(index).Desc = vbNullString
    Item(index).Sound = "None."
    
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

Sub ClearAnimInstance(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(index)), LenB(AnimInstance(index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimation(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(index)), LenB(Animation(index)))
    Animation(index).Name = vbNullString
    Animation(index).Sound = "None."
    
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

Sub ClearNPC(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(NPC(index)), LenB(NPC(index)))
    NPC(index).Name = vbNullString
    NPC(index).Sound = "None."
    
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

Sub ClearSpell(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(index)), LenB(Spell(index)))
    Spell(index).Name = vbNullString
    Spell(index).Desc = vbNullString
    Spell(index).Sound = "None."
    
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

Sub ClearShop(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Shop(index).Name = vbNullString
    
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

Sub ClearResource(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).Name = vbNullString
    Resource(index).TranslatedName = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).Sound = "None."
    
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

Sub ClearMapItem(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(index)), LenB(MapItem(index)))
    
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

Sub ClearMapNpc(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapNpc(index)), LenB(MapNpc(index)))
    
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
Function GetPlayerName(ByVal index As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(index).Name)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Name = Name
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(index).Class
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Class = ClassNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(index).sprite
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal sprite As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).sprite = sprite
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(index).Level
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Level = Level
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerExp(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(index).Exp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Exp = Exp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(index).Access
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Access = Access
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(index).PK
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).PK = PK
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(index).vital(vital)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal vital As Vitals, ByVal value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).vital(vital) = value

    If GetPlayerVital(index, vital) > GetPlayerMaxVital(index, vital) Then
        Player(index).vital(vital) = GetPlayerMaxVital(index, vital)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMaxVital(ByVal index As Long, ByVal vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    
    GetPlayerMaxVital = Player(index).MaxVital(vital)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMaxVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerStat(ByVal index As Long, stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(index).stat(stat)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerStat(ByVal index As Long, stat As Stats, ByVal value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    If value <= 0 Then value = 1
    If value > MAX_BYTE Then value = MAX_BYTE
    Player(index).stat(stat) = value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(index).points
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal points As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).points = points
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or index <= 0 Then Exit Function
    GetPlayerMap = Player(index).map
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal mapnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).map = mapnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(index).X
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerX(ByVal index As Long, ByVal X As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).X = X
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(index).Y
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerY(ByVal index As Long, ByVal Y As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Y = Y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(index).dir
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal dir As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).dir = dir
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemNum(ByVal index As Long, ByVal invslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invslot).num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invslot As Long, ByVal ItemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).num = ItemNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal invslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invslot).value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipment = Player(index).Equipment(EquipmentSlot)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    Player(index).Equipment(EquipmentSlot) = InvNum
    
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

Sub ClearProjectile(ByVal index As Long, ByVal PlayerProjectile As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Player(index).ProjecTile(PlayerProjectile)
        .direction = 0
        .Pic = 0
        .TravelTime = 0
        .X = 0
        .Y = 0
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

Sub ClearDoor(ByVal index As Long)
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        Call ZeroMemory(ByVal VarPtr(Doors(index)), LenB(Doors(index)))
        Doors(index).Name = vbNullString
   
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
Sub ClearMovement(ByVal index As Long)
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        Call ZeroMemory(ByVal VarPtr(Movements(index)), LenB(Movements(index)))
        Movements(index).Name = vbNullString
   
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

Sub ClearAction(ByVal index As Long)
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        Call ZeroMemory(ByVal VarPtr(Actions(index)), LenB(Actions(index)))
        Actions(index).Name = vbNullString
   
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

Function GetPlayerVisible(ByVal index As Long) As Long
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If index > MAX_PLAYERS Then Exit Function
GetPlayerVisible = Player(index).Visible

' Error handler
Exit Function
errorhandler:
HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Function
End Function

Sub SetPlayerVisible(ByVal index As Long, ByVal Visible As Long)
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If index > MAX_PLAYERS Then Exit Sub
Player(index).Visible = Visible

' Error handler
Exit Sub
errorhandler:
HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Sub ClearPet(ByVal index As Long)
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        Call ZeroMemory(ByVal VarPtr(Pet(index)), LenB(Pet(index)))
        Pet(index).Name = vbNullString
   
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

Function GetPlayerPetStat(ByVal index As Long, stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    If Player(index).ActualPet < 1 Or Player(index).ActualPet > MAX_PLAYER_PETS Then Exit Function
    GetPlayerPetStat = Player(index).Pet(Player(index).ActualPet).StatsAdd(stat)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPetStat(ByVal index As Long, stat As Stats, ByVal value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    If value <= 0 Then value = 1
    If value > MAX_BYTE Then value = MAX_BYTE
    If Player(index).ActualPet < 1 Or Player(index).ActualPet > MAX_PLAYER_PETS Then Exit Sub
    Player(index).Pet(Player(index).ActualPet).StatsAdd(stat) = value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPetPOINTS(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    
    If Player(index).ActualPet < 1 Or Player(index).ActualPet > MAX_PLAYER_PETS Then Exit Function
    
    GetPlayerPetPOINTS = Player(index).Pet(Player(index).ActualPet).points
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPetPOINTS(ByVal index As Long, ByVal points As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    
    If Player(index).ActualPet < 1 Or Player(index).ActualPet > MAX_PLAYER_PETS Then Exit Sub
    
    
    Player(index).Pet(Player(index).ActualPet).points = points
    
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
    Dim Filename As String
    Filename = "data\accounts\" & Trim$(Name) & ".bin"

    If FileExist(Filename) Then
        AccountExist = True
    End If

End Function

Sub SavePlayer(ByRef Data() As Byte, ByRef Login As String)
    Dim Filename As String
    Dim f As Long
    
    If Trim$(Login) = ".bin" Then Exit Sub

    Filename = App.Path & "\data\accounts\" & Login
    f = FreeFile
    
    Open Filename For Binary As #f
    Put #f, , Data
    Close #f

End Sub

Sub SaveBank(ByRef Bank As ServerBankRec, ByVal Login As String)
    Dim Filename As String
    Dim f As Long
    

    Filename = App.Path & "\data\banks\" & Trim$(Login) & ".bin"
    f = FreeFile
    
    Open Filename For Binary As #f
    Put #f, , Bank
    Close #f
    
End Sub

Sub SaveGuild(ByRef Guild As ServerGuildRec, ByVal N As Long)
    Dim Filename As String
    Dim f As Long
    

    Filename = App.Path & "\data\guilds\Guild" & CStr(N) & ".dat"
    f = FreeFile
    
    Open Filename For Binary As #f
    Put #f, , Guild
    Close #f
    
    Dim GuildName As String
    GuildName = RTrim(Guild.Guild_Name)
    
    Filename = App.Path & "\data\guildnames\" & GuildName & ".dat"
    f = FreeFile
    
    Open Filename For Binary As #f
    Put #f, , ""
    Close #f
    
End Sub



Function FindChar(ByVal Name As String) As Boolean
    Dim f As Long
    Dim s As String
    f = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Input As #f

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

Sub ClearCustomSprite(ByVal index As Long)
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        Call ZeroMemory(ByVal VarPtr(CustomSprites(index)), LenB(CustomSprites(index)))
        CustomSprites(index).Name = vbNullString
   
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

Public Function SetPlayerCustomSprite(ByVal index As Long, ByVal CustomSprite As Byte)
    If CustomSprite > MAX_CUSTOM_SPRITES Then Exit Function
    Player(index).CustomSprite = CustomSprite
End Function

Public Function GetPlayerCustomSprite(ByVal index As Long) As Byte
    If Player(index).CustomSprite > MAX_CUSTOM_SPRITES Then Exit Function
    GetPlayerCustomSprite = Player(index).CustomSprite
End Function

