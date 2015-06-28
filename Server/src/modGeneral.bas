Attribute VB_Name = "modGeneral"
Option Explicit
' Get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

Public Sub Main()

'setup the translation stuff
LangTo = "en"
LangFrom = "es"
strTransPath = App.Path & "\" & LangTo & ".dat"
strOrigPath = App.Path & "\" & LangFrom & "-" & LangTo & ".dat"

frmTransLog.Show
frmTransLog.txtLog.Text = "GTranslate and modTranslate by Dragoon/LostSoulFly!"

    Call InitServer
End Sub

Public Sub InitServer()
    Dim i As Long
    Dim F As Long
    Dim time1 As Long
    Dim time2 As Long
    Call InitMessages
    time1 = GetRealTickCount
    frmServer.Show
    ' Initialize the random-number generator
    Randomize ', seed
    
    Rainon = True
    
    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\Data\", "accounts"
    ChkDir App.Path & "\Data\", "animations"
    ChkDir App.Path & "\Data\", "banks"
    ChkDir App.Path & "\Data\", "items"
    ChkDir App.Path & "\Data\", "logs"
    ChkDir App.Path & "\Data\", "maps"
    ChkDir App.Path & "\Data\", "npcs"
    ChkDir App.Path & "\Data\", "resources"
    ChkDir App.Path & "\Data\", "shops"
    ChkDir App.Path & "\Data\", "spells"
    ChkDir App.Path & "\Data\", "guildnames"
    ChkDir App.Path & "\Data\", "guilds"
    ChkDir App.Path & "\Data\", "quests"
    ChkDir App.Path & "\Data\", "doors"
    ChkDir App.Path & "\Data\", "movements"

    ' set quote character
    vbQuote = ChrW$(34) ' "
    
    ' load options, set if they dont exist
    If Not FileExist(App.Path & "\data\options.ini", True) Then
        Options.Game_Name = "The Legend of Zelda: The Golden Kingdom"
        Options.Port = 4000
        Options.MOTD = "Welcome to The Legend of Zelda: The Golden Kingdom"
        Options.Website = "http://trollparty.org/"
        Options.DisableAdmins = 0
        Options.Update = "http://trollparty.org/Zelda/Launcher.zip"
        Options.Instructions = "Your client is out of date! Please run the Launcher to be updated to the current version." 'GetTranslation("Descargar")
        Options.ExpMultiplier = 1
        
        SaveOptions
    Else
        LoadOptions
    End If
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Options.Port
    frmServer.spExp.Value = Options.ExpMultiplier
    
    ' Init all the player sockets
    Call SetStatus("Initializing player array...")

    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next

    ' Serves as a constructor
    Call ClearGameData
    
    Call LoadGameData
    Call SetStatus("Initializing Temp Data...")
    Call InitTempTiles
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    'Call SetStatus("Creating map cache...")
    'Call CreateFullMapCache
    Call SetStatus("Loading System Tray...")
    Call LoadSystemTray
    

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("data\accounts\charlist.txt") Then
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Output As #F
        Close #F
    End If
    
    Call Set_Default_Guild_Ranks
    
    Call InitTempMaps
    
    'Spawning system
    Call SpawnRandomNPCS(NPC_SKULLTULA, SKULLTULAS)
    
    'sleep time
    CalculateSleepTime
    

    GenerateExp

    ' Start listening
    frmServer.Socket(0).Listen
    Call UpdateCaption
    time2 = GetRealTickCount
    Call SetStatus("Initialization complete. Server loaded in " & time2 - time1 & "ms.")
    
    ServerLog = True
    Set LogPlayers = New clsGenericSet
    Set GodPlayers = New Collection
    Set SpellPlayers = New Collection
    SetAD ADMIN_DISABLED
    ClearIPTries
    ReadRunningSprites
    CheckGuildNamesFile
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop
End Sub


Public Sub DestroyServer()
    Dim i As Long
    ServerOnline = False
    
    Call saveLang(strTransPath, langCol, True)
    Call saveLang(strOrigPath, origCol, True)
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("Unloading sockets...")

    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next

    End
    
End Sub

Public Sub RestartServer()
    Dim i As Long
    ServerOnline = False
    
    
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("Unloading sockets...")

    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next
    
    frmServer.Socket(0).Close
    'frmServer.Socket(0).Connect

    Main
End Sub

Public Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status)
    DoEvents
End Sub

Public Sub ClearGameData()
    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTiles
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing Resources...")
    Call ClearResources
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing animations...")
    Call ClearAnimations
    Call SetStatus("Clearing guilds...")
    Call ClearGuilds
    Call SetStatus("Clearing quests...")
    Call ClearQuests
    Call SetStatus("Clearing Doors...")
    Call ClearDoors
    Call SetStatus("Clearing Movements...")
    Call Clearmovements
    
End Sub

Private Sub LoadGameData()
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading Resources...")
    Call LoadResources
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading animations...")
    Call LoadAnimations
    
    'ALATAR
    Call SetStatus("Loading quests...")
    Call ClearQuests
    Call LoadQuests
    '/ALATAR
    'Call migrateQuests
    
    Call SetStatus("Loading Doors...")
    Call LoadDoors
    
    Call SetStatus("Loading Movements...")
    Call Loadmovements
    
    Call SetStatus("Loading Actions...")
    Call LoadActions
    
    Call SetStatus("Loading Pets...")
    Call LoadPets
    
    Call SetStatus("Loading Custom Sprites...")
    Call LoadCustomSprites
End Sub

Public Sub TextAdd(msg As String)
    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmServer.txtText.Text = vbNullString
        NumLines = 0
    End If

    frmServer.txtText.Text = frmServer.txtText.Text & vbNewLine & Time & " " & msg
    frmServer.txtText.SelStart = Len(frmServer.txtText.Text)
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function
