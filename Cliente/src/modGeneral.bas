Attribute VB_Name = "modGeneral"
Option Explicit

Global wtf As Boolean

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long

Dim strMissing As String

'For Clear functions
Public Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)
Public DX7 As New DirectX7  ' Master Object, early binding

Public Sub Main()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    On Error Resume Next
    
    If Len(App.Path & "\data\graphics\GUI\main\buttons\skills_click.jpg") >= 230 Then MsgBox "The folder that this game is in could cause issues due to length of its path." & vbNewLine & _
    "Please place the game in a location with a shorter path!", vbCritical, "Oh no!": End

    If FileExist("launcher.update.exe") = True Then
        'msgbox "launcher update found."
    'replace old launcher with the new one and then run it!
    'verify all names are correct before pushing update out
        If ShellExecute(1, "Open", "taskkill /F /IM launcher.exe", "", 0&, 0) > 0 Then
            'msgbox "Killing launcher.exe"
            Sleep 100
            DoEvents
            Kill App.Path & "\launcher.exe"
            Sleep 100
            If FileExist("launcher.exe") Then
                'msgbox "launcher.exe still exists.."
                ShellExecute 1, "Open", "taskkill /F /IM launcher.exe", "", 0&, 0
                Sleep 100
                DoEvents
                Kill App.Path & "\launcher.exe"
                Sleep 100
            End If
            'msgbox "Replacing launcher.exe"
            FileCopy App.Path & "\launcher.update.exe", App.Path & "\launcher.exe"
            If FileExist("launcher.exe") = True Then
                'msgbox "Starting new launcher.exe"
                Shell App.Path & "\launcher.exe", vbNormalFocus
                Sleep 50
                End
            Else
            'msgbox "Launcher didn't copy? Try again."
                FileCopy App.Path & "\launcher.update.exe", App.Path & "\launcher.exe"
                Sleep 50
                DoEvents
                Shell App.Path & "\launcher.exe", vbNormalFocus
                End
            End If
        End If
    
End If

    LangTo = "en"
    LangFrom = "es"
    strTransPath = App.Path & "\" & LangTo & ".dat"
    strOrigPath = App.Path & "\" & LangFrom & "-" & LangTo & ".dat"

    'frmTransLog.Show
    'frmTransLog.txtLog.text = "GTranslate and modTranslate by Dragoon/LostSoulFly!"
    'DoEvents

    ' set loading screen
    loadGUI True
    frmLoad.Visible = True

    ' load options
    Call SetStatus("Cargando Opciones\Traducción...")
    LoadOptions
    
    ' load main menu
    Call SetStatus("Cargando Menú...")
    Load frmMenu
    
    ' load gui
    Call SetStatus("Cargando Interface...")
    loadGUI
    
    Dim strsplit() As String, i As Integer
    strsplit = Split(Command, " ")
    If (UBound(strsplit)) Mod 2 > 0 Then
        For i = 0 To UBound(strsplit) Step 2
            Select Case LCase$(strsplit(i))
            
            Case Is = "-user"
                Options.Username = Trim$(strsplit(i + 1))
                
            Case Is = "-pass"
                Options.Password = Trim$(strsplit(i + 1))
            
            Case Is = "-server"
                'anything above 0 will enable it.
                Options.ip = Trim$(strsplit(i + 1))
                frmMain.Caption = Options.Game_Name & " - " & Options.ip
                
            Case Is = "-port"
                'anything above 0 will enable it.
                Options.port = Val(Trim$(strsplit(i + 1)))
                'MsgBox CStr(Options.port)
            Case Is = "-auto"
                AutoLogin = Val(Trim$(strsplit(i + 1)))
                
            Case Is = "-wtf"
                wtf = True
            
            End Select
        Next
    End If
    
    If frmMain.Caption = "" Then frmMain.Caption = Options.Game_Name
    
    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\data\", "graphics"
    ChkDir App.Path & "\data\graphics\", "animations"
    ChkDir App.Path & "\data\graphics\", "characters"
    ChkDir App.Path & "\data\graphics\", "items"
    ChkDir App.Path & "\data\graphics\", "paperdolls"
    ChkDir App.Path & "\data\graphics\", "resources"
    ChkDir App.Path & "\data\graphics\", "spellicons"
    ChkDir App.Path & "\data\graphics\", "tilesets"
    ChkDir App.Path & "\data\graphics\", "faces"
    ChkDir App.Path & "\data\graphics\", "gui"
    ChkDir App.Path & "\data\graphics\gui\", "menu"
    ChkDir App.Path & "\data\graphics\gui\", "main"
    ChkDir App.Path & "\data\graphics\gui\menu\", "buttons"
    ChkDir App.Path & "\data\graphics\gui\main\", "buttons"
    ChkDir App.Path & "\data\graphics\gui\main\", "bars"
    ChkDir App.Path & "\data\", "logs"
    ChkDir App.Path & "\data\", "maps"
    ChkDir App.Path & "\data\", "music"
    ChkDir App.Path & "\data\", "sound"
        
    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34) ' "
    
    ' initialize DirectX
    If Not InitDirectDraw Then
        MsgBox "Error initiating DirectX7 - DirectDraw."
        DestroyGame
        End
    End If
    
    ' randomize rnd's seed
    Randomize
    Call SetStatus("Iniciando ajustes de TCP...")
    Call TcpInit
    Call InitMessages
    Call SetStatus("Iniciando DirectX...")
    
    ' DX7 Master Object is already created, early binding
    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSpellIcons
    Call CheckFaces
    Call CheckProjectiles
    
    ' temp set music/sound vars
    Music_On = True
    Sound_On = True
    
    ' load music/sound engine
    InitSound
    InitMusic
    
    ' check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then PlayMidi Trim$(Options.MenuMusic)
    
    ' Reset values
    Ping = -1
    
    'Load frmMainMenu
    Load frmMenu
    
    ' cache the buttons then reset & render them
    Call SetStatus("Cargando Botones...")
    cacheButtons
    resetButtons_Menu
    
    ' we can now see it
    frmMenu.Visible = True
    
    ' hide all pics
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' set the paperdoll order
    ReDim PaperdollOrder(1 To Equipment.Equipment_Count - 1) As Long
    PaperdollOrder(1) = Equipment.Armor
    PaperdollOrder(2) = Equipment.Helmet
    PaperdollOrder(3) = Equipment.Shield
    PaperdollOrder(4) = Equipment.Weapon
    
    ' hide the load form
    frmLoad.Visible = False
    
    Call InitChatRooms
    
    If AutoLogin = True Then
        If isLoginLegal(frmMenu.txtLUser.text, frmMenu.txtLPass.text) Then
            Call MenuState(MENU_STATE_LOGIN)
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetRealTickCount() As Long
    If GetTickCount < 0 Then
        GetRealTickCount = GetTickCount + MAX_LONG
    Else
        GetRealTickCount = GetTickCount
    End If
End Function

Public Sub loadGUI(Optional ByVal loadingScreen As Boolean = False)
Dim i As Long

    ' if we can't find the interface
    On Error GoTo errorhandler
    
    ' loading screen
    If loadingScreen Then
        frmLoad.Picture = LoadPicture(App.Path & "\data\graphics\gui\menu\loading.jpg")
        Exit Sub
    End If

    ' menu
    frmMenu.Picture = LoadPicture(App.Path & "\data\graphics\gui\menu\background.jpg")
    frmMenu.picMain.Picture = LoadPicture(App.Path & "\data\graphics\gui\menu\main.jpg")
    frmMenu.picLogin.Picture = LoadPicture(App.Path & "\data\graphics\gui\menu\login.jpg")
    frmMenu.picRegister.Picture = LoadPicture(App.Path & "\data\graphics\gui\menu\register.jpg")
    frmMenu.picCredits.Picture = LoadPicture(App.Path & "\data\graphics\gui\menu\credits.jpg")
    frmMenu.picCharacter.Picture = LoadPicture(App.Path & "\data\graphics\gui\menu\character.jpg")
    ' main
    frmMain.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\main.jpg")
    frmMain.picInventory.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\inventory.jpg")
    frmMain.picCharacter.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\character.jpg")
    frmMain.picSpells.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\skills.jpg")
    frmMain.picOptions.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\options.jpg")
    frmMain.picGuild.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\guild.jpg")
    frmMain.picParty.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\party.jpg")
    frmMain.picItemDesc.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\description_item.jpg")
    frmMain.picSpellDesc.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\description_spell.jpg")
    frmMain.picTempInv.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\dragbox.jpg")
    frmMain.picTempBank.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\dragbox.jpg")
    frmMain.picTempSpell.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\dragbox.jpg")
    frmMain.picShop.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\shop.jpg")
    frmMain.picBank.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\bank.jpg")
    frmMain.picTrade.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\trade.jpg")
    frmMain.picHotbar.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\hotbar.jpg")
    frmMain.picSpeechClose.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\speechbutton.jpg")
    frmMain.picQuestLog.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\questlog.jpg")
    frmMain.picQuestDialogue.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\questdialogue.jpg")
    frmMain.picPets.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\pets.jpg")
    frmMain.picPetStats.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\petstats.jpg")
    frmMain.picCurrency.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\textboxes.jpg")
    frmMain.picDialogue.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\textboxes.jpg")
    frmMain.picSpeech.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\textboxes.jpg")
    
    ' main - bars
    frmMain.imgMPBar.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\bars\spirit.jpg")
    frmMain.imgEXPBar.Picture = LoadPicture(App.Path & "\data\graphics\gui\main\bars\experience.jpg")
    ' main - party bars
    For i = 1 To MAX_PARTY_MEMBERS
        frmMain.imgPartyHealth(i).Picture = LoadPicture(App.Path & "\data\graphics\gui\main\bars\party_health.jpg")
        frmMain.imgPartySpirit(i).Picture = LoadPicture(App.Path & "\data\graphics\gui\main\bars\party_spirit.jpg")
    Next
    
    ' store the bar widths for calculations
    HPBar_Width = frmMain.imgHPBar.Width
    SPRBar_Width = frmMain.imgMPBar.Width
    EXPBar_Width = frmMain.imgEXPBar.Width
    ' party
    Party_HPWidth = frmMain.imgPartyHealth(1).Width
    Party_SPRWidth = frmMain.imgPartySpirit(1).Width
    
    Exit Sub
    
' let them know we can't load the GUI
errorhandler:
    If Len(strMissing) = 0 Then
        MsgBox "Cannot find one or more interface images." & vbNewLine & "If they exist then you have not extracted the project properly." & vbNewLine & "Please follow the installation instructions fully.", vbCritical
    Else
        MsgBox "Unable to locate files: " & strMissing & "", vbCritical, "Files missing!"
    End If
    DestroyGame
    Exit Sub
End Sub

Public Sub MenuState(ByVal State As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmLoad.Visible = True

    Select Case State
        Case MENU_STATE_ADDCHAR
            frmMain.txtMyChat.Locked = True
            frmMain.picTutorial.Visible = True
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Conectado, enviando información de personaje...")

                If frmMenu.optMale.value Then
                    Call SendAddChar(frmMenu.txtCName, SEX_MALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite)
                Else
                    Call SendAddChar(frmMenu.txtCName, SEX_FEMALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite)
                End If
            End If
            
        Case MENU_STATE_NEWACCOUNT
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Conectado, enviando información de la nueva cuenta...")
                Call SendNewAccount(frmMenu.txtRUser.text, frmMenu.txtRPass.text)
            End If

        Case MENU_STATE_LOGIN
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Conectado, recibiendo datos...")
                Call SendLogin(frmMenu.txtLUser.text, frmMenu.txtLPass.text)
                Exit Sub
            End If
    End Select

    If frmLoad.Visible Then
        If Not IsConnected Then
            frmMenu.Visible = True
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False
            frmLoad.Visible = False
            Call MsgBox("The server seems to be offline. Please try again later.", vbOKOnly, Options.Game_Name)
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MenuState", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub logoutGame()
Dim Buffer As clsBuffer, i As Long

    DoEvents
    isLogging = True
    InGame = False
    Set Buffer = New clsBuffer
    Buffer.WriteLong CQuit
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    Call DestroyTCP
    
    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    ' destroy temp values
    DragInvSlotNum = 0
    InvX = 0
    InvY = 0
    EqX = 0
    EqY = 0
    SpellX = 0
    SpellY = 0
    LastItemDesc = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    tmpCurrencyItem = 0
    
    ' unload editors
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Spell
    
    ' hide main form stuffs
    frmMenu.picMain.Visible = True
    frmMain.txtChat.text = vbNullString
    frmMain.txtMyChat.text = vbNullString
    frmMain.picCurrency.Visible = False
    frmMain.picDialogue.Visible = False
    frmMain.picInventory.Visible = False
    frmMain.picTrade.Visible = False
    frmMain.picSpells.Visible = False
    frmMain.picCharacter.Visible = False
    frmMain.picOptions.Visible = False
    frmMain.picParty.Visible = False
    frmMain.picAdmin.Visible = False
    frmMain.picBank.Visible = False
End Sub

Sub GameInit()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EnteringGame = True
    frmMenu.Visible = False
    EnteringGame = False
    
    ' bring all the main gui components to the front
    'frmMain.picShop.ZOrder (0)
    'frmMain.picBank.ZOrder (0)
    frmMain.picTrade.ZOrder (0)
    
    ' hide gui
    InBank = False
    InShop = False
    InTrade = False
    
    ' Set font
    'Call SetFont(FONT_NAME, FONT_SIZE)
    frmMain.Font = "Georgia"
    frmMain.FontSize = 14
    
    ' show the main form
    frmLoad.Visible = False
    frmMain.Show
    
    ' Set the focus
    frmMain.picScreen.Visible = True
    
    'Call SetFocusOnChat

    
    
    If Options.WASD = 1 Then
        ChatFocus = False
        frmMain.picScreen.SetFocus
    Else
        ChatFocus = True
    End If
    
    
    ' Blt inv
    BltInventory
    
    ' blt hotbar
    BltHotbar
    
    ' get ping
    GetPing
    DrawPing
    
    ' set values for amdin panel
    frmMain.scrlAItem.Max = MAX_ITEMS
    frmMain.scrlAItem.value = 1
    
    'stop the song playing
    StopMidi
    
    InitVideo
       
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameInit", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyGame()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call saveLang(strTransPath, langCol, True)
    Call saveLang(strOrigPath, origCol, True)
    
    SaveOptions
    
    ' break out of GameLoop
    InGame = False
    Call DestroyTCP
    
    'destroy objects in reverse order
    Call DestroyDirectDraw

    ' destory DirectX7 master object
    If Not DX7 Is Nothing Then
        Set DX7 = Nothing
    End If

    Call UnloadAllForms
    End
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "destroyGame", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadAllForms()
Dim frm As Form

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For Each frm In VB.Forms
        Unload frm
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadAllForms", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetStatus(ByVal Caption As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    Caption = GetTranslation(Caption)
    frmLoad.lblStatus.Caption = Caption
    DoEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, msg As String, NewLine As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NewLine Then
        Txt.text = Txt.text + msg + vbCrLf
    Else
        Txt.text = Txt.text + msg
    End If

    Txt.SelStart = Len(Txt.text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetFocusOnChat()

    On Error Resume Next 'prevent RTE5, no way to handle error
If Options.WASD = 0 Then frmMain.txtMyChat.SetFocus Else frmMain.picScreen.SetFocus

End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Rand = Int((High - Low + 1) * Rnd) + Low
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Rand", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim GlobalX As Long
Dim GlobalY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GlobalX = PB.Left
    GlobalY = PB.Top

    If Button = 1 Then
        PB.Left = GlobalX + X - SOffsetX
        PB.Top = GlobalY + Y - SOffsetY
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MovePicture", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isLoginLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent high ascii chars
    For i = 1 To Len(sInput)

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, Options.Game_Name)
            Exit Function
        End If

    Next

    isStringLegal = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isStringLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' ####################
' ## Buttons - Menu ##
' ####################
Public Sub cacheButtons()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' menu - login
    With MenuButton(1)
        .Filename = "login"
        .State = 0 ' normal
    End With
    
    ' menu - register
    With MenuButton(2)
        .Filename = "register"
        .State = 0 ' normal
    End With
    
    ' menu - credits
    With MenuButton(3)
        .Filename = "credits"
        .State = 0 ' normal
    End With
    
    ' menu - exit
    With MenuButton(4)
        .Filename = "exit"
        .State = 0 ' normal
    End With
    
    ' main - inv
    With MainButton(1)
        .Filename = "inv"
        .State = 0 ' normal
    End With
    
    ' main - skills
    With MainButton(2)
        .Filename = "skills"
        .State = 0 ' normal
    End With
    
    ' main - char
    With MainButton(3)
        .Filename = "char"
        .State = 0 ' normal
    End With
    
    ' main - opt
    With MainButton(4)
        .Filename = "opt"
        .State = 0 ' normal
    End With
    
    ' main - trade
    With MainButton(5)
        .Filename = "trade"
        .State = 0 ' normal
    End With
    
    ' main - party
    With MainButton(6)
        .Filename = "party"
        .State = 0 ' normal
    End With
    
    'Alatar v1.2
    ' main - quest
    With MainButton(7)
        .Filename = "quest"
        .State = 0 ' normal
    End With
    '/Alatar v1.2
    
    With MainButton(8)
        .Filename = "pet"
        .State = 0 ' normal
    End With
    
    With MainButton(9)
        .Filename = "map"
        .State = 0 ' normal
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cacheButtons", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' menu specific buttons
Public Sub resetButtons_Menu(Optional ByVal exceptionNum As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_MENUBUTTONS
        ' only change if different and not exception
        If Not MenuButton(i).State = 0 And Not i = exceptionNum Then
            ' reset state and render
            MenuButton(i).State = 0 'normal
            renderButton_Menu i
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Menu = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Menu(ByVal buttonNum As Long)
Dim bSuffix As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' get the suffix
    Select Case MenuButton(buttonNum).State
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMenu.imgButton(buttonNum).Picture = LoadPicture(App.Path & MENUBUTTON_PATH & MenuButton(buttonNum).Filename & bSuffix & ".jpg")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Menu(ByVal buttonNum As Long, ByVal bState As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MenuButton(buttonNum).State = bState Then Exit Sub
        ' change and render
        MenuButton(buttonNum).State = bState
        renderButton_Menu buttonNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' main specific buttons
Public Sub resetButtons_Main(Optional ByVal exceptionNum As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_MAINBUTTONS
        ' only change if different and not exception
        If Not MainButton(i).State = 0 And Not i = exceptionNum Then
            ' reset state and render
            MainButton(i).State = 0 'normal
            renderButton_Main i
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Main = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Main(ByVal buttonNum As Long)
Dim bSuffix As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' get the suffix
    Select Case MainButton(buttonNum).State
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMain.imgButton(buttonNum).Picture = LoadPicture(App.Path & MAINBUTTON_PATH & MainButton(buttonNum).Filename & bSuffix & ".jpg")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Main(ByVal buttonNum As Long, ByVal bState As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MainButton(buttonNum).State = bState Then Exit Sub
        ' change and render
        MainButton(buttonNum).State = bState
        renderButton_Main buttonNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PopulateLists()
Dim strLoad As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Cache music list
    strLoad = dir(App.Path & MUSIC_PATH & "*.mid")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop
    
    strLoad = dir(App.Path & MUSIC_PATH & "*.mp3")
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop
    
    ' Cache sound list
    strLoad = dir(App.Path & SOUND_PATH & "*.wav")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To i) As String
        soundCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateGuildData()
Dim i As Long, GuildRank As String, RankNum As Long
With frmMain
Dim index As Long

If Not Player(MyIndex).GuildName = vbNullString Then
Call GuildSave(4, MyIndex)
.lstGuildMembers.Visible = True
.lblGuildLeave.Visible = True
.lblGuildC.Visible = False
.frmGuildC.Visible = False
.lblGuildInv.Visible = True
.lblGuildKick.Visible = True
.lblGuildAdminPanel.Visible = True
If Player(MyIndex).GuildMemberId > 0 Then
If GuildData.Guild_Members(Player(MyIndex).GuildMemberId).Rank > 5 Then
frmMain.lblGuildDisband.Visible = True
frmMain.lblGuildTransfer.Visible = True
Else
frmMain.lblGuildDisband.Visible = False
frmMain.lblGuildTransfer.Visible = False
End If
.lblGuildFounder.Visible = False
End If
.lstGuildMembers.Clear
Else
.lstGuildMembers.Visible = False
.lblGuildInv.Visible = False
.lblGuildKick.Visible = False
.lblGuildLeave.Visible = False
.lblGuildC.Visible = True
.frmGuildC.Visible = False
.lblGuildDisband.Visible = False
.lblGuildFounder.Visible = False
.lblGuildAdminPanel.Visible = False
.lblGuildTransfer.Visible = False
End If
For i = 1 To MAX_GUILD_MEMBERS
If Not Player(MyIndex).GuildName = vbNullString Then
.lblGuild.Caption = "Clan: " & Player(MyIndex).GuildName
.lblGuildInv.Caption = GetTranslation("Invitar al Clan")
.lblGuildKick.Caption = GetTranslation("Expulsar del Clan")
.lblGuildLeave.Caption = GetTranslation("Abandonar Clan")
.lblGuildDisband.Caption = GetTranslation("Deshacer Clan")
.lblGuildFounder.Caption = GetTranslation("Hacer Fundador")
.lblGuildAdminPanel.Caption = GetTranslation("Administrar Clan")
.lblGuildTransfer.Caption = GetTranslation("Transferir Fundador")
Else
.lblGuild.Caption = GetTranslation("No estás en un Clan")
.lblGuildC.Caption = GetTranslation("Crear Clan")
.lblGuildCAccept.Caption = GetTranslation("Crear")
.lblGuildCCancel.Caption = GetTranslation("Cancelar")
.frmGuildC.Caption = GetTranslation("Creación del Clan")
.txtGuildC.text = ""
.picGuildInvitation.Visible = False
.lblGuildInvitation.Caption = GetTranslation("Invitación al Clan")
.lblGuildAcceptInvitation.Caption = GetTranslation("Aceptar")
.lblGuildDeclineInvitation.Caption = GetTranslation("Rechazar")
End If
If Not Trim$(GuildData.Guild_Members(i).User_Name) = vbNullString Then
RankNum = GuildData.Guild_Members(i).Rank
GuildRank = GuildData.Guild_Ranks(RankNum).Name
.lstGuildMembers.AddItem ("[" & GuildRank & "] " & GuildData.Guild_Members(i).User_Name)
.lstGuildMembers.ListIndex = 0
End If
Next i
End With
End Sub
