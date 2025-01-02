Attribute VB_Name = "modMain"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal _
lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Sub Main()
    ' Check if the config file exists.
    ' Don't randomly exit, at least notify the user.
    
    If App.PrevInstance = True Then DestroyUpdater
    
    ClearServers
    
    
    ' Set the base value for a single %.
    ProgressP = 63.75 ' frmMain.picprogress.Width / 100
    
    If Not FileExist(App.Path & "\data\UpdateConfig.ini") Then
        'MsgBox "The configuration file appears to be missing." & vbNewLine & "Please check if it exists or re-install the application.", vbCritical, "Error"
        'DestroyUpdater
        'ChangeStatus "Comfiguration file is missing! Can't update.": Exit Sub
        ' Load the values we need into memory.
        
        UpdateURL = "http://127.0.0.1/Zelda/"
        'NewsURL = GetVar(App.Path & "\data\UpdateConfig.ini", "UPDATER", "NewsURL")
        'GameWebsite = GetVar(App.Path & "\data\UpdateConfig.ini", "UPDATER", "GameWebsite")
        GameName = "Zelda: The Golden Kingdom"
        ClientName = "Zelda - The Golden Kingdom.exe"
        serverIP = "127.0.0.1"
        
        PutVar App.Path & "\data\UpdateConfig.ini", "UPDATER", "UpdateURL", UpdateURL
        PutVar App.Path & "\data\UpdateConfig.ini", "Options", "Game_Name", GameName
        PutVar App.Path & "\data\UpdateConfig.ini", "UPDATER", "ClientName", ClientName
        PutVar App.Path & "\data\UpdateConfig.ini", "UPDATER", "ServerIP", serverIP
        
        ' Set the base value for a single %.
        ProgressP = 63.75 ' frmMain.picprogress.Width / 100
        
        Call PathCheck
        ' Load the main form
        Load frmMain
        'CheckVersion
        
    Else
        
        ' Load the values we need into memory.
        UpdateURL = GetVar(App.Path & "\data\UpdateConfig.ini", "UPDATER", "UpdateURL")
        'NewsURL = GetVar(App.Path & "\data\UpdateConfig.ini", "UPDATER", "NewsURL")
        'GameWebsite = GetVar(App.Path & "\data\UpdateConfig.ini", "UPDATER", "GameWebsite")
        GameName = GetVar(App.Path & "\data\UpdateConfig.ini", "Options", "Game_Name")
        ClientName = Trim(GetVar(App.Path & "\data\UpdateConfig.ini", "UPDATER", "ClientName"))
        serverIP = Trim(GetVar(App.Path & "\data\UpdateConfig.ini", "UPDATER", "ServerIP"))
        
        ' Set the base value for a single %.
        ProgressP = 63.75 ' frmMain.picprogress.Width / 100
        
        ' Load the main form
        'Load frmMain
        'CheckVersion
        
    End If
    
    Call PathCheck
    
    ' Load the main form
    Load frmMain
    
End Sub

Sub PathCheck()
    On Error Resume Next
    MkDir App.Path & "\data\"
End Sub

Public Sub DestroyUpdater()
    ' Delete all temporary doodles.
    If FileExist(App.Path & "\data\tmpUpdate.ini") Then Kill App.Path & "\data\tmpUpdate.ini"
    If FileExist(App.Path & "\data\status.ini") Then Kill App.Path & "\data\status.ini"
    
    ' End the program.
    frmMain.inetDownload.Cancel
    Unload frmMain
    End
End Sub

Public Sub ChangeStatus(ByVal NewStatus As String)
    frmMain.lblstatus.Caption = NewStatus
End Sub

Public Sub CheckVersion()
Dim Filename As String
    
    'On Error GoTo ErrorHandle
    Sleep 100
    If FileExist("launcher.update.exe") Then
    ChangeStatus "Updating Launcher.."
    DoEvents
    Sleep 300
        'MsgBox "launcher.update found"
        If App.EXEName = "launcher.update" Then
            'MsgBox "launcher.update is the app name. Updating!"
            ShellExecute 1, "Open", "taskkill /F /IM launcher.exe", "", 0&, 0
            'ShellExecute 1, "Open", "taskkill /F /IM " & Chr(34) & "El Reino Dorado Translated.exe" & Chr(34), "", 0&, 0
            ShellExecute 1, "Open", "taskkill /F /IM " & Chr(34) & ClientName & Chr(34), "", 0&, 0
            ChangeStatus "Replacing Old Launcher.."
            DoEvents
            Sleep 200
            Kill App.Path & "\launcher.exe"
            FileCopy App.Path & "\launcher.update.exe", App.Path & "\launcher.exe"
            Sleep 200
            Shell App.Path & "\launcher.exe", vbNormalFocus
            End
        Else
        ChangeStatus "Removing Old Launcher.."
        DoEvents
        Sleep 200
            Kill App.Path & "\launcher.update.exe"
        End If
        
    'replace myself with the new version?
    End If
    
    If App.PrevInstance = True Then End
    
    ' Enable our timeout timer, so it doesn't endlessly keep
    ' trying to connect.
    'frmMain.tmrTimeout.Enabled = True
    
    ' Change the status of our updater, and progress down to 0.
    'ChangeStatus "Connecting to update server..."
    SetProgress 25
    DoEvents
    Sleep 100
    ' Get the file which contains the info of updated files
        
    
    DownloadFile UpdateURL & "news.txt", App.Path & "\data\news.txt"
    
    DoEvents
    
    DownloadFile UpdateURL & "update.txt", App.Path & "\data\tmpUpdate.ini"
    Sleep 100
    DoEvents
    ' Done with the download, update the progress and continue!
    SetProgress 50
    ChangeStatus "Retrieving version information.."
    Sleep 100
    DoEvents
    ' read the version count
    
    If Not FileExist(App.Path & "\data\tmpUpdate.ini") Then
        SetProgress 100
        UpToDate = 1
        frmMain.tmrCheck.Enabled = False
        frmMain.tmrServerStatus.Enabled = False
        frmMain.tmrUpToDate.Enabled = False
        ResetClientConfig
        frmMain.inetDownload.Cancel
        frmMain.lblConnect.Visible = True
        ChangeStatus "Can't connect to update server: " & UpdateURL
        Exit Sub
    End If
    
    VersionCount = GetVar(App.Path & "\data\tmpUpdate.ini", "UPDATER", "Version")
    
    ' check if we've got a current client version saved
    If FileExist(App.Path & "\data\Config.ini") Then
        CurVersion = Val(GetVar(App.Path & "\data\Config.ini", "UPDATER", "Version"))
    Else
        CurVersion = 0
    End If
    
    SetProgress 100
    Sleep 100
    DoEvents
    ' Disable it, we have progress!
    frmMain.tmrTimeout.Enabled = False

    ' are we up to date?
    If CurVersion < VersionCount Then
        UpToDate = 0
        ChangeStatus "Your client is outdated!"
        VersionsToGo = VersionCount - CurVersion
        PercentToGo = 100 / (VersionsToGo * 2)
        Sleep 100
        DoEvents
        'If MsgBox("Your client is out of date! Would you like to update your client now?", vbYesNo, "Update Required!") = vbYes Then
            RunUpdates
        'Else
        '    DestroyUpdater
        'End If
    Else
        UpToDate = 1
        If SelectedServer > 0 Then
            If Server(SelectedServer).CurrentPlayers > 0 Then
                ChangeStatus "Your client is up to date! Server: " & Server(SelectedServer).Name & " - " & " Players Online: " & Server(SelectedServer).CurrentPlayers
            Else
                ChangeStatus "Your client is up to date! Server: " & Server(SelectedServer).Name
            End If
            frmMain.lblConnect.Visible = True
        Else
            ChangeStatus "Your client is up to date!"
            frmMain.lblConnect.Visible = True
        End If
        SetProgress 100
        ' Load a GUI image, if it does not exist.. Exit out of the program.
        'Form_LoadPicture (App.Path & "\data\graphics\gui\updater\launch.jpg")
        frmMain.lblConnect.Visible = True
    End If
    
    
    Exit Sub
ErrorHandle:
    MsgBox "An unexpected error has occured: " & Err.Description & vbNewLine & vbNewLine & "It is likely that your configuration is incorrect.", vbCritical
    Err.Clear
    DestroyUpdater
End Sub

Public Sub RunUpdates()
Dim Filename As String
Dim i As Long
Dim UpdateID As Long
Dim CurProgress As Long
    
    On Error GoTo ErrorHandle
    
    If CurVersion = 0 Then CurVersion = 1 Else CurVersion = CurVersion + 1
    UpdateID = 0
    CurProgress = 0
    ' loop around, download and unrar each update
    Sleep 100
    DoEvents
    ChangeStatus "Checking components.."
    If Not FileExist(App.Path & "\unrar.dll") Then
        DownloadFile UpdateURL & "/unrar.dll", App.Path & "\unrar.dll"
    End If
    Sleep 100
    DoEvents
    
    For i = CurVersion To VersionCount
        ' Increase Update ID by 1
        UpdateID = UpdateID + 1
        
        ' let them know we're actually doing something..
        CurProgress = CurProgress + PercentToGo
        ChangeStatus "Downloading update " & Str(UpdateID) & "/" & Trim$(Str(VersionsToGo)) & ".."
        SetProgress CurProgress
        Sleep 100
        DoEvents
        ' Download time!
        Filename = "version" & Trim(Str(i)) & ".rar"
        DownloadFile UpdateURL & "/" & Filename, App.Path & "\" & Filename
        Sleep 100
        DoEvents
        ' Done downloading? Awesome.. Time to change the status
        CurProgress = CurProgress + PercentToGo
        ChangeStatus "Unpacking update " & Str(UpdateID) & "/" & Trim$(Str(VersionsToGo)) & "..."
        Sleep 100
        DoEvents
        SetProgress CurProgress
        Sleep 100
        DoEvents
        ' Extract date from the update file.
        RARExecute OP_EXTRACT, Filename
        
        ' Delete the update file.
        Kill App.Path & "\" & Filename
        
        If FileExist(App.Path & "\launcher.update.exe") Then
            PutVar App.Path & "\data\Config.ini", "UPDATER", "Version", Str(i)
            DoEvents
            Shell App.Path & "\launcher.update.exe"
            DoEvents
            DestroyUpdater
            End
    End If
        
    Next
    
    If FileExist(App.Path & "\temp.txt") Then Kill App.Path & "\temp.txt"
    
    ' Update the version of the client.
    PutVar App.Path & "\data\Config.ini", "UPDATER", "Version", Str(i)
    
     ' Load the Launch Backdrop.
    'Form_LoadPicture (App.Path & "\data\graphics\gui\updater\launch.jpg")
    
    ' Set the Update variable to 1, to prevent running this sub again.
    UpToDate = 1
    'MsgBox "copy launcher.update into folder now!"
    'check for a launcher update..
    ChangeStatus "Updating Launcher."
    DoEvents
    Sleep 100
    'If FileExist(App.Path & "\launcher.update.exe") Then
    '    Shell App.Path & "\launcher.update.exe"
    '    End
    'End If
    SetProgress 95
    
    ResetClientConfig
    
    ' Done? Niiiice..
    SetProgress 100 'Just to be sure, sometimes it misses ~1% due to the lack of decimals.
    ChangeStatus "Your client is up to date!"
    Sleep 100
    DoEvents
    
    Exit Sub
    
ErrorHandle:
    MsgBox "An error has occured while extracting the update(s): " & Err.Description & vbNewLine & vbNewLine & "Please relay this message to an administrator.", vbCritical
    Err.Clear
    DestroyUpdater
End Sub

Public Sub ResetClientConfig()
ChangeStatus "Resetting Config Settings.."
    
    PutVar App.Path & "\data\Config.ini", "Options", "SafeMode", "1"
    PutVar App.Path & "\data\Config.ini", "Options", "Game_Name", "The Legend Of Zelda: The Golden Kingdom"
    PutVar App.Path & "\data\Config.ini", "Options", "SavePass", "1"
    PutVar App.Path & "\data\Config.ini", "Options", "Port", "4000"
    PutVar App.Path & "\data\Config.ini", "Options", "Music", "1"
    PutVar App.Path & "\data\Config.ini", "Options", "Sound", "1"
    PutVar App.Path & "\data\Config.ini", "Options", "Debug", "1"
    PutVar App.Path & "\data\Config.ini", "Options", "Names", "1"
    PutVar App.Path & "\data\Config.ini", "Options", "Level", "1"
    PutVar App.Path & "\data\Config.ini", "Options", "WASD", "1"
    PutVar App.Path & "\data\Config.ini", "Options", "MiniMapBltElse", "1"
    PutVar App.Path & "\data\Config.ini", "Options", "Chat", "1"
    PutVar App.Path & "\data\Config.ini", "Options", "DefaultVolume", "50"
    PutVar App.Path & "\data\Config.ini", "Options", "MiniMap", "1"
    PutVar App.Path & "\data\Config.ini", "Options", "IP", serverIP
    PutVar App.Path & "\data\Config.ini", "Options", "ChatToScreen", "2"
    PutVar App.Path & "\data\Config.ini", "Options", "MappingMode", "0"
    PutVar App.Path & "\data\Config.ini", "Options", "RequireLauncher", "1"
    
    PutVar App.Path & "\data\Config.ini", "ChatOptions", "1", "1"
    PutVar App.Path & "\data\Config.ini", "ChatOptions", "2", "1"
    PutVar App.Path & "\data\Config.ini", "ChatOptions", "3", "1"
    PutVar App.Path & "\data\Config.ini", "ChatOptions", "4", "1"
    PutVar App.Path & "\data\Config.ini", "ChatOptions", "5", "1"
    PutVar App.Path & "\data\Config.ini", "ChatOptions", "6", "1"
    
ChangeStatus "Client Config Settings Reset!"
End Sub

Public Sub SetProgress(ByVal Percent As Long)
    If Percent = 0 Then
        frmMain.picprogress.Width = 0
    ElseIf Percent = 100 Then
        frmMain.picprogress.Width = (ProgressP * Percent)
    End If
End Sub

Public Sub DownloadFile(ByVal URL As String, ByVal Filename As String)
    Dim fileBytes() As Byte
    Dim fileNum As Integer
    
    On Error GoTo DownloadError
        
    If frmMain.inetDownload.StillExecuting Then
        Dim i As Integer
        Do While i < 500
            Sleep 10
            i = i + 1
            DoEvents
            If frmMain.inetDownload.StillExecuting = False Then Exit Do
        Loop
    End If
    
    If frmMain.inetDownload.StillExecuting Then Exit Sub
    
    'ChangeStatus "Retrieving " & URL
    
    ' download data to byte array
    fileBytes() = frmMain.inetDownload.OpenURL(URL, icByteArray)

    fileNum = FreeFile
    Open Filename For Binary Access Write As #fileNum
        ' dump the byte array as binary
        Put #fileNum, , fileBytes()
    Close #fileNum
    
    Exit Sub
    
DownloadError:
    'MsgBox Err.Description & vbCrLf & URL
End Sub

Public Sub Form_LoadPicture(ByVal Filename As String)
    If FileExist(Filename) Then
            frmMain.Picture = LoadPicture(Filename)
        Else
            DestroyUpdater
        End If
End Sub

Sub ReadServerFile()
Dim File As String, header As String
Dim i As Integer
Dim NumServers As Integer
File = App.Path & "\data\status.ini"

If Not FileExist(File) Then Exit Sub

frmSelect.lstServers.Clear

header = "Servers"
NumServers = Val(GetVar(File, header, "NumServers"))

If NumServers > UBound(Server) Then NumServers = UBound(Server)

For i = 1 To NumServers

header = "Server" & i

    Server(i).CurrentPlayers = Val(GetVar(File, header, "Players"))
    Server(i).MaxPlayers = Val(GetVar(File, header, "MaxPlayers"))
    'PutVar File, Header, "PvPOnly", "0"
    Server(i).Name = GetVar(File, header, "Name")
    Server(i).Port = Val(GetVar(File, header, "Port"))
    Server(i).Online = Val(GetVar(File, header, "Online"))

If frmSelect Is Nothing Then Load frmSelect
    With frmSelect.lstServers
        If Server(i).Online = True Then
            .AddItem Server(i).Name & " - Players: " & Server(i).CurrentPlayers & "/" & Server(i).MaxPlayers
        End If
    End With

Next i

'If CheckServerFull(SelectedServer) Then CheckServerFull


End Sub

Public Function CheckServerFull(Optional index As Integer = 0) As Boolean
Dim i As Integer
Dim anyOnline As Boolean
On Error Resume Next


If UBound(Server) > 0 Then frmMain.lblServer.Visible = True

For i = 1 To UBound(Server)
    If Server(i).Online = True Then anyOnline = True
Next i

If index = 0 Then

    For i = 1 To UBound(Server)
            If (Server(i).CurrentPlayers > Server(SelectedServer).CurrentPlayers) And (Server(i).CurrentPlayers < Server(SelectedServer).MaxPlayers) Then SelectedServer = i
    Next

    For i = 1 To UBound(Server)
        If SelectedServer = 0 Then SelectedServer = i
        
        If Server(SelectedServer).CurrentPlayers >= Server(SelectedServer).MaxPlayers & Server(SelectedServer).MaxPlayers > 0 Then
            If Server(i + 1).CurrentPlayers < Server(i + 1).MaxPlayers Then
                If Server(i + 1).Online = True Then SelectedServer = i + 1
                frmMain.lblServer.Visible = True
            End If
        Else
            CheckServerFull = False
            Exit Function
        End If
    
    Next
    If anyOnline = False Then
        MsgBox "I'm unable to retrieve server status information!" & vbNewLine & _
        "So we'll just try the default port and hope for the best!", vbInformation, "Unable to Get Status"
        SelectedServer = 1
        Server(SelectedServer).Port = 4000
        Server(SelectedServer).MaxPlayers = 40
        Server(SelectedServer).Online = True
        Server(SelectedServer).CurrentPlayers = 0
    Else
        MsgBox "It appears that all servers are full at the moment! :(", vbInformation, "Sorry! All Full?!"
        CheckServerFull = True
    End If
    Exit Function
Else

    If Server(index).CurrentPlayers < Server(index).MaxPlayers Then CheckServerFull = False: Exit Function

End If

CheckServerFull = True

End Function

Sub WriteClientInfo()
Dim File As String

If SelectedServer = 0 Then SelectedServer = 1

CheckServerFull (SelectedServer)

If Server(SelectedServer).Port = 0 Then Exit Sub

File = App.Path & "\Data\config.ini"

'MsgBox "Server chosen: " & CStr(Server(SelectedServer).Name)

PutVar File, "Options", "Port", CStr(Server(SelectedServer).Port)
PutVar File, "Options", "RequireLauncher", "0"

DoEvents

End Sub

Sub ClearServers()

    Dim i As Integer
    For i = 1 To UBound(Server)
        Server(i).CurrentPlayers = 0
        Server(i).MaxPlayers = 0
        Server(i).Name = ""
        Server(i).Online = False
        Server(i).Port = 0
    Next i

End Sub
