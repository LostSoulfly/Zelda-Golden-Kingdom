VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LulzSrv"
   ClientHeight    =   5340
   ClientLeft      =   8235
   ClientTop       =   6120
   ClientWidth     =   8370
   ForeColor       =   &H80000008&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Select an Image"
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdSave 
         Caption         =   "Use Selected Image"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   2415
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "Application Dir"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   855
      End
      Begin VB.FileListBox File1 
         Height          =   2625
         Left            =   2640
         Pattern         =   "*.jpg;*.png;*.gif;*.jpeg;*.bmp"
         TabIndex        =   10
         Top             =   240
         Width           =   4335
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Timer tmrSockets 
      Interval        =   15000
      Left            =   7560
      Top             =   600
   End
   Begin MSWinsockLib.Winsock srvIn 
      Index           =   0
      Left            =   7560
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7455
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   4920
         Width           =   3255
      End
      Begin VB.CommandButton cmdRestart 
         Caption         =   "Restart"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   4600
         Width           =   1095
      End
      Begin VB.CheckBox chkEvents 
         Caption         =   "LogEvents"
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   4860
         Width           =   1215
      End
      Begin VB.CommandButton cmdImage 
         Caption         =   "Change Image"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   4600
         Width           =   1215
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   3360
         TabIndex        =   3
         Text            =   "80"
         Top             =   4600
         Width           =   615
      End
      Begin VB.CheckBox chkSave 
         Caption         =   "Save Settings"
         Height          =   255
         Left            =   4560
         TabIndex        =   6
         Top             =   4860
         Width           =   1335
      End
      Begin VB.CheckBox chkLog 
         Caption         =   "Log URLs"
         Height          =   255
         Left            =   6000
         TabIndex        =   4
         Top             =   4600
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstProxyLog 
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Time"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Event"
            Object.Width           =   8291
         EndProperty
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   210
         Left            =   3000
         TabIndex        =   13
         Top             =   4635
         Width           =   495
      End
   End
   Begin VB.Image imgChocobo 
      Height          =   2760
      Left            =   8160
      Picture         =   "Form1.frx":08CA
      Top             =   1800
      Visible         =   0   'False
      Width           =   2445
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTotal As Integer
Dim strVersion As String
Dim intSeconds As Integer
Dim strPicture As String
Dim DelChocobo As Boolean
Dim IndexSrc As String
Dim ErrorSrc As String


Private Sub chkEvents_Click()
SaveSettings
End Sub

Private Sub chkLog_Click()
SaveSettings
End Sub

Private Sub chkSave_Click()
If srvIn(0).State = sckListening Then Call SaveSettings
If chkSave.Value = vbChecked Then chkSave.Visible = False
End Sub

Private Sub cmdCancel_Click()
Frame2.Visible = False
End Sub

Private Sub cmdDefault_Click()
Dir1.Path = App.Path
End Sub

Private Sub cmdImage_Click()
If Frame2.Visible = True Then Frame2.Visible = False Else Frame2.Visible = True
End Sub

Private Sub cmdRestart_Click()
ShutServer
Call LoadSettings
InitServer
lstProxyLog.ListItems.Clear
AddLog "System", "Server restarted on Port: " & ProxyListenPort
End Sub
    
Private Sub cmdSave_Click()
If Dir(File1.Path & "\" & File1.FileName) <> "" Then
    strPicture = File1.Path & "\" & File1.FileName
    Frame2.Visible = False
    SaveSettings
Else
    MsgBox "I can't access that file or it doesn't exist. Please try again!", vbCritical
End If
    
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_Click()
File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
'strPicture = File1.Path & "\" & File1.FileName
End Sub

Private Sub File1_DblClick()
cmdSave_Click
End Sub

    Private Sub Form_Load()
        ' Set form dimensions
        Dim ff As Integer
        Dim strTemp As String
        IndexSrc = "<html><head><title>Index</title></head><body><center><img src=""site/header.png""><br>This is not a fancy <i>web server</i>.<br>You can't view many other pages or pictures.<br><br><br><font size=1><i>Powered by an old VB6 server I wrote v" & App.Major & "." & App.Minor & "." & App.Revision & " </font></i></center></body></html>"
        ErrorSrc = "<html><head><title>Error</title></head><body><center><img src=""site/header.png""><br><br><h1>Error. 404 most likely.</h1><br><br>There's nothing here for you, brah.<br />kbai<br><br><br><font size=1><i>Powered by an old VB6 server I wrote " & App.Major & "." & App.Minor & "." & App.Revision & " </font></i></center></html>"

        With frmMain
            .Width = 7400
            .Height = 5450
            .Caption = "an old VB6 server I wrote v" & App.Major & "." & App.Minor & "." & App.Revision
        End With
        
        Call LoadSettings
        
        ' Set the port to listen on
        ProxyListenPort = txtPort.Text
        
        ' Set the port to connect on (port detection not implemented, temp)
        ProxyConnectPort = 80
        
        ' Initiate the server
        InitServer
        On Error GoTo errorinit
        ff = FreeFile
    
        'strVersion = Replace(LoadFile("version.txt"), vbCrLf, "")
        
Exit Sub
        
errorinit:
        MsgBox Err.Number & ": " & Err.Description
    End Sub
        
Private Sub LoadSettings()
    On Error Resume Next
    Dim arrTemp() As String
    Dim arrSplit As String
    Dim i As Integer
    arrTemp = Split(LoadFile("settings.txt"), vbCrLf)
    For i = 0 To UBound(arrTemp)
    
        'MsgBox arrTemp(i)

        Select Case LCase(Split(arrTemp(i), " ")(0))
            
            Case Is = "picture"
            If Not Trim(Replace(arrTemp(i), "picture", "")) = "" Then strPicture = Trim(Replace(arrTemp(i), "picture", ""))
               'Split(arrTemp(i), " ")(1)
               'If strVersion = "" Then strVersion = "2.8.6,2.8.7,2.8.8,2.8.9"
                'AddLog "System", "Settings loaded successfully"
                
            Case Is = "maxusers"

                If IsNumeric(Split(arrTemp(i), " ")(1)) And (Split(arrTemp(i), " ")(1) > 0) Then
                    If Split(arrTemp(i), " ")(1) > 300 Then
                        ProxyClients = 300
                    Else
                        ProxyClients = Split(arrTemp(i), " ")(1)
                    End If
                Else
                    ProxyClients = 15
                End If
                
            Case Is = "port"
                If IsNumeric(Split(arrTemp(i), " ")(1)) And (Split(arrTemp(i), " ")(1) < 65536) And (Split(arrTemp(i), " ")(1) > 0) Then
                    ProxyListenPort = Split(arrTemp(i), " ")(1)
                    txtPort.Text = ProxyListenPort
                Else
                    ProxyListenPort = "80"
                End If
                
            Case Is = "logevents"
                If Split(arrTemp(i), " ")(1) = "1" Or Split(arrTemp(i), " ")(1) = True Then
                chkEvents.Value = vbChecked
                Else
                chkEvents.Value = vbUnchecked
                End If
                
            Case Is = "logurls"
                If Split(arrTemp(i), " ")(1) = "1" Or Split(arrTemp(i), " ")(1) = True Then
                chkLog.Value = vbChecked
                Else
                chkLog.Value = vbUnchecked
                End If
                
            Case Is = "save"
                If Split(arrTemp(i), " ")(1) = "1" Or Split(arrTemp(i), " ")(1) = True Then
                chkSave.Value = vbChecked
                Else
                chkSave.Value = vbUnchecked
                End If
                
            Case Is = "sockcheck"
                If IsNumeric(Split(arrTemp(i), " ")(1)) And (Split(arrTemp(i), " ")(1) < 60001) And (Split(arrTemp(i), " ")(1) > 0) Then
                    tmrSockets.Interval = Split(arrTemp(i), " ")(1)
                Else
                    tmrSockets.Interval = "15000"
                End If
            Case Is = "txtpath"
                If Not Trim(Replace(arrTemp(i), "txtpath", "")) = "" Then txtPath.Text = Trim(Replace(arrTemp(i), "txtpath", ""))
                
        End Select
    
    Next i
    
    If txtPath = "" Then txtPath = App.Path
    
If strPicture = "" Then strPicture = App.Path & "\ChocoboDomDomSoft.png"
If Not Dir(strPicture) <> "" Then
    strPicture = App.Path & "\ChocoboDomDomSoft.png"
    SavePicture imgChocobo, "ChocoboDomDomSoft.png"
    DelChocobo = False
    imgChocobo.Picture = Nothing
    Unload imgChocobo
End If
SetDir strPicture
strPicture = Replace(strPicture, App.Path & "\", "")
If ProxyClients = 0 Then ProxyClients = 15
ReDim Preserve ProxyUsers(ProxyClients) As ProxyType

End Sub
    
Private Sub SetDir(Path As String)
Dim sPath() As String
Dim temp As String
Dim i As Integer

sPath = Split(Path, "\")

    For i = 0 To UBound(sPath) - 1
    temp = temp & sPath(i) & "\"
    Next i
    
    'Path = sPath(UBound(sPath))

Dir1.Path = temp
For i = 0 To File1.ListCount - 1
If File1.List(i) = sPath(UBound(sPath)) Then
    File1.ListIndex = i
    i = File1.ListCount - 1
End If
Next i
End Sub
    
Private Sub SaveSettings()
If chkSave.Value = Checked Then
    Dim strTemp As String
        strTemp = "# This is the port it starts on, 80 is recommended." & vbCrLf
        strTemp = strTemp & "port " & ProxyListenPort & vbCrLf
            strTemp = strTemp & "# Maximum number of concurrent connections, default: 15, max 300." & vbCrLf
        strTemp = strTemp & "MaxUsers " & ProxyClients & vbCrLf
        strTemp = strTemp & "# This changes whether the program saves events to file." & vbCrLf
        strTemp = strTemp & "LogEvents " & chkEvents.Value & vbCrLf
        strTemp = strTemp & "# This changes whether the program logs URLs requestion to file." & vbCrLf
        strTemp = strTemp & "LogURLs " & chkLog.Value & vbCrLf
        strTemp = strTemp & "# This changes whether the program saves settings in this file or not at all." & vbCrLf
        strTemp = strTemp & "save " & chkSave.Value & vbCrLf
        strTemp = strTemp & "# This changes the interval all sockets are checked and unloaded, default: 15000ms" & vbCrLf
        strTemp = strTemp & "SockCheck " & tmrSockets.Interval & vbCrLf
        strTemp = strTemp & "# Path to the placeholder picture when Manga doesn't have an image to display." & vbCrLf
        strTemp = strTemp & "picture " & strPicture & vbCrLf
        'strTemp = strTemp & "# Path to the server directory." & vbCrLf
        'strTemp = strTemp & "txtpath " & txtPath.Text & vbCrLf
        strTemp = strTemp & vbCrLf & vbCrLf & "//* Any text written in this file not by an old VB6 server I wrote will be deleted *//"
        'strTemp = strTemp & vbCrLf & "//* " & GetRandom & " *//"
        
        
        WriteData strTemp, "Settings.txt", False
End If
End Sub
    
    Private Function GetRandom() As String
    Dim i As Integer
    Randomize
    For i = 0 To 24
    'Randomize (Now)
    GetRandom = GetRandom & Asc(Int(Rnd * 127))
    Next i
    End Function
    
    
    Private Sub Form_Unload(Cancel As Integer)
    
        ' Terminate server
        ShutServer
        
        SaveSettings
        On Error Resume Next
        If DelChocobo = True Then Kill "ChocoboDomDomSoft.png"
        
        
        ' Terminate program
        End
    
    End Sub

    Public Sub InitServer()
    On Error GoTo errar
    Dim i As Long
        
    
        ' Load winsock objects
        For i = 1 To ProxyClients
            Load srvIn(i)
        Next i
        '
        srvIn(0).Close
        ' Set the port to listen on
        srvIn(0).LocalPort = ProxyListenPort
        
        ' Set the port to connect on
        
        ' Start listening
        srvIn(0).Listen
        
        ' Add to log
        AddLog "System", "Server started, listening on port " & ProxyListenPort
        AddLog "System", "Max connections set to: " & ProxyClients & " (Change in settings.ini)"
        AddLog "System", "Running check on sockets every " & tmrSockets.Interval & "ms."
        
        
        
    Exit Sub
errar:
    Select Case Err.Number
        Case Is = "10048"
            MsgBox "Port " & txtPort.Text & " is in use! This port must be free. Please close the program using this port and then click Restart in the bottom corner to try again!"
        Case Else
        MsgBox Err.Number & ": " & Err.Description
    End Select
    
    End Sub

    Public Sub ShutServer()
    On Error Resume Next
    Dim i As Long
    
        ' Unload winsock objects
        For i = 1 To ProxyClients
            Unload srvIn(i)
        Next i
        '
        If Dir(App.Path & "\Logfile.txt") <> "" Then AddLog "System", "Server shutting down.."
    End Sub


    Private Sub srvIn_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i As Long
            
            If lstProxyLog.ListItems.Count > 350 Then lstProxyLog.ListItems.Clear
            
        ' Handle the ConnectionRequest
        For i = 1 To ProxyClients
            If (srvIn(i).State = sckClosing) Then srvIn(i).Close
            If (srvIn(i).State = sckClosed) And (ProxyUsers(i).ProxyUserIp = Empty) Then
                Exit For
            End If
        Next i
        
        If i > ProxyClients Then
        Exit Sub
        End If
        
        AddLog "System", "Connection on socket " & i
        
        ' Just in case all sockets are full
        If Not (srvIn(i).State = sckClosed) Then Exit Sub
        
        ' Only accept local connections
        ' If Not (srvIn(i).RemoteHostIP = srvIn(0).LocalIP) Then Exit Sub
        
        ProxyUsers(i).ProxyUserIp = srvIn(i).RemoteHost
        If ProxyUsers(i).ProxyUserIp = "" Then ProxyUsers(i).ProxyUserIp = "127.0.0.1"
        ProxyUsers(i).SendCompleted = True
        ' Set the ip of the proxy user
        srvIn(i).Accept requestID
        
    UpdateStatus

End Sub

Private Sub srvIn_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo Error
    Dim SckData As String   ' Socket data
    Dim remHost As String   ' What is the remote host ?
    Dim tarSite As String   ' What site is the proxy user requesting ?
    Dim unfSite As String   ' Same as above but unformated!
    Dim strHeader As String
    Dim strUpdater As String

        srvIn(Index).GetData SckData, vbString, bytesTotal
        
        unfSite = URLDecode(FetchRequestedPage(SckData))
        WriteUrlLog (unfSite)
        'writedata "<" & SckData, "debug", True
        ProxyUsers(Index).url = unfSite
        'Debug.Print SckData
        'Sleep 250

    remHost = GetHttpVars(SckData, "Host")
        
    'If ProxyUsers(Index).url = "/" Then
       ' AddLog "WebSrv", srvIn(Index).RemoteHostIP & " requested root directory."
       ' strHeader = "HTTP/1.1 200 OK" & vbCrLf & _
       ' "Content-Length: " & Len(IndexSrc) & vbCrLf & _
       ' "Content-Type: text/html" & vbCrLf & _
       ' "Server: an old VB6 server I wrote" & vbCrLf & _
       ' "Connection: Close" & vbCrLf & _
       ' "Accept-Ranges: none" & vbCrLf & vbCrLf
       '
        '''srvIn(Index).SendData strHeader
        ''Call SendData(strHeader, Index)
        ''Call SendData(IndexSrc, Index)
        '''srvIn(Index).SendData IndexSrc
        'SckData = ""
        'ProxyUsers(Index).ProxyUserIp = srvIn(Index).RemoteHostIP
        
        If InStr(1, LCase(ProxyUsers(Index).url), "/site/header.png") >= 1 Then
        
    If FileExists(strPicture) Then
        strHeader = "HTTP/1.1 200 OK" & vbCrLf & _
        "Content-Length: " & FileLen(strPicture) & vbCrLf & _
        "Content-Type: image/png" & vbCrLf & _
        "Server: vbHTTPd/1.0" & vbCrLf & _
        "Connection: Close" & vbCrLf & _
        "Accept-Ranges: none" & vbCrLf & vbCrLf

        'srvIn(Index).SendData strHeader
         ProxyUsers(Index).SendCompleted = True
        Call SendData(strHeader, Index)
        'ProxyUsers(Index).SendCompleted = False
        ''doevents
        SendFile srvIn(Index), strPicture, "4096", Index
        SckData = ""
        ProxyUsers(Index).ProxyUserIp = srvIn(Index).RemoteHostIP
        UnloadProxyUser (Index): Exit Sub
    Else
        Send404 (Index)
        'doevents
        AddLog "Error", "Manga placeholder image not found: " & strPicture
        ProxyUsers(Index).ProxyUserIp = srvIn(Index).RemoteHostIP
        UnloadProxyUser (Index): Exit Sub
    End If
End If

    If Left(LCase(ProxyUsers(Index).url), 1) = "/" Then
        Call ServeFiles(Index)
   
   ' ElseIf LCase(ProxyUsers(Index).url) = "/favicon.ico" Then
        'AddLog "WebSvr", "Client requested favicon.ico"
        'If FileExists(App.Path & "/web/favicon.ico") Then
        '    strHeader = "HTTP/1.1 200 OK" & vbCrLf & _
        '        "Content-Length: " & FileLen(App.Path & "/web/favicon.ico") & vbCrLf & _
        '        "Content-Type: " & GetMime("favicon.ico") & vbCrLf & _
        '        "Server: an old VB6 server I wrote" & vbCrLf & _
        '        "Accept-Ranges: none" & vbCrLf & vbCrLf
        '    Call SendData(strHeader, Index)
        '    DoEvents
        '    Call SendFile(srvIn(Index), App.Path & "/web/favicon.ico", 2048, Index)
        'Else
        '    Send404 (Index)
        'End If
'GoTo debugg
    Else
        If Not ProxyUsers(Index).url = "" Then Call Send404(Index)
    End If

If ProxyUsers(Index).Wait = False Then
    If srvIn(0).LocalPort = "80" Then
        remHost = ""
        'If Not SckData = "" Then SckData = ""
        ProxyUsers(Index).ProxyUserIp = srvIn(Index).RemoteHostIP
        UnloadProxyUser (Index): Exit Sub
        'End If
    End If
End If

debugg:
        If IsEmpty(unfSite) Or IsEmpty(remHost) Then UnloadProxyUser (Index): Exit Sub
    Exit Sub
    
Error:
    ProxyUsers(Index).ProxyUserIp = "127.0.0.1"
    'UnloadProxyUser (Index)
    Select Case Err.Number
        Case Is = 5
            AddLog "Error", "Invalid procedure call or argument!"
        Case Is = 40006
            AddLog "Error", "Client disconnected before we could send data!"
        Case Else
        'MsgBox Err.Number & Err.Description
        AddLog "Error", Err.Number & " " & Err.Description
    End Select
    
'doevents
Exit Sub

    End Sub
    
Public Function ServeFiles(Index As Integer)
    Dim strFile As String, strHeader As String
    DoEvents
strFile = ProxyUsers(Index).url  'Replace(LCase(ProxyUsers(Index).URL), "/web/", "")
'Debug.Print strFile

AddLog "WebSrv", srvIn(Index).RemoteHostIP & " requested " & URLDecode(strFile)

If Not FileExists(txtPath.Text & strFile) Then
        If DirExists(txtPath.Text & strFile) Then
        If Not Right(strFile, 1) = "/" Then strFile = strFile & "/"
        SendDir (Index), strFile
    Else
        Call Send404(Index)
    End If
Else
    strHeader = "HTTP/1.1 200 OK" & vbCrLf & _
        "Content-Length: " & FileLen(txtPath.Text & strFile) & vbCrLf & _
        "Content-Type: " & GetMime(strFile) & vbCrLf & _
        "Server: an old VB6 server I wrote" & vbCrLf & _
        "Accept-Ranges: none" & vbCrLf & vbCrLf

        'srvIn(Index).SendData strHeader
        Call SendData(strHeader, Index)
        DoEvents
        Call SendFile(srvIn(Index), txtPath.Text & strFile, 2048, Index)
        'Call SendData(LoadFile(txtPath.text & Replace(strFile, "/", "\")), Index)
    'End If
End If
End Function
    
Public Function GetMime(File As String) As String
 Dim strMime As String
 Dim fH As Integer
 Dim MimeTypes As String
 Dim SortMimes As Variant
 Dim strTemp As Variant
    strMime = Mid(File, InStrRev(File, ".") + 1&)
    If InStr(1, strMime, "/") Then GetMime = "text/html": Exit Function
    
    If Not FileExists(App.Path & "\mimes.txt") Then
        AddLog "Error", "mimes.txt not found, creating default mimes.txt"
        Call WriteMimes
    End If
    
    SortMimes = Split(LoadFile(App.Path & "\mimes.txt"), vbCrLf)
    
    Dim i As Integer
    
    For i = 0 To UBound(SortMimes)
        
    SortMimes(i) = Replace(SortMimes(i), vbTab, " ")
            
    strTemp = Split(SortMimes(i), " ")
            
        If UBound(strTemp) >= 1 Then
                
            If strTemp(0) = strMime Then GetMime = strTemp(1): Exit For
                
        End If
        
    Next i
    
End Function
    
Public Function WriteMimes()
    WriteData "#MIME TYPES" & vbCrLf & "#FORMAT: MIME [space/tab] MIMETYPE" & vbCrLf & "txt text/plain" & vbCrLf & "log text/plain" & vbCrLf & "ini text/plain" & vbCrLf & "inf text/plain" & vbCrLf & "htm text/html" & vbCrLf & "html text/html" & vbCrLf & "shtm text/html" & vbCrLf & "as text/html" & vbCrLf & "php text/html" & vbCrLf & "zip application/x-zip-compressed" & vbCrLf & "rar application/x-rar-compressed" & vbCrLf & "7z application/x-7z-compressed" & vbCrLf & "jpg image/jpeg" & vbCrLf & "jpeg image/jpeg" & vbCrLf & "png image/png" & vbCrLf & "bmp image/bmp" & vbCrLf & "gif image/gif" & vbCrLf & "avi video/avi" & vbCrLf & "mp4 video/mp4" & vbCrLf & "flv video/flv" & vbCrLf & "wmv video/wmv" & vbCrLf & "mp3 audio/mp3" & vbCrLf & "mp2 audio/mp2" & vbCrLf & "wav audio/wav" & vbCrLf & "wma audio/wma" & vbCrLf & "iso application/x-iso-disc-image" & vbCrLf & "dmg application/x-dmg-disc-image" & vbCrLf & "daa application/x-daa-disc-image" & vbCrLf & "exe application/windows-executeable", "mimes.txt", False
End Function
    
Public Function FileExists(sFullPath As String) As Boolean
    Dim oFile As New Scripting.FileSystemObject
    FileExists = oFile.FileExists(sFullPath)
End Function

Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
End Function
    
Public Function SendDir(Index As Integer, Path As String)
    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim File As File
    Dim Directory As Variant
    Dim strDirectory As String
    Dim i As Integer
    Dim strDir As String
    Dim strHeader As String
    strDir = "<html><head><title>Dir list for " & Path & "</title></head><body><center><img src=""site/header.png""><br><br>"
    

    
strDir = strDir & "<table width=""60%"" cellspacing=""1"" cellpadding=""1"" border=""1"">"
                strDir = strDir & "<tr>"
                strDir = strDir & "<th>File Name</th>"
                strDir = strDir & "<th width=""15%"">File Type</th>"
                'strDir = strDir & "<th width=>Refresh</th>"
                '    <th width="10%">Delete</th>
                '    <th width="15%">Share It</th>
                strDir = strDir & "</tr>"
                
    strDir = strDir & "<tr><td align=""left""> <a href=""" & "." & """>.</a></td><td align=""left"">Change Dir</td></tr>"
    strDir = strDir & "<tr><td align=""left""> <a href=""" & ".." & """>..</a></td><td align=""left"">Change Dir</td></tr>"
    
Set FSfolder = FS.GetFolder(txtPath.Text & Path)
For Each Directory In FSfolder.SubFolders
DoEvents

    strDir = strDir & "<tr>"
    strDir = strDir & "<td align=""left""> <a href=""" & URLEncode(Path & Mid(Directory, InStrRev(Directory, "\") + 1&)) & """>" & Mid(Directory, InStrRev(Directory, "\") + 1&) & "</a></td>"
    strDir = strDir & "<td align=""left""> "
    If InStr(1, Directory, ".") > 0 Then
        strDir = strDir & Mid(Directory, InStrRev(Directory, ".") + 1&)
    Else
        strDir = strDir & "Directory"
    End If
    strDir = strDir & "</td>"
    strDir = strDir & "</tr>"
        
Next Directory

For Each File In FSfolder.Files
DoEvents
        
    strDir = strDir & "<tr>"
    strDir = strDir & "<td align=""left""> <a href=""" & URLEncode(Path & Mid(File, InStrRev(File, "\") + 1&)) & """>" & Mid(File, InStrRev(File, "\") + 1&) & "</a></td>"
    strDir = strDir & "<td align=""left""> "
    If InStr(1, File, ".") > 0 Then
        strDir = strDir & Mid(File, InStrRev(File, ".") + 1&)
    Else
        strDir = strDir & "Guess: txt"
    End If
    strDir = strDir & "</td>"
    strDir = strDir & "</tr>"
        
Next File
    
    strDir = strDir & "</table>" & "<br><font size=1><i>Powered by an old VB6 server I wrote " & App.Major & "." & App.Minor & "." & App.Revision & " </font></i></center><br /></body></html>"
    ' strDir
    Set FSfolder = Nothing
    

    strHeader = "HTTP/1.1 200 OK" & vbCrLf & _
        "Content-Length: " & Len(strDir) & vbCrLf & _
        "Content-Type: text/html" & vbCrLf & _
        "Server: an old VB6 server I wrote" & vbCrLf & _
        "Connection: Close" & vbCrLf & _
        "Accept-Ranges: none" & vbCrLf & vbCrLf & strDir
        
        SendData strHeader, Index
        DoEvents
        SendData strDir, Index
        'Call SendData(strHeader, Index)
    
End Function
    
Public Sub Send404(Index As Integer)
Dim strHeader As String
AddLog "WebSrv", srvIn(Index).RemoteHostIP & " got a 404 from " & ProxyUsers(Index).url
    strHeader = "HTTP/1.1 404 Not Found" & vbCrLf & _
        "Content-Length: " & Len(ErrorSrc) & vbCrLf & _
        "Content-Type: text/html" & vbCrLf & _
        "Server: an old VB6 server I wrote" & vbCrLf & _
        "Connection: Close" & vbCrLf & _
        "Accept-Ranges: none" & vbCrLf & vbCrLf
        
        '"ETag: ""20101031212956;""" & vbCrLf & _

        'srvIn(Index).SendData strHeader
        Call SendData(strHeader, Index)
        'ProxyUsers(Index).SendCompleted = False
        Call SendData(ErrorSrc, Index)
        'srvIn(Index).SendData ErrorSrc
        'writedata ">" & strHeader, "debug"
        ProxyUsers(Index).ProxyUserIp = srvIn(Index).RemoteHostIP
End Sub
    
Public Function SendData(Data As String, Index As Integer) As Boolean
    Dim Continue As Boolean
    Continue = True
    Dim Timeout As Integer
'writedata ">" & Data, "debug"
Do While Continue = True
Timeout = Timeout + 1
    If ProxyUsers(Index).SendCompleted = True Then
        If SendData = True Then Continue = False: Exit Do
            srvIn(Index).SendData Data
            ProxyUsers(Index).SendCompleted = False
            SendData = True
    End If
    If Timeout = 100 Then
    ProxyUsers(Index).SendCompleted = True
    End If
DoEvents
Loop

    End Function
    
    Private Sub srvIn_Close(Index As Integer)
    
        ' Add to log
        AddLog "System", "Socket closed (" & Index & ")"
    
        If (Index = 0) Then
            srvIn(0).Close
            srvIn(0).Listen
        End If
    
        If ProxyUsers(Index).Wait = True Then ProxyUsers(Index).Wait = False
        ProxyUsers(Index).SendCompleted = True
        ' Unload the user
        UnloadProxyUser (Index)
    

    
    End Sub
    
Private Sub srvIn_SendComplete(Index As Integer)
ProxyUsers(Index).SendCompleted = True
End Sub

Private Sub srvIn_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
        ' Add to log
        'AddLog "SrvIn", "Socket error! (" & Index & ")"
        AddLog ">>>", Description & " - i:" & Index

        ' Error occured, unload the user
            If (srvIn(Index).State = sckClosing) Then ProxyUsers(Index).ProxyUserIp = "127.0.0.1": srvIn(Index).Close
            If (srvIn(Index).State = sckError) Then ProxyUsers(Index).ProxyUserIp = "127.0.0.1": srvIn(Index).Close
        UnloadProxyUser (Index)
        
        ' Close the socket
        If (Index = 0) Then srvIn(0).Close
        
End Sub

Private Sub Timer1_Timer()
DoEvents
End Sub

Private Sub tmrSockets_Timer()
On Error Resume Next
Dim i As Integer
'AddLog "System", "Ports check initiating.."
        For i = 1 To ProxyClients
            If (srvIn(i).State = sckClosing) Then ProxyUsers(i).ProxyUserIp = "127.0.0.1": srvIn(i).Close
            If (srvIn(i).State = sckError) Then ProxyUsers(i).ProxyUserIp = "127.0.0.1": srvIn(i).Close
            If (srvIn(i).State = sckClosed) Then
                UnloadProxyUser (i)
            End If
        Next i
    UpdateStatus

End Sub

Private Sub txtPort_Change()
On Error Resume Next
If IsNumeric(txtPort.Text) Then ProxyListenPort = txtPort.Text Else MsgBox "Must be numeric, don't go over 65535."
SaveSettings
End Sub
