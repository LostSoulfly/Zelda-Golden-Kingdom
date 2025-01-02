Attribute VB_Name = "mProxy"
Option Explicit
   
    ' Set max clients
    Public ProxyClients As Integer

    Public Type ProxyType
    
        ProxyUserData As String         ' The data recieved from the proxy user        ->
        ProxyUserIp As String           ' The ip of the proxy user (not used yet!)
        
        ProxyRequData As String         ' The data received after the request          <-
        ProxyRequHost As String         ' The host/domain to connect on
        Wait As Boolean
        url As String
        SendCompleted As Boolean
    End Type
    
    ' Declare the ProxyUser array
    'Public ProxyUsers(10) As ProxyType
    
   Global ProxyUsers() As ProxyType

    ' Declare som public vars
    Public ProxyListenPort As Integer
    Public ProxyConnectPort As Integer

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_WRITE = &H2
Public Const FILE_SHARE_READ = &H1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3

Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub SendFile(SocketObject As Variant, ByVal FilePath As String, ByVal PacketSize As Long, Optional Index As Integer)
Dim lonFF As Long, bytData() As Byte
Dim lonCurByte As Long, lonSize As Long
Dim lonPrevSize As Long
Dim strHeader As String
Dim strTotal As String
Dim hFile As Long
Dim Continue As Boolean

Continue = True
On Error GoTo ErrorHandler
'writedata ">" & FilePath, "debug"
hFile = CreateFile(FilePath, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

lonSize = FileLen(FilePath)

lonFF = FreeFile
        
'strHeader = "HTTP/1.1 200 OK" & vbCrLf & _
"Content-Type: image/png" & vbCrLf & _
"Last-Modified: " & Now & vbCrLf & _
"Accept -Ranges: none" & vbCrLf & _
"ETag: ""f0ac5f81c4e6ca1:0""" & vbCrLf & _
"Server: Microsoft -IIS / 7#" & vbCrLf & _
"X -Powered - By: ASP.NET" & vbCrLf & _
"Date: " & Now & vbCrLf & _
"Content-Length: " & lonSize & vbCrLf & vbCrLf & vbCrLf
ProxyUsers(Index).SendCompleted = True
Dim tBuff() As Byte, Readed As Long, tS As String
Do While Continue = True
If ProxyUsers(Index).SendCompleted = True Then
        ReDim tBuff(PacketSize - 1)
        ReadFile hFile, tBuff(0), UBound(tBuff) + 1, Readed, ByVal 0
        If Readed = 0 Then
        Continue = False
        Else
            If Readed <> UBound(tBuff) + 1 Then ReDim Preserve tBuff(Readed - 1)
            SocketObject.SendData tBuff
            ProxyUsers(Index).SendCompleted = False
            
       End If
End If
DoEvents
Loop

CloseHandle hFile
'Close #lonFF
Exit Sub

ErrorHandler:
    Debug.Print "Send File Error"
    Debug.Print "---------------"
    Debug.Print "Number:      " & Err.Number
    Debug.Print "Description: " & Err.Description
    Debug.Print "File:        " & FilePath
    ProxyUsers(Index).SendCompleted = True
End Sub

    Public Function LoadFile(Path As String) As String
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile

    Open (Path) For Input As #ff
        LoadFile = Input(LOF(ff), #ff)
    Close #ff
    End Function

    Public Function UnloadProxyUser(Index As Integer)
    Dim i As Integer
        ' Check so its not trying to unload the main socket
        If (Index = 0) Then Exit Function
        
        ' Check so the user hasent already been unloaded
        If IsEmpty(ProxyUsers(Index).ProxyUserIp) Then Exit Function
        
        ' Add to log
        'AddLog "System", "Unloading client " & Index & " (" & ProxyUsers(Index).ProxyUserIp & ")"
        
        With frmMain
        
            ' Close the sockets
            .srvIn(Index).Close
            
        End With
    
        ' Empty the proxy user array
        With ProxyUsers(Index)
        
            .ProxyRequData = Empty
            .ProxyRequHost = Empty
            .ProxyUserData = Empty
            .ProxyUserIp = Empty
            
        End With
        
    UpdateStatus

    End Function
    
Public Function UpdateStatus()
    Dim i As Integer, intNum As Integer
        For intNum = 1 To ProxyClients
        If Not frmMain.srvIn(intNum).State = 0 Then Debug.Print "Socket " & intNum & " Status: " & frmMain.srvIn(intNum).State
            If frmMain.srvIn(intNum).State = sckClosed Or frmMain.srvIn(intNum).State = sckClosing Or frmMain.srvIn(intNum).State = sckError Then
                Else
                i = i + 1
            End If
        Next
    frmMain.Caption = "an old VB6 server I wrote " & App.Major & "." & App.Minor & "." & App.Revision & " - Open Sockets: " & i & "/" & ProxyClients
End Function
    
    Public Function IsEmpty(ByRef CheckVal) As Boolean
    
        ' Check if the value is empty, if so then return true
        If (Len(CheckVal) = 0) Then IsEmpty = True
        
    End Function

Public Function AddLog(tSocket As String, nEvent As String)
On Error GoTo clearit
    'DoEvents
        With frmMain.lstProxyLog
        
            ' List the time
            .ListItems.Add , , Time
            
            ' What socket triggered the event ?
            .ListItems(.ListItems.Count).SubItems(1) = tSocket
            
            ' List the event
            .ListItems(.ListItems.Count).SubItems(2) = nEvent
            
        End With
'DoEvents
    frmMain.lstProxyLog.ListItems(frmMain.lstProxyLog.ListItems.Count).EnsureVisible

If frmMain.chkEvents.Value = Unchecked Then Exit Function
    WriteData "[" & Time & "] " & nEvent, "Logfile.txt"
    
    Exit Function
clearit:
frmMain.lstProxyLog.ListItems.Clear
Resume Next
End Function
    
    Public Function WriteUrlLog(url As String)
    Dim ff As Integer
    Dim eDate As String
    If frmMain.chkLog.Value = vbUnchecked Then Exit Function
        ' Get a free file number
        
        If IsEmpty(url) Then Exit Function
        
        ff = FreeFile
        

        
        ' Set the date of the event
        eDate = ("[" & Replace(Date, "-", "/") & Chr(32) & Time & "]")
        
        ' Update database
        Open (App.Path & "/ProxyLog.txt") For Append As #ff
            Print #ff, (eDate & Chr(32) & url)
        Close #ff
    
    End Function

    Public Function WriteData(Data As String, strFile As String, Optional Append As Boolean = True)
    Dim ff As Integer
    
        ff = FreeFile
        If Append = True Then Open (App.Path & "\" & strFile) For Append As #ff Else: Open (App.Path & "\" & strFile) For Output As #ff
        
            Print #ff, (Data)
        Close #ff
    
    End Function

    ' Remove the Get Parameters that is sometimes included in a request.
    Public Function RemoveGet(Data As String) As String
    Dim i As Integer
        
        ' Just in case
        If (Len(Data) = 1) Then RemoveGet = Data: Exit Function
        
        ' Parse info
        For i = 1 To Len(Data)
            
            ' Find the char pos
            If (Mid(Data, i, 1) = "?") Then
                
                ' Match found
                If Not (i = (Len(Data) - 1)) Then
                    
                    ' Return result
                    RemoveGet = Mid(Data, 1, i - 1)
                    Exit For
                    
                End If
                
            ElseIf (i = (Len(Data) - 1)) Then
                
                ' Return result
                RemoveGet = Data
                Exit For
                
            End If
            
        Next i
        
    End Function
    
    Public Function StripHost(url As String) As String
    Dim i As Long
        
        ' Just in case
        If (InStr(url, "\")) Then url = Replace(url, "\", "/")
        
        ' Convert string to lowercase
        url = (LCase(url))
        
        ' Check if its a valid url (not ssl)
        If Not (Left(url, 7) = "http://") Then Exit Function
        
        For i = 8 To Len(url)
            If (Mid(url, i, 1) = "/") Then Exit For
        Next i
        
        ' Return result
        StripHost = (Right(url, (Len(url) - i) + 1))
    
    End Function

    Public Function GetDomain(url As String) As String
    Dim i As Long
    
        ' Just in case
        If (InStr(url, "\")) Then url = Replace(url, "\", "/")
        
        ' Convert string to lowercase
        url = (LCase(url))
        
        ' Check if its a valid url (not ssl)
        If Not (Left(url, 7) = "http://") Then Exit Function
        
        For i = 8 To Len(url)
            If (Mid(url, i, 1) = "/") Then Exit For
        Next i
        
        ' Return result
        GetDomain = Left(url, i)
        
    End Function
    
    ' Fetch the requested page (param included)
    Public Function FetchRequestedPage(Data As String) As String
    Dim sData() As String
    Dim i As Integer
    
        ' Check so theres no empty string
        If Not (Data = "") Then
        
            ' Split data
            sData = Split(Data, Chr(13) & Chr(10))
            
            ' Check http method
            If Mid(sData(0), 1, 3) = "GET" Then
            
                ' Parse string
                For i = 5 To Len(sData(0))
                    If Mid(sData(0), i, 1) = Chr(32) Then Exit For
                Next i
                
                ' Return result
                FetchRequestedPage = Mid(sData(0), 5, i - 5)
                
            Else
                
                ' Parse string
                For i = 6 To Len(sData(0))
                    If Mid(sData(0), i, 1) = Chr(32) Then Exit For
                Next i
                
                ' Return result
                FetchRequestedPage = Mid(sData(0), 6, i - 6)
                
            End If
        Else
            FetchRequestedPage = "/"
        End If
        
    End Function

    ' Feth the http vars that are included in a request.
    Public Function GetHttpVars(Data As String, tVar As String) As String
    Dim i As Integer, j As Integer
    Dim sData() As String
    
        ' Split the http request
        sData = Split(Data, vbCrLf)
        
        ' Parse the info
        For i = LBound(sData) To UBound(sData)
            For j = 1 To Len(sData(i))
            
                If Mid(sData(i), j, 2) = (Chr(58) & Chr(32)) Then
                    
                    ' If a match then return value
                    If Mid(sData(i), 1, j - 1) = tVar Then
                        GetHttpVars = Mid(sData(i), j + 2, Len(sData(i)) - (j + 1))
                        Exit For
                    End If
                    
                End If
                
            Next j
        Next i
        
    End Function
Public Function URLDecode(sEncodedURL As String) As String
 On Error GoTo Catch
 
 Dim iLoop As Integer
 Dim sRtn As String
 Dim sTmp As String
 
 If Len(sEncodedURL) > 0 Then
    For iLoop = 1 To Len(sEncodedURL)
        sTmp = Mid(sEncodedURL, iLoop, 1)
        sTmp = Replace(sTmp, "+", " ")
    If sTmp = "%" And Len(sEncodedURL) + 1 > iLoop + 2 Then
        sTmp = Mid(sEncodedURL, iLoop + 1, 2)
        sTmp = Chr(CDec("&H" & sTmp))
        iLoop = iLoop + 2
    End If
        sRtn = sRtn & sTmp
    Next iLoop
 URLDecode = sRtn
 End If
Finally:
 Exit Function
Catch:
 URLDecode = ""
 Resume Finally
End Function

Function URLEncode(strBefore As String) As String
 Dim strAfter As String
 Dim intLoop As Integer
 
 If Len(strBefore) > 0 Then
For intLoop = 1 To Len(strBefore)
 Select Case Asc(Mid(strBefore, intLoop, 1))
    Case 48 To 57, 65 To 90, 97 To 122, 46, 45, 95, 42 '0-9, A-Z, a-z . - _ *
strAfter = strAfter & Mid(strBefore, intLoop, 1)
    Case 32
strAfter = strAfter & "%20"
    Case 47 '/
    strAfter = strAfter & "/"
    Case Else
strAfter = strAfter & "%" & Right("0" & Hex(Asc(Mid(strBefore, intLoop, 1))), 2)
 End Select
Next
 End If
 
URLEncode = strAfter
End Function
