Attribute VB_Name = "modHubTCP"
Public isHubConnected As Boolean
Public hubLastHeard As Long
Public useHubServer As Boolean

Private hubBuffer As clsBuffer

Public Enum HubPackets
    HHello = 1
    HServerInfo
    HShutdown
    HLog
    HGlobalMsg
    HMSG_COUNT
End Enum

Public Const MAX_SERVERS As Long = 10
Public HandleDataHub(HMSG_COUNT) As Long

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Sub InitHubMessages()

HandleDataHub(HHello) = GetAddress(AddressOf Hello)
HandleDataHub(HShutdown) = GetAddress(AddressOf HandleShutdownFromHub)
HandleDataHub(HGlobalMsg) = GetAddress(AddressOf HandleForwardGlobalMsg)

End Sub

Sub HandleDataHubSub(ByRef Data() As Byte)
'Dim buffer As clsBuffer
Dim MsgType As Long

    'Set buffer = New clsBuffer
    'buffer.WriteBytes Data()

    'MsgType = buffer.ReadLong
    MsgType = ReadHandleDataType(Data)

    If MsgType < 0 Then
        Exit Sub
    End If

    If MsgType >= HMSG_COUNT Then
        Exit Sub
    End If

    'CallWindowProc HandleDataSub(MsgType), index, buffer.ReadBytes(buffer.length), 0, 0
    CallWindowProc HandleDataHub(MsgType), index, Data, 0, 0
'Set buffer = Nothing
End Sub

Sub SendDataHub(ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim TempData() As Byte
If Not useHubServer Then Exit Sub
    If frmServer.hubSocket.state = sckConnected Then
        Set buffer = New clsBuffer
        TempData = Data
        
        buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteBytes TempData()
              
        frmServer.hubSocket.SendData buffer.ToArray()
        
        PacketsSent = PacketsSent + 1
        BytesSent = BytesSent + (UBound(TempData) - LBound(TempData)) + 1
        
    End If
End Sub

Public Sub CheckHubConnection()

If hubLastHeard = 0 Then hubLastHeard = GetRealTickCount

    With frmServer.hubSocket
        If GetRealTickCount > hubLastHeard + 5000 Then
            If .state <> sckConnected Then
                If isHubConnected = True Then
                    TextAdd "Hub server disconnected."
                    GlobalMsg "Hub Server disconnected.", Green, False
                End If
                
                isHubConnected = False
                'so let's try to connect again..
                .Close
                .Connect
                hubLastHeard = GetRealTickCount
                
                Do While GetRealTickCount < (hubLastHeard + 500)
                    
                    DoEvents
                    Sleep 10
                    
                    If .state = sckConnected Then
                        'we've connected!
                        isHubConnected = True
                        GlobalMsg "Hub server connected.", Green, False
                        Exit Do
                    End If
                    
                Loop
            Else
                isHubConnected = True
            End If
        End If
    End With
    

End Sub

Public Sub HubIncomingData(ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

If hubBuffer Is Nothing Then Set hubBuffer = New clsBuffer

    frmServer.hubSocket.GetData buffer, vbUnicode, DataLength
    
    hubBuffer.WriteBytes buffer()
    
    If hubBuffer.length >= 4 Then pLength = hubBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= hubBuffer.length - 4
        If pLength <= hubBuffer.length - 4 Then
            hubBuffer.ReadLong
            HandleDataHubSub hubBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If hubBuffer.length >= 4 Then pLength = hubBuffer.ReadLong(False)
    Loop
    hubBuffer.Trim

End Sub

Public Sub SendServerInfo()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong HServerInfo
    buffer.WriteLong TotalPlayersOnline
    buffer.WriteLong MAX_PLAYERS
    buffer.WriteString frmServer.Socket(0).LocalPort
    buffer.WriteLong StartTick
    
    SendDataHub buffer.ToArray

    Set buffer = Nothing
End Sub

Public Sub SendHubLog(Text As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong HLog
    buffer.WriteString Text
    
    SendDataHub buffer.ToArray

    Set buffer = Nothing
End Sub

Public Sub ForwardGlobalMsg(ByVal msg As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong HGlobalMsg
    buffer.WriteString msg
    
    SendDataHub buffer.ToArray

    Set buffer = Nothing

End Sub

Private Sub Hello(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next

    TextAdd "Received HELLO from Hub."
    SendServerInfo
    
End Sub

Private Sub HandleShutdownFromHub(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next

    TextAdd "ShutdownFromHub received."
        
    isShuttingDown = True
    frmServer.cmdShutDown.Caption = "Cancel"

End Sub

Private Sub HandleForwardGlobalMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As New clsBuffer
    Set buffer = New clsBuffer
    Dim msg As String, color As Long
    buffer.WriteBytes Data()
    msg = buffer.ReadString
    'color = buffer.ReadLong
    
    Call GlobalMsg(msg, White, False, False)
    
    Set buffer = Nothing
End Sub
