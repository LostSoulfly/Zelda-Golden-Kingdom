Attribute VB_Name = "modTCP"

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long
AddLog "Received new connection.."
    If (Index = 0) Then
        i = FindOpenServerSlot

        If i <> 0 Then
            ' we can connect them
            ClearServer i
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
            
        Else
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            DoEvents
            frmServer.Socket(0).Close
            frmServer.Socket(0).Listen
        End If
    End If

End Sub

Function FindOpenServerSlot() As Long
    Dim i As Long
    FindOpenServerSlot = 0

    For i = 1 To MAX_SERVERS

        If frmServer.Socket(i).State <> 7 Then
            FindOpenServerSlot = i
            Exit Function
        End If

    Next

End Function

Sub SocketConnected(ByVal Index As Long)
'this runs when a connection has been accepted.
'Dim Buffer As New clsBuffer
'Set Buffer = New clsBuffer

'Buffer.WriteLong HHello

SendDataHub Index, BuildGeneric(HHello, "")

AddLog "Requesting server info.."

UpdateCaption

'Set BuildGeneric = Nothing

End Sub

Public Function BuildGeneric(PacketNum As Long, Data As String) As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong PacketNum
Buffer.WriteString Data

BuildGeneric = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function SendCommand(CommandNum As Long, Data As String) As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong HCommand
Buffer.WriteLong CommandNum
Buffer.WriteString Data

SendCommand = Buffer.ToArray

Set Buffer = Nothing

End Function

Sub CloseSocket(ByVal Index As Long)
Dim i As Integer

    Call AddLog("Connection from " & frmServer.Socket(Index).RemoteHostIP & " has been terminated.")
    frmServer.Socket(Index).Close
    ClearServer Index
    UpdateCaption
    WriteServerFile
    
End Sub

Sub ClearServer(Index As Long)
    Call ZeroMemory(ByVal VarPtr(Server(Index)), LenB(Server(Index)))
    Set Server(Index).Buffer = New clsBuffer
    
    With Server(Index)
        .CurrentPlayers = 0
        .DataBytes = 0
        .DataPackets = 0
        .DataTimer = 0
        .InactiveTime = 0
        .MaxPlayers = 0
        .Name = vbNullString
        .Online = False
        .Uptime = 0
    End With
    
End Sub

Sub SendDataToAllHub(ByRef Data() As Byte, Optional Index As Long)
Dim i As Long

For i = 1 To MAX_SERVERS
    If frmServer.Socket(i).State = sckConnected And i <> Index Then
        SendDataHub i, Data()
    End If
Next

End Sub

Sub SendDataHub(Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte

    If frmServer.Socket(Index).State = sckConnected Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        frmServer.Socket(Index).SendData Buffer.ToArray()
        
        PacketsSent = PacketsSent + 1
        BytesSent = BytesSent + (UBound(TempData) - LBound(TempData)) + 1
        
    End If
End Sub
