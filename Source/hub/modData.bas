Attribute VB_Name = "modData"
Option Explicit

Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Public Enum HubPackets
    HHello = 1
    HServerInfo
    HShutdown
    HLog
    HGlobalMsg
    HCommand
    HMSG_COUNT
End Enum

Public Type ServerRec
    Buffer As clsBuffer
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    InactiveTime As Long
    Name As String
    CurrentPlayers As String
    MaxPlayers As String
    Uptime As Long
    Online As Boolean
    Port As Long
End Type

Public Const MAX_SERVERS As Long = 50
Public HandleDataSub(HMSG_COUNT) As Long
Public Server(1 To MAX_SERVERS) As ServerRec

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()

    HandleDataSub(HServerInfo) = GetAddress(AddressOf HandleServerInfo)
    HandleDataSub(HLog) = GetAddress(AddressOf HandleLog)
    HandleDataSub(HGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(HCommand) = GetAddress(AddressOf HandleServerCommand)
    
    
End Sub

Function ReadHandleDataType(ByRef Data() As Byte) As Long
    Dim length As Long
    length = UBound(Data) - LBound(Data) - 4
    If length = -1 Then
        Call CopyMemory(ReadHandleDataType, Data(0), 4)
    ElseIf length >= 0 Then
        Call CopyMemory(ReadHandleDataType, Data(0), 4)
        Call CopyMemory(Data(0), Data(4), length + 1)
        ReDim Preserve Data(0 To length)
    End If
End Function

Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
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
    CallWindowProc HandleDataSub(MsgType), Index, Data, 0, 0
'Set buffer = Nothing
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long
    
    ' Check if elapsed time has passed
    Server(Index).DataBytes = Server(Index).DataBytes + DataLength
    If GetRealTickCount >= Server(Index).DataTimer Then
        Server(Index).DataTimer = GetRealTickCount + 1000
        Server(Index).DataBytes = 0
        Server(Index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(Index).GetData Buffer(), vbUnicode, DataLength
    Server(Index).Buffer.WriteBytes Buffer()
    
    If Server(Index).Buffer.length >= 4 Then
        pLength = Server(Index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= Server(Index).Buffer.length - 4
        If pLength <= Server(Index).Buffer.length - 4 Then
            Server(Index).DataPackets = Server(Index).DataPackets + 1
            Server(Index).Buffer.ReadLong
            HandleData Index, Server(Index).Buffer.ReadBytes(pLength)
            
            PacketsReceived = PacketsReceived + 1
        End If
        
        pLength = 0
        If Server(Index).Buffer.length >= 4 Then
            pLength = Server(Index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
    
    BytesReceived = BytesReceived + DataLength

    Server(Index).Buffer.Trim
End Sub

Private Sub HandleServerInfo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    With Server(Index)
        .CurrentPlayers = Buffer.ReadLong
        .MaxPlayers = Buffer.ReadLong
        .Name = Buffer.ReadString
        .Uptime = Buffer.ReadLong
        .Online = True
        .Port = Buffer.ReadLong

    AddLog "ServerInfo from: " & .Name & " Players: " & .CurrentPlayers & "/" & .MaxPlayers & " uptime: " & ConvertTime(GetRealTickCount - .Uptime)
    UpdateCaption
    UpdateComboList
    WriteServerFile
    End With
    Set Buffer = Nothing
End Sub

Private Sub HandleLog(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim text As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    text = Buffer.ReadString

    AddLog "[" & Server(Index).Name & "] " & text
    
    Set Buffer = Nothing
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

If frmServer.chkChat.Value = vbUnchecked Then Exit Sub

    Dim Buffer As New clsBuffer
    Set Buffer = New clsBuffer
    Dim msg As String
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    
    Buffer.Flush
    
    Buffer.WriteLong HGlobalMsg
    Buffer.WriteString msg
    
    AddLog "[" & Server(Index).Name & "] " & msg
    
    SendDataToAllHub Buffer.ToArray, Index

    Set Buffer = Nothing

End Sub

Private Sub HandleServerCommand(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim Buffer As New clsBuffer
    Set Buffer = New clsBuffer
    Dim Command As Long
    Dim strData As String
    Dim msg As String
    Buffer.WriteBytes Data()
    Command = Buffer.ReadLong
    strData = Buffer.ReadString
    
    Buffer.Flush
    
    Select Case Command
        Case Is = CommandsType.Classes
            msg = "Classes"
        Case Is = CommandsType.Maps
            msg = "Map #" & strData
        Case Is = CommandsType.Spells
            msg = "Spell #" & strData
        Case Is = CommandsType.Shops
            msg = "Shop #" & strData
        Case Is = CommandsType.npcs
            msg = "Npc #" & strData
        Case Is = CommandsType.Items
            msg = "Item #" & strData
        Case Is = CommandsType.Resources
            msg = "Resource #" & strData
        Case Is = CommandsType.Animations
            msg = "Animations"
        Case Is = CommandsType.Language
            msg = "Language"
        Case Is = CommandsType.SOptions
            msg = "Options"
        Case Is = CommandsType.SPets
            msg = "pet #" & strData
        Case Is = CommandsType.Weather
            msg = "Weather"
            If strData = "True" Then strData = 1 Else strData = 0
            If frmServer.cmbWeather.text = "NONE" Then Exit Sub
            If frmServer.cmbWeather.text <> "ALL" Then
                If frmServer.cmbWeather.text <> Server(Index).Name Then
                    AddLog "Ignoring weather update from: " & Server(Index).Name
                    Exit Sub
                End If
            End If
            
    End Select
    
    AddLog "Broadcasting " & msg & " update to all servers."
    
    
    Buffer.WriteLong HCommand
    Buffer.WriteLong Command
    Buffer.WriteString strData

    SendDataToAllHub Buffer.ToArray, Index

    Set Buffer = Nothing

End Sub
