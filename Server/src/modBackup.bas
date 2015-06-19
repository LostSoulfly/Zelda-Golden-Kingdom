Attribute VB_Name = "modBackup"

Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)


Public Const DATA_PATH As String = "\Data"
Public Const CLIENT_DATA_PATH As String = "\Data Files"
Public Const ACCOUNTS_PATH As String = "\accounts"
Public Const BANKS_PATH As String = "\banks"
Public Const GUILDS_PATH As String = "\guilds"
Public Const GUILDNAMES_PATH As String = "\guildnames"
Public Const CODES_PATH As String = "\codes"


Sub HandleNeedAccounts(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    If GetPlayerAccess_Mode(index) < ADMIN_CREATOR Then Exit Sub
    
    Dim buffer As clsBuffer
    Dim password As String
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    password = buffer.ReadString
    
    If Not password = GetAccountsPassword Then
        Call GlobalMsg(GetPlayerName(index) & " " & GetTranslation(" ha sido expulsado de ") & " " & Options.Game_Name & " " & GetTranslation(" por el servidor!"), White)
        Call AddLog(0, "el servidor ha expulsado a " & GetPlayerName(index) & ".", ADMIN_LOG)
        Call AlertMsg(index, "Has sido expulsado")
        Exit Sub
    Else
        Call SendAllDirFiles(index, DATA_PATH & ACCOUNTS_PATH, CLIENT_DATA_PATH & ACCOUNTS_PATH, True)
        Call SendAllDirFiles(index, DATA_PATH & BANKS_PATH, CLIENT_DATA_PATH & BANKS_PATH, True)
        Call SendAllDirFiles(index, DATA_PATH & GUILDS_PATH, CLIENT_DATA_PATH & GUILDS_PATH, True)
        Call SendAllDirFiles(index, DATA_PATH & CODES_PATH, CLIENT_DATA_PATH & CODES_PATH, True)
    End If

End Sub

Public Sub SendAllDirFiles(ByVal index As Long, ByRef dir As String, ByRef ClientDir As String, ByVal Compress As Boolean)
    Dim buffer As clsBuffer
    Dim FileName As String
    
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSaveFiles
    
    Dim AccountsFolder As Folder
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    Set AccountsFolder = FSO.GetFolder(App.Path & dir)
    

    Dim Archivo As File
    
    
    
    buffer.WriteByte Compress
    
    Dim StuffBuffer As clsBuffer
    Set StuffBuffer = New clsBuffer
    
    StuffBuffer.WriteString ClientDir
    StuffBuffer.WriteLong AccountsFolder.Files.Count
    
    For Each Archivo In AccountsFolder.Files

    Dim Data() As Byte
        FileName = Archivo.Path
        Dim NotNull As Boolean
        Data = ReadFile(FileName, NotNull)
        StuffBuffer.WriteByte NotNull
        StuffBuffer.WriteString "\" & Archivo.Name
        If NotNull Then
            StuffBuffer.WriteLong UBound(Data) - LBound(Data) + 1
            StuffBuffer.WriteBytes Data
        End If
        
        DoEvents
    Next
    
    If Compress Then
        StuffBuffer.BufferCompress
    End If
    
    buffer.WriteBytes StuffBuffer.ToArray
     
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub


