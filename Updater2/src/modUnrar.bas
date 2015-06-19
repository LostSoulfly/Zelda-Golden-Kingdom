Attribute VB_Name = "modUnrar"
Option Explicit

Private Const ERAR_END_ARCHIVE As Byte = 10
Private Const ERAR_NO_MEMORY As Byte = 11
Private Const ERAR_BAD_DATA As Byte = 12
Private Const ERAR_BAD_ARCHIVE As Byte = 13
Private Const ERAR_UNKNOWN_FORMAT As Byte = 14
Private Const ERAR_EOPEN As Byte = 15
Private Const ERAR_ECREATE As Byte = 16
Private Const ERAR_ECLOSE As Byte = 17
Private Const ERAR_EREAD As Byte = 18
Private Const ERAR_EWRITE As Byte = 19
Private Const ERAR_SMALL_BUF As Byte = 20
 
Private Const RAR_OM_LIST As Byte = 0
Private Const RAR_OM_EXTRACT As Byte = 1
 
Private Const RAR_SKIP As Byte = 0
Private Const RAR_TEST As Byte = 1
Private Const RAR_EXTRACT As Byte = 2
 
Private Const RAR_VOL_ASK As Byte = 0
Private Const RAR_VOL_NOTIFY As Byte = 1

Public Enum RarOperations
    OP_EXTRACT = 0
    OP_TEST
    OP_LIST
End Enum
 
Private Type RARHeaderData
    ArcName As String * 260
    Filename As String * 260
    Flags As Long
    PackSize As Long
    UnpSize As Long
    HostOS As Long
    FileCRC As Long
    FileTime As Long
    UnpVer As Long
    Method As Long
    FileAttr As Long
    CmtBuf As String
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
End Type
 
Private Type RAROpenArchiveData
    ArcName As String
    OpenMode As Long
    OpenResult As Long
    CmtBuf As String
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
End Type
 
Private Declare Function RAROpenArchive Lib "unrar.dll" (ByRef ArchiveData As RAROpenArchiveData) As Long
Private Declare Function RARCloseArchive Lib "unrar.dll" (ByVal hArcData As Long) As Long
Private Declare Function RARReadHeader Lib "unrar.dll" (ByVal hArcData As Long, ByRef HeaderData As RARHeaderData) As Long
Private Declare Function RARProcessFile Lib "unrar.dll" (ByVal hArcData As Long, ByVal Operation As Long, ByVal DestPath As String, ByVal DestName As String) As Long
Private Declare Sub RARSetChangeVolProc Lib "unrar.dll" (ByVal hArcData As Long, ByVal Mode As Long)
Private Declare Sub RARSetPassword Lib "unrar.dll" (ByVal hArcData As Long, ByVal Password As String)

Public Sub RARExecute(ByVal Mode As RarOperations, ByVal RarFile As String, Optional ByVal Password As String)
    ' Description:-
    ' Extract file(s) from RAR archive.
    ' Parameters:-
    ' Mode = Operation to perform on RAR Archive
    ' RARFile = RAR Archive filename
    ' sPassword = Password (Optional)
    Dim lHandle As Long
    Dim iStatus As Integer
    Dim uRAR As RAROpenArchiveData
    Dim uHeader As RARHeaderData
    Dim sStat As String, Ret As Long
     
    uRAR.ArcName = RarFile
    uRAR.CmtBuf = Space(16384)
    uRAR.CmtBufSize = 16384
    
    If Mode = OP_LIST Then
        uRAR.OpenMode = RAR_OM_LIST
    Else
        uRAR.OpenMode = RAR_OM_EXTRACT
    End If
    
    lHandle = RAROpenArchive(uRAR)
    If uRAR.OpenResult <> 0 Then
        Kill RarFile
        OpenError uRAR.OpenResult, RarFile
    End If
 
    If Password <> "" Then RARSetPassword lHandle, Password
    
    If (uRAR.CmtState = 1) Then MsgBox uRAR.CmtBuf, vbApplicationModal + vbInformation, "Comment"
    
    iStatus = RARReadHeader(lHandle, uHeader)
    
    Do Until iStatus <> 0
        sStat = Left(uHeader.Filename, InStr(1, uHeader.Filename, vbNullChar) - 1)
        Select Case Mode
        Case RarOperations.OP_EXTRACT
            Ret = RARProcessFile(lHandle, RAR_EXTRACT, "", uHeader.Filename)
        Case RarOperations.OP_TEST
            Ret = RARProcessFile(lHandle, RAR_TEST, "", uHeader.Filename)
        Case RarOperations.OP_LIST
            Ret = RARProcessFile(lHandle, RAR_SKIP, "", "")
        End Select
        
        If Ret = 0 Then
            ProcessError Ret
        End If
        
        iStatus = RARReadHeader(lHandle, uHeader)
    Loop
    
    If iStatus = ERAR_BAD_DATA Then
        MsgBox "File header broken", vbCritical
        DestroyUpdater
    End If
    
    RARCloseArchive lHandle
End Sub

' Error handling
Private Sub OpenError(ErroNum As Long, ArcName As String)
Dim erro As String

    Select Case ErroNum
        Case ERAR_NO_MEMORY
            erro = "Not enough memory"
            GoTo errorbox
        Case ERAR_EOPEN:
            erro = "Cannot open " & ArcName
            GoTo errorbox
        Case ERAR_BAD_ARCHIVE:
            erro = ArcName & " is not RAR archive"
            GoTo errorbox
        Case ERAR_BAD_DATA:
            erro = ArcName & ": archive header broken"
            GoTo errorbox
    End Select
    
    Exit Sub
    
errorbox:
    MsgBox erro, vbCritical
    DestroyUpdater
End Sub

Private Sub ProcessError(ErroNum As Long)
Dim erro As String

    Select Case ErroNum
        Case ERAR_UNKNOWN_FORMAT
            erro = "Unknown archive format"
            GoTo errorbox
        Case ERAR_BAD_ARCHIVE:
            erro = "Bad volume"
            GoTo errorbox
        Case ERAR_ECREATE:
            erro = "File create error"
            GoTo errorbox
        Case ERAR_EOPEN:
            erro = "Volume open error"
            GoTo errorbox
        Case ERAR_ECLOSE:
            erro = "File close error"
            GoTo errorbox
        Case ERAR_EREAD:
            erro = "Read error"
            GoTo errorbox
        Case ERAR_EWRITE:
            erro = "Write error"
            GoTo errorbox
        Case ERAR_BAD_DATA:
            erro = "CRC error"
            GoTo errorbox
    End Select
    
    Exit Sub

errorbox:
    MsgBox erro, vbCritical
    DestroyUpdater
End Sub
