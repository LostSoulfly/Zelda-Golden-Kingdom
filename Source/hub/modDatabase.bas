Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Function FileExist(ByVal Filename As String) As Boolean
    If LenB(Dir(Filename)) > 0 Then FileExist = True
End Function

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Sub WriteServerFile()
Dim File As String, Header As String
Dim i As Integer
Dim NumServers As Integer
File = App.Path & "\Zelda\Status.txt"
Header = "Servers"
NumServers = TotalServers

If FileExist(File) Then Kill File
DoEvents

PutVar File, Header, "NumServers", CStr(NumServers)

For i = 1 To NumServers

Header = "Server" & i

    PutVar File, Header, "Players", CStr(Server(i).CurrentPlayers)
    PutVar File, Header, "MaxPlayers", CStr(Server(i).MaxPlayers)
    PutVar File, Header, "PvPOnly", "0"
    PutVar File, Header, "Name", Server(i).Name
    PutVar File, Header, "Port", CStr(Server(i).Port)
    PutVar File, Header, "Online", " 1" 'IIf(Server(i).Online, "1", "0")
Next i


End Sub
