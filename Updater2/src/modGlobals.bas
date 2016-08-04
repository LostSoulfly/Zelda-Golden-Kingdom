Attribute VB_Name = "modGlobals"
' Version Control Globals.
Public VersionCount As Long
Public CurVersion As Long

' Progress Tracking Globals
Public ProgressP As Double
Public UpToDate As Long
Public VersionsToGo As Long
Public PercentToGo As Long

' Locally stored variabled obtained with GetVar
Public NewsURL As String
Public GameName As String
Public GameWebsite As String
Public UpdateURL As String
Public ClientName As String

Public Type ServerRec
    Name As String
    CurrentPlayers As String
    MaxPlayers As String
    Online As Boolean
    Port As Long
End Type

Public Server(1 To 50) As ServerRec
Public SelectedServer As Integer
