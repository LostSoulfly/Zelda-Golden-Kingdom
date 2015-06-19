Attribute VB_Name = "modGlobals"
Option Explicit

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back npcs hp
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean

' Text vars
Public vbQuote As String

' Maximum classes
Public Max_Classes As Long

' Used for server loop
Public ServerOnline As Boolean

' Used for outputting text
Public NumLines As Long

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Public Secs As Long
Public TotalPlayersOnline As Long

' GameCPS
Public GameCPS As Long
Public ElapsedTime As Long

' high indexing
Public Player_HighIndex As Long

' lock the CPS?
Public CPSUnlock As Boolean

' Active weathers
Public Rainon As Boolean

' Time set for weather going (mili seconds)
Public WeatherTime As Long

' Used for calculate interval weather times
Public LastWeatherUpdate As Long

' Used for probabilities
Public WeatherProbability As Byte

'Sleep Time in ms
Public SleepTime As Long

'Map upper bound
Public Map_highindex As Long

'Server is bug
Public IsServerBug As Boolean

'Flood lapse
Public FloodTimer As Long

'Step Lapse
Public StepTimer As Long

'Test
Public TimeToMove As Long

'Chat Enabled
Public GlobalChatMinAccess As Long

'Stadistics
Public PacketsSent As Long
Public PacketsReceived As Long
Public BytesSent As Long
Public BytesReceived As Long


Public ByteCounter As Long
