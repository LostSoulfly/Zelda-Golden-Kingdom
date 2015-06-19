Attribute VB_Name = "modGuild"
Public Const MAX_GUILD_MEMBERS As Long = 50
Public Const MAX_GUILD_RANKS As Long = 6
Public Const MAX_GUILD_RANKS_PERMISSION As Long = 6


Public GuildData As GuildRec

Public Type GuildRanksRec
    'General variables
    Used As Boolean
    Name As String
    
    'Rank Variables
    RankPermission(1 To MAX_GUILD_RANKS_PERMISSION) As Byte
    RankPermissionName(1 To MAX_GUILD_RANKS_PERMISSION) As String
End Type

Public Type GuildMemberRec
    'User login/name
    Used As Boolean
    
    User_Login As String
    User_Name As String
    Founder As Boolean
    
    'Guild Variables
    Rank As Integer
    Comment As String * 100
     
End Type

Public Type GuildRec
    In_Use As Boolean
    
    Guild_Name As String
    
    'Guild file number for saving
    Guild_Fileid As Long
    
    Guild_Members(1 To MAX_GUILD_MEMBERS) As GuildMemberRec
    Guild_Ranks(1 To MAX_GUILD_RANKS) As GuildRanksRec
    
    'Message of the day
    Guild_MOTD As String * 100
    
    'The rank recruits start at
    Guild_RecruitRank As Integer
    Guild_Color As Integer

End Type
Public Sub HandleAdminGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Integer
Dim b As Integer

    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    If Buffer.ReadByte = 1 Then
        frmGuildAdmin.Visible = True
    Else
        frmGuildAdmin.Visible = False
    End If

    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAdminGuild", "modGuild", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
Public Sub HandleSendGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Integer
Dim b As Integer

    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    GuildData.Guild_Name = Buffer.ReadString
    GuildData.Guild_Color = Buffer.ReadInteger
    GuildData.Guild_MOTD = Buffer.ReadString
    GuildData.Guild_RecruitRank = Buffer.ReadInteger
    
    'Get Members
    For i = 1 To MAX_GUILD_MEMBERS
        GuildData.Guild_Members(i).User_Name = Buffer.ReadString
        GuildData.Guild_Members(i).Rank = Buffer.ReadInteger
        GuildData.Guild_Members(i).Comment = Buffer.ReadString
    Next i
    
    'Get Ranks
    For i = 1 To MAX_GUILD_RANKS
        GuildData.Guild_Ranks(i).Name = Buffer.ReadString
        For b = 1 To MAX_GUILD_RANKS_PERMISSION
            GuildData.Guild_Ranks(i).RankPermission(b) = Buffer.ReadByte
            GuildData.Guild_Ranks(i).RankPermissionName(b) = Buffer.ReadString
        Next b
    Next i
    
    'Update Guildadmin data
    Call frmGuildAdmin.Load_Guild_Admin
    
    
    Set Buffer = Nothing
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSendGuild", "modGuild", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
Public Sub GuildMsg(ByVal text As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayGuild
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GuildMsg", "modGuild", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GuildCommand(ByVal Command As Integer, ByVal SendText As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CGuildCommand
    Buffer.WriteInteger Command
    Buffer.WriteString SendText
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GuildMsg", "modGuild", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GuildSave(ByVal SaveType As Integer, ByVal Index As Integer)
Dim Buffer As clsBuffer
Dim i As Integer
Dim b As Integer
'SaveType
'1=options
'2=users
'3=ranks
 If Index = 0 Then Exit Sub


    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSaveGuild
    
    Buffer.WriteInteger SaveType
    Buffer.WriteInteger Index
    
    Select Case SaveType
    Case 1
        'options
        Buffer.WriteInteger GuildData.Guild_Color
        Buffer.WriteInteger GuildData.Guild_RecruitRank
        Buffer.WriteString GuildData.Guild_MOTD
    Case 2
        'users
        Buffer.WriteInteger GuildData.Guild_Members(Index).Rank
        Buffer.WriteString GuildData.Guild_Members(Index).Comment
    Case 3
        'ranks
        Buffer.WriteString GuildData.Guild_Ranks(Index).Name
        For i = 1 To MAX_GUILD_RANKS_PERMISSION
            Buffer.WriteByte GuildData.Guild_Ranks(Index).RankPermission(i)
        Next i
    End Select

    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GuildMsg", "modGuild", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleGuildData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteBytes Data()

If Buffer.ReadByte = 1 Then
    Player(MyIndex).GuildName = Buffer.ReadString
    Player(MyIndex).GuildMemberId = Buffer.ReadLong
Else
    Player(MyIndex).GuildName = vbNullString
    Player(MyIndex).GuildMemberId = 0
End If

End Sub
