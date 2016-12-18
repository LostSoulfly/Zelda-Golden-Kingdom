Attribute VB_Name = "modGuild"
Option Explicit
'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

'Max Members Per Guild
Public Const MAX_GUILD_MEMBERS As Long = 50
'Max Ranks Guilds Can Have
Public Const MAX_GUILD_RANKS As Long = 6
'Max Different Permissions
Public Const MAX_GUILD_RANKS_PERMISSION As Long = 6
'Max guild save files(aka max guilds)
Public Const MAX_GUILD_SAVES As Long = 200
'Guild make price
Public Const GUILD_PRICE As Long = 100

Public Const GUILD_PATH As String = "\guilds"
Public Const GUILD_NAMES_FILE As String = "\guildnames.txt"

'Default Ranks Info
'1: Open Admin
'2: Can Recruit
'3: Can Kick
'4: Can Edit Ranks
'5: Can Edit Users
'6: Can Edit Options

Public Guild_Ranks_Premission_Names(1 To MAX_GUILD_RANKS_PERMISSION) As String
Public Default_Ranks(1 To MAX_GUILD_RANKS_PERMISSION) As Byte


'Max is set to MAX_PLAYERS so each online player can have his own guild
Public GuildData(1 To MAX_PLAYERS) As GuildRec

Public Type GuildRanksRec
    'General variables
    Used As Boolean
    Name As String
    
    'Rank Variables
    RankPermission(1 To MAX_GUILD_RANKS_PERMISSION) As Byte
End Type

Public Type GuildMemberRec
    'User login/name
    Used As Boolean
    
    User_Login As String
    User_Name As String
    Founder As Boolean
    
    Online As Boolean
    
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
    'Color of guild name
    Guild_Color As Integer

End Type
Public Sub Set_Default_Guild_Ranks()
    'Max sure this starts at 1 and ends at MAX_GUILD_RANKS_PERMISSION (Default 7)
    '0 = Cannot, 1 = Able To
    Guild_Ranks_Premission_Names(1) = "Open Admin"
    Default_Ranks(1) = 0
    
    Guild_Ranks_Premission_Names(2) = "Can Recruit"
    Default_Ranks(2) = 1
    
    Guild_Ranks_Premission_Names(3) = "Can Kick"
    Default_Ranks(3) = 0
    
    Guild_Ranks_Premission_Names(4) = "Can Edit Ranks"
    Default_Ranks(4) = 0
    
    Guild_Ranks_Premission_Names(5) = "Can Edit Users"
    Default_Ranks(5) = 0
    
    Guild_Ranks_Premission_Names(6) = "Can Edit Options"
    Default_Ranks(6) = 0
End Sub
Public Function GuildCheckName(index As Long, MemberSlot As Long, AttemptCorrect As Boolean) As Boolean
Dim i As Integer

    If player(index).GuildFileId = 0 Or TempPlayer(index).tmpGuildSlot = 0 Or IsPlaying(index) = False Or MemberSlot = 0 Then
        GuildCheckName = False
        Exit Function
    End If
    
    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(MemberSlot).User_Login = player(index).login Then
        GuildCheckName = True
        Exit Function
    End If
    
    If AttemptCorrect = True Then
        If TempPlayer(index).tmpGuildSlot > 0 And player(index).GuildFileId > 0 Then
            'did they get moved?
            For i = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).User_Login = player(index).login Then
                    player(index).GuildMemberId = i
                    Call SavePlayer(index)
                    GuildCheckName = True
                    Exit Function
                Else
                    player(index).GuildMemberId = 0
                End If
            Next i
                
            'Remove from guild if we can't find them
            If player(index).GuildMemberId = 0 Then
                player(index).GuildFileId = 0
                TempPlayer(index).tmpGuildSlot = 0
                Call SavePlayer(index)
                PlayerMsg index, "No pudimos encontrarte en la lista de tu clan", BrightRed
                PlayerMsg index, "1)Te han echado   2)Tu clan ha sido sobreescribido", BrightRed
            End If
        End If
    End If
    
    
    GuildCheckName = False


End Function
Public Sub MakeGuild(Founder_Index As Long, Name As String)
    Dim tmpGuild As GuildRec
    Dim GuildSlot As Long
    Dim GuildFileId As Long
    Dim i As Integer
    Dim b As Integer
    Dim ItemAmount As Long
    Dim N As String

    If player(Founder_Index).GuildFileId > 0 Then
        PlayerMsg Founder_Index, "Debes abandonar tu clan actual para poder crear otro!", BrightRed
        Exit Sub
    End If
    
    If Len(Name) > 30 Then
        PlayerMsg Founder_Index, "El nombre no puede tener mas de 30 caracteres!", BrightRed
        Exit Sub
    End If
    
    ' Prevent hacking
    For i = 1 To Len(Name)
        N = AscW(Mid$(Name, i, 1))

        If Not isNameLegal(N) Then
            PlayerMsg Founder_Index, "Nombre inválido, solo letras, números, espacios y _ están permitidos.", BrightRed
            Exit Sub
        End If

    Next

    If GuildNameBlacklist(Name) And GetPlayerAccess_Mode(Founder_Index) <= ADMIN_MAPPER Then
        PlayerMsg Founder_Index, "That guild name has been blacklisted!", BrightRed, True, False
        Exit Sub
    End If

    If GuildNameExist(Name) Then
        PlayerMsg Founder_Index, "Nombre ya existente", BrightRed
        Exit Sub
    End If
    
    
    
    GuildFileId = Find_Guild_Save
    GuildSlot = FindOpenGuildSlot
    
    If Not IsPlaying(Founder_Index) Then Exit Sub

If Not GetPlayerAccess_Mode(Founder_Index) > ADMIN_MAPPER Then

    'We are unable for an unknown reason
    If GuildSlot = 0 Or GuildFileId = 0 Then
        PlayerMsg Founder_Index, "¡Incapaz de crear clan!", BrightRed
        Exit Sub
    End If
    
    'Change 10 to any level you want
    If GetPlayerLevel(Founder_Index) < 15 Then
        PlayerMsg Founder_Index, "Nivel Insuficiente. Debes tener al menos 15.", BrightGreen
        Exit Sub
    End If
    
    
    If LenB(Name$) = 0 Then
        PlayerMsg Founder_Index, "¡Tu clan necesita un nombre!", BrightRed
        Exit Sub
    End If
    
    'Change 1 to item number
    ItemAmount = HasItem(Founder_Index, 1)
    
    'Change 500 to amount
    If ItemAmount = 0 Or ItemAmount < GUILD_PRICE Then
        PlayerMsg Founder_Index, GetTranslation("Rupias insuficientes. Debes tener al menos", , UnTrimBack) & GUILD_PRICE & GetTranslation("Rupias Verdes", , UnTrimFront), BrightRed, , False
        Exit Sub
    End If
    
    'Change 1 to item number 5000 to amount
    TakeInvItem Founder_Index, 1, GUILD_PRICE

End If
    
    GuildData(GuildSlot).Guild_Name = Name
    GuildData(GuildSlot).In_Use = True
    GuildData(GuildSlot).Guild_Fileid = GuildFileId
    GuildData(GuildSlot).Guild_Members(1).Founder = True
    GuildData(GuildSlot).Guild_Members(1).User_Login = player(Founder_Index).login
    GuildData(GuildSlot).Guild_Members(1).User_Name = player(Founder_Index).Name
    GuildData(GuildSlot).Guild_Members(1).Rank = MAX_GUILD_RANKS
    GuildData(GuildSlot).Guild_Members(1).Comment = "Founder of Clan"
    GuildData(GuildSlot).Guild_Members(1).Used = True
    GuildData(GuildSlot).Guild_Members(1).Online = True
    

    'Set up Admin Rank with all permission which is just the max rank
    GuildData(GuildSlot).Guild_Ranks(MAX_GUILD_RANKS).Name = "Leader"
    GuildData(GuildSlot).Guild_Ranks(MAX_GUILD_RANKS).Used = True
    
    For b = 1 To MAX_GUILD_RANKS_PERMISSION
        GuildData(GuildSlot).Guild_Ranks(MAX_GUILD_RANKS).RankPermission(b) = 1
    Next b
    
    'Set up rest of the ranks with default permission
    For i = 1 To MAX_GUILD_RANKS - 1
        GuildData(GuildSlot).Guild_Ranks(i).Name = "Grado " & i
        GuildData(GuildSlot).Guild_Ranks(i).Used = True
        
        For b = 1 To MAX_GUILD_RANKS_PERMISSION
            GuildData(GuildSlot).Guild_Ranks(i).RankPermission(b) = Default_Ranks(b)
        Next b

    Next i
    
    player(Founder_Index).GuildFileId = GuildFileId
    player(Founder_Index).GuildMemberId = 1
    TempPlayer(Founder_Index).tmpGuildSlot = GuildSlot
    
    
    'Save
    Call SaveGuild(GuildSlot)
    Call SavePlayer(Founder_Index)
    Call SaveGuildName(Name)
    
    'Send to player
    Call SendGuild(False, Founder_Index, GuildSlot)
    
    'Inform users
    PlayerMsg Founder_Index, "¡Clan creado satisfactoriamente!", BrightGreen
    PlayerMsg Founder_Index, GetTranslation("¡Bienvenido a", , UnTrimBack) & GuildData(GuildSlot).Guild_Name & ".", BrightGreen, , False
    
    PlayerMsg Founder_Index, "Puedes hablar en el chat del clan seleccionando desde Opciones Chat: Clan", BrightGreen
    
    'Update user for guild name display
    'Call SendPlayerData(Founder_Index)
    Call SendPlayerGuildData(Founder_Index)

    
End Sub
Public Function CheckGuildPermission(index As Long, Permission As Integer) As Boolean
Dim GuildSlot As Long

    'Get slot
    GuildSlot = TempPlayer(index).tmpGuildSlot
    
    'Make sure we are talking about the same person
    If Not GuildData(GuildSlot).Guild_Members(player(index).GuildMemberId).User_Login = player(index).login Then
        'Something went wrong and they are not allowed to do anything
        CheckGuildPermission = False
        Exit Function
    End If
    
    'If founder, true in every case
    If GuildData(GuildSlot).Guild_Members(player(index).GuildMemberId).Founder = True Then
        CheckGuildPermission = True
        Exit Function
    End If
    
    'Make sure this slot is being used aka they are still a member
    If GuildData(GuildSlot).Guild_Members(player(index).GuildMemberId).Used = False Then
        'Something went wrong and they are not allowed to do anything
        CheckGuildPermission = False
        Exit Function
    End If
    
    'Check if they are able to
    If GuildData(GuildSlot).Guild_Ranks(GuildData(GuildSlot).Guild_Members(player(index).GuildMemberId).Rank).RankPermission(Permission) = 1 Then
        CheckGuildPermission = True
    Else
        CheckGuildPermission = False
    End If
    
End Function
Public Sub Request_Guild_Invite(index As Long, GuildSlot As Long, Inviter_Index As Long)

    If player(index).GuildFileId > 0 Then
        PlayerMsg index, GetTranslation("¡Debe abandonar su clan actual antes de poder unirse a", , UnTrimBack) & GuildData(GuildSlot).Guild_Name & "!", BrightRed, , False
        PlayerMsg Inviter_Index, "¡No pueden unirse al clan porque ya están en otro!", BrightRed
        Exit Sub
    End If

    If TempPlayer(index).tmpGuildInviteSlot > 0 Then
        PlayerMsg Inviter_Index, "Éste usuario tiene una invitación pendiente, vuelva a intentarlo.", BrightRed
        Exit Sub
    End If

    'Permission 2 = Can Recruit
    If CheckGuildPermission(Inviter_Index, 2) = False Then
        PlayerMsg Inviter_Index, "¡Tu grado no es suficientemente alto!", BrightRed
        Exit Sub
    End If
    
    TempPlayer(index).tmpGuildInviteSlot = GuildSlot
    '2 minute
    TempPlayer(index).tmpGuildInviteTimer = GetRealTickCount + 120000
    
    TempPlayer(index).tmpGuildInviteId = player(Inviter_Index).GuildFileId
    
    PlayerMsg Inviter_Index, "¡Invitación al Clan enviada!", Green
    PlayerMsg index, player(Inviter_Index).Name & GetTranslation("te ha invitado al clan", , UnTrimBoth) & GuildData(GuildSlot).Guild_Name & "!", Green, , False
    PlayerMsg index, "Para aceptarla, pulsa en Aceptar desde el Panel del Clan, el cual se abre desde Opciones, antes de 2 minutos.", Green
    PlayerMsg index, "O bien pulsa en Rechazar, para rechazar la oferta.", Green
End Sub
Public Sub Join_Guild(index As Long, GuildSlot As Long)
Dim OpenSlot As Long



    If IsPlaying(index) = False Then Exit Sub
    
    OpenSlot = FindOpenGuildMemberSlot(GuildSlot)
        'Guild full?
        If OpenSlot > 0 Then
            'Set guild data
            GuildData(GuildSlot).Guild_Members(OpenSlot).Used = True
            GuildData(GuildSlot).Guild_Members(OpenSlot).User_Login = player(index).login
            GuildData(GuildSlot).Guild_Members(OpenSlot).User_Name = player(index).Name
            GuildData(GuildSlot).Guild_Members(OpenSlot).Rank = GuildData(GuildSlot).Guild_RecruitRank
            GuildData(GuildSlot).Guild_Members(OpenSlot).Comment = "Unido: " & DateValue(Now)
            GuildData(GuildSlot).Guild_Members(OpenSlot).Online = True
            
            'Set player data
            player(index).GuildFileId = GuildData(GuildSlot).Guild_Fileid
            player(index).GuildMemberId = OpenSlot
            TempPlayer(index).tmpGuildSlot = GuildSlot
            
            'Save
            Call SaveGuild(GuildSlot)
            Call SavePlayer(index)
            
            'Send player guild data and display welcome
            Call SendGuild(False, index, GuildSlot)
            PlayerMsg index, GetTranslation("Bienvenido a", , UnTrimBack) & GuildData(GuildSlot).Guild_Name & ".", BrightGreen, , False
            
            PlayerMsg index, "Puedes hablar en el chat del clan seleccionando desde Opciones Chat: Clan", BrightGreen
            
            'Update player to display guild name
            'Call SendPlayerData(index)
            Call SendPlayerGuildData(index)
            
        Else
            'Guild full display msg
            PlayerMsg index, "¡El clan está lleno!", BrightRed
        End If
    
        
    
End Sub
Public Function Find_Guild_Save() As Long
Dim FoundSlot As Boolean
Dim Current As Integer
FoundSlot = False
Current = 1

Do Until FoundSlot = True
    
    If Not FileExist("\Data\guilds\Guild" & Current & ".dat") Then
        Find_Guild_Save = Current
        FoundSlot = True
    Else
        Current = Current + 1
    End If
    
    'Max Guild Files check
    If Current > MAX_GUILD_SAVES Then
        'send back 0 for no slot found
        Find_Guild_Save = 0
        FoundSlot = True
    End If
    
    
Loop

End Function
Public Function FindOpenGuildSlot() As Long
    Dim i As Integer
    
    For i = 1 To MAX_PLAYERS
        If GuildData(i).In_Use = False Then
            FindOpenGuildSlot = i
            Exit Function
        End If
        
        'No slot found how?
        FindOpenGuildSlot = 0
    Next i
End Function
Public Function FindOpenGuildMemberSlot(GuildSlot As Long) As Long
Dim i As Integer
    
    For i = 1 To MAX_GUILD_MEMBERS
        If GuildData(GuildSlot).Guild_Members(i).Used = False Then
            FindOpenGuildMemberSlot = i
            Exit Function
        End If
    Next i
    
    'Guild is full sorry bub
    FindOpenGuildMemberSlot = 0

End Function
Public Sub ClearGuildMemberSlot(GuildSlot As Long, MembersSlot As Long)
            GuildData(GuildSlot).Guild_Members(MembersSlot).Used = False
            GuildData(GuildSlot).Guild_Members(MembersSlot).User_Login = vbNullString
            GuildData(GuildSlot).Guild_Members(MembersSlot).User_Name = vbNullString
            GuildData(GuildSlot).Guild_Members(MembersSlot).Rank = 0
            GuildData(GuildSlot).Guild_Members(MembersSlot).Comment = vbNullString
            GuildData(GuildSlot).Guild_Members(MembersSlot).Founder = False
            GuildData(GuildSlot).Guild_Members(MembersSlot).Online = False
            
            'Save guild after we remove member
            Call SaveGuild(GuildSlot)
End Sub
Public Sub LoadGuild(GuildSlot As Long, GuildFileId As Long)
Dim i As Integer
'If 0 something is wrong
If GuildFileId = 0 Then Exit Sub

'Does this file even exist?
If Not FileExist("\Data\guilds\Guild" & GuildFileId & ".dat") Then Exit Sub

    Dim FileName As String
    Dim F As Long
    
        FileName = App.Path & "\data\guilds\Guild" & GuildFileId & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , GuildData(GuildSlot)
        Close #F
        
        GuildData(GuildSlot).In_Use = True
        
        'Make sure an online flag didn't manage to slip through
        For i = 1 To MAX_GUILD_MEMBERS
            If GuildData(GuildSlot).Guild_Members(i).Online = True Then
                GuildData(GuildSlot).Guild_Members(i).Online = False
            End If
        Next i
        
End Sub
Public Sub SaveGuild(GuildSlot As Long)

'Dont save unless a fileid was assigned
If GuildData(GuildSlot).Guild_Fileid = 0 Then Exit Sub

    If IsServerBug Then Exit Sub

    Dim FileName As String
    Dim F As Long
    
    FileName = App.Path & "\data\guilds\Guild" & GuildData(GuildSlot).Guild_Fileid & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , GuildData(GuildSlot)
    Close #F
    
End Sub
Public Sub UnloadGuildSlot(GuildSlot As Long)
    'Exit on error
    If GuildSlot = 0 Or GuildSlot > MAX_GUILD_SAVES Then Exit Sub
    If GuildData(GuildSlot).In_Use = False Then Exit Sub
    
    'Save it first
    Call SaveGuild(GuildSlot)
    'Clear and reset for next use
    Call ClearGuild(GuildSlot)
End Sub
Public Sub ClearGuilds()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        Call ClearGuild(i)
    Next i
End Sub
Public Sub ClearGuild(index As Long)
    Call ZeroMemory(ByVal VarPtr(GuildData(index)), LenB(GuildData(index)))
    GuildData(index).Guild_Name = vbNullString
    GuildData(index).In_Use = False
    GuildData(index).Guild_Fileid = 0
    GuildData(index).Guild_Color = 1
    GuildData(index).Guild_RecruitRank = 1
End Sub
Public Sub CheckUnloadGuild(GuildSlot As Long)
Dim i As Integer
Dim UnloadGuild As Boolean

UnloadGuild = True

If GuildData(GuildSlot).In_Use = False Then Exit Sub

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If player(i).GuildFileId = GuildData(GuildSlot).Guild_Fileid Then
                UnloadGuild = False
                Exit For
            End If
        End If
    Next i
    
    
    If UnloadGuild = True Then
        Call UnloadGuildSlot(GuildSlot)
    End If
End Sub
Public Sub GuildKick(GuildSlot As Long, index As Long, playerName As String)

Dim FoundOffline As Boolean
Dim IsOnline As Boolean
Dim OnlineIndex As Long
Dim MemberSlot As Long
Dim i As Integer
    
    
    
    OnlineIndex = FindPlayer(playerName)
    
    If OnlineIndex = index Then
        PlayerMsg index, "¡No puedes expulsarte a ti mismo!", BrightRed
        Exit Sub
    End If
    
    
    
    'If OnlineIndex > 0 they are online
    If OnlineIndex > 0 Then
        IsOnline = True
        
        If player(OnlineIndex).GuildMemberId > 0 Then
            MemberSlot = player(OnlineIndex).GuildMemberId
        Else
            'Prevent error, rest of this code assumes this is greater than 0
            Exit Sub
        End If
        
    Else
        IsOnline = False
    End If
    
    
    'Handle kicking online user
    If IsOnline = True Then
        If Not player(index).GuildFileId = player(OnlineIndex).GuildFileId Then
            PlayerMsg index, "¡El usuario debe estar en tu clan para poder expulsarlo!", BrightRed
            Exit Sub
        End If
        
        If GuildData(GuildSlot).Guild_Members(MemberSlot).Founder = True Then
            PlayerMsg index, "¡No puedes expulsar al fundador!", BrightRed
            Exit Sub
        End If
        
        player(OnlineIndex).GuildFileId = 0
        player(OnlineIndex).GuildMemberId = 0
        TempPlayer(OnlineIndex).tmpGuildSlot = 0
        Call ClearGuildMemberSlot(GuildSlot, MemberSlot)
        PlayerMsg OnlineIndex, "¡Has sido expulsado del clan!", BrightRed
        PlayerMsg index, "¡Jugador expulsado!", BrightRed
        Call SavePlayer(OnlineIndex)
        Call SaveGuild(GuildSlot)
        'Call SendPlayerData(OnlineIndex)
        Call SendPlayerGuildData(OnlineIndex)
        Exit Sub
    End If
    
    
    
    'Handle Kicking Offline User
    FoundOffline = False
    If IsOnline = False Then
        'Lets Try to find them in the roster
        For i = 1 To MAX_GUILD_MEMBERS
            If playerName = Trim$(GuildData(GuildSlot).Guild_Members(i).User_Name) Then
                'Found them
                FoundOffline = True
                MemberSlot = i
                Exit For
            End If
        Next i
        
        If FoundOffline = True Then
        
            If MemberSlot = 0 Then Exit Sub
            
            Call ClearGuildMemberSlot(GuildSlot, MemberSlot)
            Call SaveGuild(GuildSlot)
            PlayerMsg index, "Jugador fuera de línea expulsado!", BrightRed
            Exit Sub
        End If
        
        If FoundOffline = False And IsOnline = False Then
            PlayerMsg index, "No se pudo encontrar a " & playerName & " en tu clan.", BrightRed
        End If
    
    End If
 
End Sub
Public Sub GuildLeave(index As Long)
Dim i As Integer
    
    'This is for the leave command only, kicking has its own sub because it handles both online and offline kicks, while this only handles online.
    
    If Not player(index).GuildFileId > 0 Then
        PlayerMsg index, "¡Debes estar en un clan para abandonarlo!", BrightRed
        Exit Sub
    End If
    
    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(player(index).GuildMemberId).Founder = True Then
        PlayerMsg index, "El fundador no puede abandonar o ser expulsado, antes de eso el estatus de fundador debe ser transferido.", BrightRed
        PlayerMsg index, "Utiliza /founder (nombre) para transferir el título.", BrightRed
        Exit Sub
    End If
    
    'They match so they can leave
    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(player(index).GuildMemberId).User_Login = player(index).login Then
        
        'Clear guild slot
        Call ClearGuildMemberSlot(TempPlayer(index).tmpGuildSlot, player(index).GuildMemberId)
        
        'Clear player data
        player(index).GuildFileId = 0
        player(index).GuildMemberId = 0
        TempPlayer(index).tmpGuildSlot = 0
        
        'Update user for guild name display
        'Call SendPlayerData(index)
        Call SendPlayerGuildData(index)
        
        PlayerMsg index, "Has abandonado el clan.", BrightRed

    Else
        'They don't match this slot remove them
        player(index).GuildFileId = 0
        player(index).GuildMemberId = 0
        TempPlayer(index).tmpGuildSlot = 0
    End If
    
    
End Sub
Public Sub GuildLoginCheck(index As Long)
Dim i As Long
Dim GuildSlot As Long
Dim GuildLoaded As Boolean
GuildLoaded = False


    'Not in guild
    If player(index).GuildFileId = 0 Then Exit Sub
    
    'Check to make sure the guild file exists
    If Not FileExist("\Data\guilds\Guild" & player(index).GuildFileId & ".dat") Then
        'If guild was deleted remove user from guild
        player(index).GuildFileId = 0
        player(index).GuildMemberId = 0
        TempPlayer(index).tmpGuildSlot = 0
        Call SavePlayer(index)
        PlayerMsg index, "Tu clan ha sido borrado!", BrightRed
        Exit Sub
    End If
    
    'First we need to see if our guild is loaded
    For i = 1 To MAX_PLAYERS
        If GuildData(i).In_Use = True Then
            'If its already loaded set true
            If GuildData(i).Guild_Fileid = player(index).GuildFileId Then
                GuildLoaded = True
                GuildSlot = i
                Exit For
            End If
        End If
    Next i
    
    'If the guild is not loaded we need to load it
    If GuildLoaded = False Then
        'Find open guild slot, if 0 none
        GuildSlot = FindOpenGuildSlot
        If GuildSlot > 0 Then
            'LoadGuild
            Call LoadGuild(GuildSlot, player(index).GuildFileId)
            
        End If
    End If
    
    'Set GuildSlot
    TempPlayer(index).tmpGuildSlot = GuildSlot
    
    'This is to prevent errors when we look for them
    If player(index).GuildMemberId = 0 Then player(index).GuildMemberId = 1

    'Make sure user didn't get kicked or guild was replaced by a different guild, both result in removal
    If GuildCheckName(index, player(index).GuildMemberId, True) = False Then
        'unload if this user is not in this guild and it was loaded for this user
        If GuildLoaded = False Then
            Call UnloadGuildSlot(GuildSlot)
            Exit Sub
        End If
    End If
    
    'Sent data and set slot if all is good
    If player(index).GuildFileId > 0 Then
        'Set online flag
        GuildData(GuildSlot).Guild_Members(player(index).GuildMemberId).Online = True
        
        
        'send
        Call SendGuild(False, index, GuildSlot)
        
        'Display motd
        If Not GuildData(GuildSlot).Guild_MOTD = vbNullString Then
            PlayerMsg index, "Guild Motd: " & GuildData(GuildSlot).Guild_MOTD, Cyan, , False
        End If
    End If
    
    
    
End Sub
Sub DisbandGuild(GuildSlot As Long, index As Long)
Dim i As Integer
Dim tmpGuildSlot As Long
Dim TmpGuildFileId As Long
Dim FileName As String

'Set some thing we need
tmpGuildSlot = GuildSlot
TmpGuildFileId = GuildData(tmpGuildSlot).Guild_Fileid

    'They are who they say they are, and are founder
    If GuildCheckName(index, player(index).GuildMemberId, False) = True And GuildData(tmpGuildSlot).Guild_Members(player(index).GuildMemberId).Founder = True Then
        'File exists right?
         If FileExist("\Data\Guilds\Guild" & TmpGuildFileId & ".dat") = True Then
            'We have a go for disband
            'First we take everyone online out, this will include the founder people who login later will be kicked out then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) = True Then
                    If player(i).GuildFileId = TmpGuildFileId Then
                        'remove from guild
                        player(i).GuildFileId = 0
                        player(i).GuildMemberId = 0
                        TempPlayer(i).tmpGuildSlot = 0
                        Call SavePlayer(i)
                        'Send player data so they don't have name over head anymore
                        'Call SendPlayerData(i)
                        Call SendPlayerGuildData(i)
                    End If
                End If
            Next i
            
            FileName = App.Path & "\Data\Guilds\Guild" & TmpGuildFileId & ".dat"
            Kill FileName
            
            
            
            DeleteGuildName GuildData(tmpGuildSlot).Guild_Name
            
            'Unload Guild from memory
            Call UnloadGuildSlot(tmpGuildSlot)

            
            
            
            PlayerMsg index, "Clan desbancado!", BrightGreen
         End If
    Else
        PlayerMsg index, "No puedes hacer eso!", BrightRed
    End If
End Sub
Sub SendDataToGuild(ByVal GuildSlot As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If player(i).GuildFileId = GuildData(GuildSlot).Guild_Fileid Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendGuild(ByVal SendToWholeGuild As Boolean, ByVal index As Long, ByVal GuildSlot As Long)
    Dim Buffer As clsBuffer
    Dim i As Integer
    Dim b As Integer
    
    If GuildSlot < 1 Or GuildSlot > MAX_PLAYERS Then Exit Sub

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SSendGuild
    
    'General data
    Buffer.WriteString GuildData(GuildSlot).Guild_Name
    Buffer.WriteInteger GuildData(GuildSlot).Guild_Color
    Buffer.WriteString GuildData(GuildSlot).Guild_MOTD
    Buffer.WriteInteger GuildData(GuildSlot).Guild_RecruitRank
    
    
    'Send Members
    For i = 1 To MAX_GUILD_MEMBERS
        Buffer.WriteString GuildData(GuildSlot).Guild_Members(i).User_Name
        Buffer.WriteInteger GuildData(GuildSlot).Guild_Members(i).Rank
        Buffer.WriteString GuildData(GuildSlot).Guild_Members(i).Comment
    Next i
    
    'Send Ranks
    For i = 1 To MAX_GUILD_RANKS
            Buffer.WriteString GuildData(GuildSlot).Guild_Ranks(i).Name
        For b = 1 To MAX_GUILD_RANKS_PERMISSION
            Buffer.WriteByte GuildData(GuildSlot).Guild_Ranks(i).RankPermission(b)
            Buffer.WriteString Guild_Ranks_Premission_Names(b)
        Next b
    Next i
    
    If SendToWholeGuild = False Then
        SendDataTo index, Buffer.ToArray()
    Else
        SendDataToGuild GuildSlot, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub
Sub ToggleGuildAdmin(ByVal index As Long, ByVal GuildSlot, ByVal OpenAdmin As Boolean)
    Dim Buffer As clsBuffer
    Dim i As Integer
    Dim b As Integer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAdminGuild
    
    
    If OpenAdmin = True Then
        Buffer.WriteByte 1
    Else
        Buffer.WriteByte 0
    End If

        SendDataTo index, Buffer.ToArray()

    
    Set Buffer = Nothing
End Sub
Sub SayMsg_Guild(ByVal GuildSlot As Long, ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess_Mode(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString message
    Buffer.WriteString "[" & GuildData(GuildSlot).Guild_Name & "]"
    Buffer.WriteLong saycolour
    Buffer.WriteLong ClanChat
    
    SendDataToGuild GuildSlot, Buffer.ToArray()

    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Party(ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess_Mode(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString message
    Buffer.WriteString "[Party]"
    Buffer.WriteLong saycolour
    Buffer.WriteLong PartyChat
    
    SendDataToParty TempPlayer(index).inParty, Buffer.ToArray()

    
    Set Buffer = Nothing

End Sub
Public Sub HandleGuildMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString

    ' Prevent hacking
    'For i = 1 To Len(Msg)

        'If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            'Exit Sub
        'End If

    'Next
    
    If Not player(index).GuildFileId > 0 Then
        PlayerMsg index, "¡Necesitas estar en un clan!", BrightRed
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seError, 1
        Exit Sub
    End If
    
    s = "[" & GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name & "]" & GetPlayerName(index) & ": " & msg
    
    Call SayMsg_Guild(TempPlayer(index).tmpGuildSlot, index, msg, QBColor(White))
    Call AddLog(index, s, PLAYER_LOG)
    Call TextAdd(msg)
    
    Set Buffer = Nothing
End Sub
Public Sub HandleGuildSave(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

Dim Buffer As clsBuffer
Dim SaveType As Integer
Dim SentIndex As Integer
Dim HoldInt As Integer
Dim i As Integer


    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    
    SaveType = Buffer.ReadInteger
    SentIndex = Buffer.ReadInteger
    
    If SaveType = 0 Or SentIndex = 0 Then Exit Sub
    
    
    Select Case SaveType
    Case 1
        'options
        If CheckGuildPermission(index, 6) = True Then
            
            'Guild Color
            HoldInt = Buffer.ReadInteger
            If HoldInt > 0 Then
                GuildData(TempPlayer(index).tmpGuildSlot).Guild_Color = HoldInt
                HoldInt = 0
            End If
            
            'Guild Recruit rank
            HoldInt = Buffer.ReadInteger
            
            'Guild MOTD
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_MOTD = Buffer.ReadString
            
            
            'Did Recruit Rank change? Make sure they didnt set recruit rank at or above their rank
            If Not GuildData(TempPlayer(index).tmpGuildSlot).Guild_RecruitRank = HoldInt Then
                If Not HoldInt >= GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(player(index).GuildMemberId).Rank Then
                    GuildData(TempPlayer(index).tmpGuildSlot).Guild_RecruitRank = HoldInt
                    
                Else
                    PlayerMsg index, "No puedes asignar tanto nivel a un recluta.", BrightRed
                End If
            End If
        Else
            PlayerMsg index, "No estas autorizado para guardar esas opciones.", BrightRed
        End If
        HoldInt = 0
    Case 2
        'users
        If CheckGuildPermission(index, 5) = True Then
            'Guild Member Rank
            HoldInt = Buffer.ReadInteger
            If HoldInt > 0 Then
                GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(SentIndex).Rank = HoldInt
            Else
                PlayerMsg index, "El rango debe ser superior a 0", BrightRed
            End If
            
            'Guild Member Comment
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(SentIndex).Comment = Buffer.ReadString
        Else
            PlayerMsg index, "No estas autorizado para guardar usuarios.", BrightRed
        End If
        
    Case 3
        'ranks
        If CheckGuildPermission(index, 4) = True Then
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_Ranks(SentIndex).Name = Buffer.ReadString
                For i = 1 To MAX_GUILD_RANKS_PERMISSION
                    GuildData(TempPlayer(index).tmpGuildSlot).Guild_Ranks(SentIndex).RankPermission(i) = Buffer.ReadByte
                Next i
        Else
            PlayerMsg index, "No estas autorizado para guardar rangos.", BrightRed
        End If
    
    End Select
    
    Call SendGuild(True, index, TempPlayer(index).tmpGuildSlot)
    
    Set Buffer = Nothing
End Sub
Public Sub HandleGuildCommands(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Integer
    Dim SelectedIndex As Long
    Dim SendText As String
    Dim SelectedCommand As Integer
    Dim MembersCount As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    
    SelectedCommand = Buffer.ReadInteger
    SendText = Buffer.ReadString
    
    'Command 1/6/7 can be used while not in a guild
    If player(index).GuildFileId = 0 And Not (SelectedCommand = 1 Or SelectedCommand = 6 Or SelectedCommand = 7) Then
        PlayerMsg index, "¡Debes estar en un clan para usar éstos comandos!", BrightRed
        Exit Sub
    End If
    
    
    
    Select Case SelectedCommand
    Case 1
        'make
        Call MakeGuild(index, SendText)
    
    Case 2
        'invite
        'Find user index
        SelectedIndex = 0
        
        'Try to find player
        SelectedIndex = FindPlayer(SendText)
        
        If SelectedIndex > 0 Then
            Call Request_Guild_Invite(SelectedIndex, TempPlayer(index).tmpGuildSlot, index)
        Else
            PlayerMsg index, GetTranslation("No se pudo encontrar el usuario", , UnTrimBack) & SendText & ".", BrightRed, , False
        End If
        
    Case 3
        'leave
        Call GuildLeave(index)
        
    Case 4
        'admin
        If CheckGuildPermission(index, 1) = True Then
            Call ToggleGuildAdmin(index, TempPlayer(index).tmpGuildSlot, True)
        Else
            PlayerMsg index, "No le está permitido abrir el panel de administración.", BrightRed
        End If
    
    Case 5
        'view
        'This sets the default option
        If LenB(SendText) = 0 Then SendText = "online"
        MembersCount = 0
        
        Select Case SendText
        Case "online"
            PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name & " Members List (" & UCase$(SendText) & ")", Green, , False
            For i = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).Used = True Then
                    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).Online = True Then
                        PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).User_Name, Green, True, False
                        MembersCount = MembersCount + 1
                    End If
                End If
            Next i
            
            PlayerMsg index, "Total: " & MembersCount, Green, , False
        
        Case "all"
            PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name & " Members List (" & UCase$(SendText) & ")", Green
            For i = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).Used = True Then
                    PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).User_Name, Green, , False
                    MembersCount = MembersCount + 1
                End If
            Next i
            
            PlayerMsg index, "Total: " & MembersCount, Green, , False
        
        Case "offline"
            PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name & " Members List (" & UCase$(SendText) & ")", Green, , False
            For i = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).Used = True Then
                    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).Online = False Then
                        PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).User_Name, Green, , False
                        MembersCount = MembersCount + 1
                    End If
                End If
            Next i
            
            PlayerMsg index, "Total: " & MembersCount, Green, , False
        
        End Select
    Case 6
        'accept
        If TempPlayer(index).tmpGuildInviteSlot > 0 Then
            If GuildData(TempPlayer(index).tmpGuildInviteSlot).In_Use = True And GuildData(TempPlayer(index).tmpGuildInviteSlot).Guild_Fileid = TempPlayer(index).tmpGuildInviteId Then
                Call Join_Guild(index, TempPlayer(index).tmpGuildInviteSlot)
                TempPlayer(index).tmpGuildInviteSlot = 0
                TempPlayer(index).tmpGuildInviteTimer = 0
                TempPlayer(index).tmpGuildInviteId = 0
            Else
                PlayerMsg index, "Nadie de este clan está en línea, por favor pida una nueva invitación.", BrightRed
            End If
        Else
            PlayerMsg index, "Debe obtener una invitación de clan para utilizar este comando.", BrightRed
        End If
    Case 7
        'decline
        If TempPlayer(index).tmpGuildInviteSlot > 0 Then
            TempPlayer(index).tmpGuildInviteSlot = 0
            TempPlayer(index).tmpGuildInviteTimer = 0
            TempPlayer(index).tmpGuildInviteId = 0
            PlayerMsg index, "Se rechazó la invitación al clan.", BrightRed
        Else
            PlayerMsg index, "Debe obtener una invitación de clan para utilizar este comando.", BrightRed
        End If
        
    Case 8
        'founder
        'Make sure the person who used the command is who they say they are
        If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(player(index).GuildMemberId).User_Login = player(index).login Then
            'Make sure they are founder
            If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(player(index).GuildMemberId).Founder = True Then
                'Find user index
                SelectedIndex = 0
                
                'Try to find player
                SelectedIndex = FindPlayer(SendText)
                
                If SelectedIndex > 0 Then
                    'Make sure the person getting founder is the correct person
                    If GuildData(TempPlayer(SelectedIndex).tmpGuildSlot).Guild_Members(player(SelectedIndex).GuildMemberId).User_Login = player(SelectedIndex).login Then
                        GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(player(index).GuildMemberId).Founder = False
                        GuildData(TempPlayer(SelectedIndex).tmpGuildSlot).Guild_Members(player(SelectedIndex).GuildMemberId).Founder = True
                        PlayerMsg index, "Has cedido el estatus de Fundador a " & SendText & ".", BrightRed
                        PlayerMsg SelectedIndex, "¡Te ha sido otorgado el grado de Fundador del Clan!.", Green
                    End If
                Else
                    PlayerMsg index, GetTranslation("No se pudo encontrar el usuario", , UnTrimBack) & SendText & ".", BrightRed, , False
                End If
            Else
                 PlayerMsg index, "Usted debe ser el fundador para utilizar este comando.", BrightRed
            End If
        End If
    Case 9
        'kick
        Call GuildKick(TempPlayer(index).tmpGuildSlot, index, SendText)
    
    Case 10
        'disband
        Call DisbandGuild(TempPlayer(index).tmpGuildSlot, index)
        
    
    End Select
  
    Set Buffer = Nothing
End Sub

Function GuildNameBlacklist(ByRef Name As String) As Boolean
        If InStr(1, LCase(Name), "admin") > 0 Then GuildNameBlacklist = True
        If InStr(1, LCase(Name), "game master") > 0 Then GuildNameBlacklist = True
        If InStr(1, Name, "GM") > 0 Then GuildNameBlacklist = True
        If InStr(1, LCase(Name), "dragoon") > 0 Then GuildNameBlacklist = True
End Function

Function GuildNameExist(ByRef Name As String) As Boolean

    Dim FileName As String
    FileName = App.Path & DATA_PATH & GUILD_PATH & GUILD_NAMES_FILE
    GuildNameExist = LineExists(FileName, Trim$(Name))
    

End Function

Sub CheckGuildNamesFile()
    If Not FileExist(App.Path & DATA_PATH & GUILDS_PATH & GUILD_NAMES_FILE, True) Then
        CreateGuildNamesFile
        FillGuildNames
    End If
End Sub

Sub FillGuildNames()
    Dim AccountsFolder As Folder
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    Set AccountsFolder = FSO.GetFolder(App.Path & DATA_PATH & GUILD_PATH)
    Dim F As Long

    Dim Archivo As File

    Dim FileName As String
    For Each Archivo In AccountsFolder.Files
        
        Dim AuxGuild As GuildRec
        FileName = Archivo.Path
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , AuxGuild
        Close #F
        
        SaveGuildName AuxGuild.Guild_Name

    Next
End Sub

Sub CreateGuildNamesFile()
    Call CreateFile(App.Path & DATA_PATH & GUILDS_PATH & GUILD_NAMES_FILE)
End Sub

Public Sub SaveGuildName(ByRef Name As String)
    AddLine App.Path & DATA_PATH & GUILDS_PATH & GUILD_NAMES_FILE, Trim$(Name)
End Sub

Public Sub DeleteGuildName(ByRef Name As String)
    DeleteLine App.Path & DATA_PATH & GUILDS_PATH & GUILD_NAMES_FILE, ".txt", Trim$(Name)
End Sub

Public Sub SendPlayerGuildData(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SGuildData

    If player(index).GuildFileId > 0 Then
        If TempPlayer(index).tmpGuildSlot > 0 Then
            Buffer.WriteByte 1
            Buffer.WriteString GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name
            Buffer.WriteLong player(index).GuildMemberId
        End If
    Else
        Buffer.WriteByte 0
    End If
    
    SendDataTo index, Buffer.ToArray
End Sub



