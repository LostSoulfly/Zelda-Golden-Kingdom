Attribute VB_Name = "modGames"
Option Explicit
'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Public Const MAX_TEAM_MEMBERS As Byte = 4
Public Const MAX_CURRENT_GAMES As Byte = 1
Public Const MAX_GAME_TEAMS As Byte = 20

Public Enum GamesCommandType
    GInviteTeam = 1
    GCreateTeam
    GJoinGame
    GCreateGame
    GStartGame
    GRequestTeamInfo
    GBet
End Enum

Public Enum GamesInfoType
    GAddTeam
    GDeleteTeam
    GAddTeamPlayer
    GDeleteTeamPlayer
    GAddGame
    GDeleteGame
End Enum


Public Enum GameTeamType
    GuildTeam
    PartyTeam
    GameTeamCount
End Enum

Public Enum GameTypeEnum
    CaptureTheFlag
    GameCount
End Enum

Private Type GameTeamRec
    TeamType As GameTeamType
    Members(1 To MAX_TEAM_MEMBERS) As Long 'link to index's
    Representant As Long
    NumMembers As Long
    IndexNum As Long 'IndexNum = GuildNum or IndexNum = PartyNum
    GameIndex As Byte
    
    'Fill the next spaces with game info(points, position..)
End Type

Private Type CaptureTheFlagRec
    Flag() As Long
    Score() As Byte
End Type

Private Type GameRec
    Teams() As Long 'link to teams (Flag)
    NTeams As Byte
    GameType As GameTypeEnum
    CTFInfo As CaptureTheFlagRec
    IsWaiting As Boolean
    IsRunning As Boolean
End Type

Public Teams(1 To MAX_GAME_TEAMS) As GameTeamRec
Public CurrentGames(1 To MAX_CURRENT_GAMES) As GameRec

Public Teams_HighIndex As Byte

Public Function FindOpenTeamSlot() As Byte
Dim i As Byte
For i = 1 To MAX_GAME_TEAMS
    If Teams(i).Representant = 0 Then
        FindOpenTeamSlot = i
        Exit Function
    End If
Next

End Function

Public Sub CreateTeam(ByVal index As Long, ByVal TeamType As GameTeamType)

If Not CanPlayerAlist(index, TeamType) Then Exit Sub
Dim i As Byte
i = FindOpenTeamSlot
If i > 0 Then
    With Teams(i)
        .Members(1) = index
        .Representant = index
        .TeamType = TeamType
        .NumMembers = 1
        .GameIndex = 0
        .IndexNum = GetPlayerGameTeamTypeIndex(index, TeamType)
    End With
End If

End Sub

Sub CreateCompleteTeam(ByVal index As Long, ByVal TeamType As GameTeamType)
If Not CanPlayerAlist(index, TeamType) Then Exit Sub
Dim i As Byte
i = FindOpenTeamSlot
If i > 0 Then
    With Teams(i)
        Dim j As Long
        .Representant = index
        .TeamType = TeamType
        .IndexNum = GetPlayerGameTeamTypeIndex(index, TeamType)
        .GameIndex = 0
        For j = 1 To Player_HighIndex
            If CanAlistToCurrentTeam(j, i, TeamType) Then
                AddTeamMember j, i
            End If
        Next
    End With
End If
End Sub

Function CanAlistToCurrentTeam(ByVal index As Long, ByVal TeamIndex As Byte, ByVal TeamType As GameTeamType) As Boolean
    If IsInTeam(index) = 0 Then
        If CanPlayerAlist(index, TeamType) Then
            If Teams(TeamIndex).IndexNum = GetPlayerGameTeamTypeIndex(index, TeamType) Then
                CanAlistToCurrentTeam = True
            End If
        End If
    End If
End Function
Function IsInTeam(ByVal index As Long) As Byte
    IsInTeam = TempPlayer(index).TeamIndex
End Function


Sub DisbandPlayerFromGame(ByVal index As Long, ByVal TeamIndex As Byte)
Dim GameIndex As Byte
GameIndex = Teams(TeamIndex).GameIndex
    If GameIndex > 0 Then
        If IsGameRunning(GameIndex) Then
            Select Case GetGameType(GameIndex)
            Case CaptureTheFlag
            
                Dim Flag As Byte
                Flag = IsPlayerCarryingFlag(index, GameIndex)
                If Flag > 0 Then
                    ReturnFlagToBase GameIndex, Flag
                    GameMsg GameIndex, GetPlayerName(index) & " se ha desconectado y la bandera que llevaba ha sido devuelta a la base!", GetFlagColour(GetPlayerFlag(index))
                End If
                
            End Select
        End If
    End If
End Sub

Sub TeamMsg(ByVal TeamIndex As Byte, ByVal msg As String, ByVal color As Long)
    Dim i As Byte
    For i = 1 To MAX_TEAM_MEMBERS
        If Teams(TeamIndex).Members(i) > 0 Then
            PlayerMsg Teams(TeamIndex).Members(i), msg, color
        End If
    Next
End Sub

Sub GameMsg(ByVal GameIndex As Byte, ByVal msg As String, ByVal color As Long)
    Dim i As Byte
    For i = 1 To CurrentGames(GameIndex).NTeams
        Call TeamMsg(CurrentGames(GameIndex).Teams(i), msg, color)
    Next
End Sub

Sub ClearTeamPlayer(ByVal index As Long)
    Dim i As Byte
    i = IsInTeam(index)
    If i > 0 Then
        With Teams(i)
            .NumMembers = .NumMembers - 1
            TempPlayer(index).TeamIndex = 0
            Call DisbandPlayerFromGame(index, i)
            
            Dim j As Byte
            For j = 1 To MAX_TEAM_MEMBERS
                If .Members(j) = index Then
                    .Members(j) = 0
                    Exit For
                End If
            Next
                   
            If .Representant = index Then
                TransferTeamRepresentant i
            End If
        End With
    End If
End Sub

Sub ClearTeamPlayerByMemberIndex(ByVal TeamIndex As Byte, ByVal MemberIndex As Byte, Optional ByVal SetNewRepresentant As Boolean = False)
    With Teams(TeamIndex)
        TempPlayer(.Members(MemberIndex)).TeamIndex = 0
        Call DisbandPlayerFromGame(.Members(MemberIndex), TeamIndex)
        Teams(TeamIndex).Members(MemberIndex) = 0
        .NumMembers = .NumMembers - 1
        If SetNewRepresentant Then
            TransferTeamRepresentant TeamIndex
        End If
    End With
End Sub

Function TransferTeamRepresentant(ByVal TeamIndex As Byte) As Boolean
    Dim j As Byte
    With Teams(TeamIndex)
    .Representant = 0
    For j = 1 To MAX_TEAM_MEMBERS
        If .Members(j) > 0 Then
            .Representant = .Members(j)
            TransferTeamRepresentant = True
            Exit Function
        End If
    Next
    End With
End Function


Sub ClearTeamFromGame(ByVal TeamIndex As Byte, ByVal GameIndex As Byte)
    If TeamIndex < 1 Or TeamIndex > MAX_GAME_TEAMS Then Exit Sub
    Dim i As Byte
    If GameIndex > 0 Then
        With CurrentGames(GameIndex)
        For i = 1 To .NTeams
            If .Teams(i) = TeamIndex Then
                .Teams(i) = 0
                '.NTeams = .NTeams - 1
                Exit For
            End If
        Next
        End With
    End If
        
    
    
    
    Teams(TeamIndex).GameIndex = 0
    
End Sub

Function GetGameMinTeams(ByVal GameIndex As Byte) As Byte
    Select Case CurrentGames(GameIndex).GameType
    Case CaptureTheFlag
        GetGameMinTeams = CTFMinTeams
    End Select
End Function
Function CTFMinTeams() As Byte
    If FileExist("\data\Games\CTF\Rules.ini") Then
        CTFMinTeams = GetVar(App.Path & "\data\Games\CTF\Rules.ini", "RULES", "MinTeams")
    End If
End Function

Function CTFMaxTeams() As Byte
    If FileExist("\data\Games\CTF\Rules.ini") Then
        CTFMaxTeams = GetVar(App.Path & "\data\Games\CTF\Rules.ini", "RULES", "MaxTeams")
    End If
End Function

Public Sub DisbandTeam(ByVal TeamIndex As Byte)
    If TeamIndex < 1 Or TeamIndex > MAX_GAME_TEAMS Then Exit Sub
    Call ClearTeamFromCurrentGames(TeamIndex)
    Call ClearTeamPlayers(TeamIndex)
    Call ZeroMemory(Teams(TeamIndex), Len(Teams(TeamIndex)))
End Sub

Public Sub ClearTeam(ByVal TeamIndex As Byte)
    If TeamIndex < 1 Or TeamIndex > MAX_GAME_TEAMS Then Exit Sub

    'If GetTeamGame(TeamIndex) > 0 Then
        Call ClearTeamFromGame(TeamIndex, GetTeamGame(TeamIndex))
        Call ClearTeamPlayers(TeamIndex)
    'End If
    Call ZeroMemory(Teams(TeamIndex), Len(Teams(TeamIndex)))
End Sub

Public Sub ClearTeamFromCurrentGames(ByVal TeamIndex As Byte)
    If TeamIndex < 1 Or TeamIndex > MAX_GAME_TEAMS Then Exit Sub
    Dim GameIndex As Byte
    GameIndex = GetTeamGame(TeamIndex)
    If GameIndex > 0 Then
        Teams(TeamIndex).GameIndex = 0
        Select Case GetGameType(GameIndex)
        Case CaptureTheFlag
        End Select
    
    End If
End Sub

Sub ClearTeamPlayers(ByVal TeamIndex As Byte)
    If TeamIndex < 1 Or TeamIndex > MAX_GAME_TEAMS Then Exit Sub
    Dim i As Byte
    For i = 1 To MAX_TEAM_MEMBERS
        If Teams(TeamIndex).Members(i) > 0 Then
            Call ClearTeamPlayerByMemberIndex(TeamIndex, i, False)
        End If
    Next
    
    For i = 1 To Player_HighIndex
        If TempPlayer(i).TeamIndex = TeamIndex Then
            TempPlayer(i).TeamIndex = 0
        End If
    Next
End Sub

Public Function FindOpenMemberSlot(ByVal TeamIndex As Byte) As Byte
    Dim i As Byte
    For i = 1 To MAX_TEAM_MEMBERS
        If Teams(TeamIndex).Members(i) = 0 Then
            FindOpenMemberSlot = i
            Exit Function
        End If
    Next
    
End Function

Public Sub AddTeamMember(ByVal index As Long, ByVal TeamIndex As Byte)
    Dim i As Byte
    If TeamIndex < 1 Or TeamIndex > MAX_GAME_TEAMS Then Exit Sub
    With Teams(TeamIndex)
    
    If Not CanPlayerAlist(index, .TeamType) Then Exit Sub
    
    i = FindOpenMemberSlot(TeamIndex)
    If i > 0 Then
        
        .Members(i) = index
        .NumMembers = .NumMembers + 1
        TempPlayer(index).TeamIndex = TeamIndex
    End If
    
    End With
End Sub

Public Sub AddGameTeam(ByVal GameIndex As Byte, ByVal TeamIndex As Byte)

    If GameIndex < 1 Or GameIndex > MAX_CURRENT_GAMES Or TeamIndex < 1 Or TeamIndex > MAX_GAME_TEAMS Then Exit Sub
    If Not CurrentGames(GameIndex).IsWaiting Then Exit Sub
    With CurrentGames(GameIndex)
    
    Select Case .GameType
    Case CaptureTheFlag
        If .NTeams < CTFMaxTeams Then
            .NTeams = .NTeams + 1
            .Teams(.NTeams) = TeamIndex
            Teams(TeamIndex).GameIndex = GameIndex
        End If
    End Select
    
    End With
    
    
End Sub

Public Function GetPlayerGameTeamTypeIndex(ByVal index As Long, ByVal TeamType As GameTeamType) As Long
    Select Case TeamType
    Case GuildTeam
        GetPlayerGameTeamTypeIndex = player(index).GuildFileId
    Case PartyTeam
        GetPlayerGameTeamTypeIndex = TempPlayer(index).inParty
    End Select
End Function

Public Function CanPlayerAlist(ByVal index As Long, ByVal TeamType As GameTeamType) As Boolean
    Select Case TeamType
    Case GuildTeam
        If player(index).GuildFileId > 0 Then
            CanPlayerAlist = True
        End If
    Case PartyTeam
        If TempPlayer(index).inParty > 0 Then
            CanPlayerAlist = True
        End If
    End Select
End Function

Function FindOpenGameSlot() As Byte
    FindOpenGameSlot = 0
    Dim i As Byte
    For i = 1 To MAX_CURRENT_GAMES
        If Not IsGameWaiting(i) Then
            FindOpenGameSlot = i
            Exit Function
        End If
    Next
End Function

Function IsGameWaiting(ByVal GameIndex As Byte) As Boolean
If GameIndex > 0 And GameIndex < MAX_CURRENT_GAMES Then
    IsGameWaiting = CurrentGames(GameIndex).IsWaiting
End If
End Function

Sub CreateGame(ByVal GameType As GameTypeEnum)
    Select Case GameType
    Case CaptureTheFlag
        Dim MaxTeams As Byte
        MaxTeams = CTFMaxTeams
        Dim i As Byte
        i = FindOpenGameSlot
        If i > 0 Then
            With CurrentGames(i)
                .GameType = CaptureTheFlag
                .IsWaiting = True
                .IsRunning = False
                ReDim .Teams(1 To MaxTeams)
                .NTeams = 0
                ReDim .CTFInfo.Flag(1 To MaxTeams)
                ReDim .CTFInfo.Score(1 To MaxTeams)
            End With
        End If
    End Select
End Sub

Function TeamExists(ByVal TeamIndex As Byte) As Boolean
    If Not (TeamIndex > 0 And TeamIndex <= MAX_GAME_TEAMS) Then Exit Function
    
    If Teams(TeamIndex).Representant > 0 Then
        TeamExists = True
    End If
End Function

Sub StartGame(ByVal GameIndex As Byte)
    If GameIndex > 0 And GameIndex <= MAX_CURRENT_GAMES Then
        CurrentGames(GameIndex).IsRunning = True
        'Teleport Players To Game if it exists
        Select Case CurrentGames(GameIndex).GameType
        Case CaptureTheFlag
             Call CTFTeleportPlayers(GameIndex)
        End Select
    End If
End Sub

Sub ClearGame(ByVal GameIndex As Byte)
    If GameIndex < 1 Or GameIndex > MAX_CURRENT_GAMES Then Exit Sub
    With CurrentGames(GameIndex)
        .IsWaiting = False
        .IsRunning = False
        Dim i As Byte
        For i = 1 To .NTeams
            .Teams(i) = 0
            .CTFInfo.Flag(i) = 0
            .CTFInfo.Score(i) = 0
        Next
        .GameType = 0
        .NTeams = 0
    End With
End Sub

Function IsMember(ByVal index As Long, ByVal TeamIndex As Byte) As Boolean
    With Teams(TeamIndex)
    Dim i As Byte
    For i = 1 To MAX_TEAM_MEMBERS
        If .Members(i) = index Then
            IsMember = True
            Exit Function
        End If
    Next
    End With
End Function

Function GetPlayerTeam(ByVal index As Long) As Byte
    Dim i As Byte
    GetPlayerTeam = 0
    For i = 1 To MAX_GAME_TEAMS
        If IsMember(index, i) Then
            GetPlayerTeam = i
            Exit Function
        End If
    Next
End Function

Function IsGameRunning(ByVal GameIndex As Byte) As Boolean
    If GameIndex > 0 Then
        IsGameRunning = CurrentGames(GameIndex).IsRunning
    End If
End Function

Function GetTeamFlag(ByVal GameIndex As Byte, ByVal TeamIndex As Byte) As Byte
    If GameIndex < 1 Or GameIndex > MAX_CURRENT_GAMES Then Exit Function
    If TeamIndex < 1 Or TeamIndex > MAX_GAME_TEAMS Then Exit Function
    
    If CurrentGames(GameIndex).NTeams > 0 Then
        Dim i As Byte
        For i = 1 To CurrentGames(GameIndex).NTeams
            If CurrentGames(GameIndex).Teams(i) = TeamIndex Then
                GetTeamFlag = i
            End If
        Next
    End If
End Function

Sub CTFTeleportPlayers(ByVal GameIndex As Byte)
    If GameIndex < 1 Or GameIndex > MAX_CURRENT_GAMES Then Exit Sub
    Dim map As Long, X As Long, Y As Long
    If CurrentGames(GameIndex).NTeams >= CTFMinTeams Then
        Dim i As Byte
        For i = 1 To CurrentGames(GameIndex).NTeams
            GetCTFSpawningSite i, map, X, Y
            If CurrentGames(GameIndex).Teams(i) > 0 Then
                TeleportTeamPlayers CurrentGames(GameIndex).Teams(i), map, X, Y
            End If
        Next
    
    End If
End Sub

Sub CTFTeleportPlayer(ByVal index As Long, ByVal Flag As Byte)
    Dim map As Long, X As Long, Y As Long
    Dim i As Byte

    GetCTFSpawningSite Flag, map, X, Y
    PlayerWarpByEvent index, map, X, Y

End Sub


Sub TeleportTeamPlayers(ByVal TeamIndex As Byte, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    Dim i As Byte
    
    If Not (TeamIndex > 0 And TeamIndex <= MAX_GAME_TEAMS) Then Exit Sub
    
    
    For i = 1 To MAX_TEAM_MEMBERS
        If Teams(TeamIndex).Members(i) > 0 Then
            Call PlayerWarpByEvent(Teams(TeamIndex).Members(i), mapnum, X, Y)
        End If
    Next
End Sub

Function IsFlagTaken(ByVal GameIndex As Byte, ByVal Flag As Byte) As Boolean

    If Not (GameIndex > 0 And GameIndex <= MAX_CURRENT_GAMES) Then Exit Function
    
    If CurrentGames(GameIndex).CTFInfo.Flag(Flag) > 0 Then
        IsFlagTaken = True
    End If

End Function


Function CTFGetFlagByPosition(ByVal GameIndex As Byte, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long) As Byte
If GameIndex < 1 Or GameIndex > MAX_CURRENT_GAMES Then Exit Function

Dim i As Byte
For i = 1 To CurrentGames(GameIndex).NTeams
    Dim fmap As Long, fx As Long, fy As Long
    Call GetCTFBaseCoordenades(i, fmap, fx, fy)
    If fmap = mapnum And fx = X And fy = Y Then
        CTFGetFlagByPosition = i
        Exit Function
    End If
Next
End Function

Sub TakeFlag(ByVal index As Long, ByVal GameIndex As Byte, ByVal Flag As Byte)
    If GameIndex > 0 And GameIndex <= MAX_CURRENT_GAMES Then
        CurrentGames(GameIndex).CTFInfo.Flag(Flag) = index
    End If
End Sub

Sub GetCTFBaseCoordenades(ByVal Flag As Byte, ByRef mapnum As Long, ByRef X As Long, ByRef Y As Long)
    If FileExist("\data\Games\CTF\Rules.ini") Then
        mapnum = GetVar(App.Path & "\data\Games\CTF\Rules.ini", "TEAM" & Flag, "BaseMap")
        X = GetVar(App.Path & "\data\Games\CTF\Rules.ini", "TEAM" & Flag, "BaseX")
        Y = GetVar(App.Path & "\data\Games\CTF\Rules.ini", "TEAM" & Flag, "BaseY")
    End If
End Sub


Function OnRivalFlag(ByVal index As Long, ByVal GameIndex As Byte) As Byte
    If GameIndex < 1 Or GameIndex > MAX_CURRENT_GAMES Then Exit Function
    
    Dim map As Long, X As Long, Y As Long
    Dim i As Byte
    For i = 1 To CurrentGames(GameIndex).NTeams
        If i <> GetPlayerFlag(index) Then
            Call GetCTFBaseCoordenades(i, map, X, Y)
            If GetPlayerMap(index) = map And GetPlayerX(index) = X And GetPlayerY(index) = Y Then
                OnRivalFlag = i
                Exit Function
            End If
        End If
    Next

End Function

Function OnOwnFlag(ByVal index As Long, ByVal Flag As Byte) As Boolean
    Dim map As Long, X As Long, Y As Long
    Call GetCTFBaseCoordenades(Flag, map, X, Y)
    If GetPlayerMap(index) = map And GetPlayerX(index) = X And GetPlayerY(index) = Y Then
        OnOwnFlag = True
    End If
End Function

Function GetCTFMaxScore() As Long
    If FileExist("\data\Games\CTF\Rules.ini") Then
        GetCTFMaxScore = GetVar(App.Path & "\data\Games\CTF\Rules.ini", "RULES", "MaxScore")
    End If
End Function

Sub GetCTFSpawningSite(ByVal Flag As Byte, ByRef mapnum As Long, ByRef X As Long, ByRef Y As Long)
    If FileExist("\data\Games\CTF\Rules.ini") Then
        mapnum = GetVar(App.Path & "\data\Games\CTF\Rules.ini", "TEAM" & Flag, "RespawnMap")
        X = GetVar(App.Path & "\data\Games\CTF\Rules.ini", "TEAM" & Flag, "RespawnX")
        Y = GetVar(App.Path & "\data\Games\CTF\Rules.ini", "TEAM" & Flag, "RespawnY")
    End If
End Sub
Sub CheckIfGameFinished(ByVal GameIndex As Byte)
    If Not (GameIndex > 0 And GameIndex <= MAX_CURRENT_GAMES) Then Exit Sub
    Select Case CurrentGames(GameIndex).GameType
    Case CaptureTheFlag
        With CurrentGames(GameIndex)
        Dim i As Byte
        For i = 1 To CurrentGames(GameIndex).NTeams
        If .CTFInfo.Score(i) >= GetCTFMaxScore Then
            If .Teams(i) > 0 Then
                GlobalMsg "El equipo de " & GetPlayerName(Teams(.Teams(i)).Representant) & " ha ganado el juego!", BrightGreen
                AddLog 0, "Ganador: " & GetPlayerName(Teams(.Teams(i)).Representant), ADMIN_LOG
                AdminMsg "Ganador: " & GetPlayerName(Teams(.Teams(i)).Representant), BrightGreen
            End If
            Dim j As Byte
            For j = 1 To .NTeams
                If .Teams(j) > 0 Then
                    AddLog 0, "CTF: " & GetPlayerName(Teams(.Teams(j)).Representant) & ": " & .CTFInfo.Score(j), ADMIN_LOG
                    AdminMsg "CTF: " & GetPlayerName(Teams(.Teams(j)).Representant) & ": " & .CTFInfo.Score(j), BrightGreen
                    TeleportTeamPlayers .Teams(j), 53, 27, 22
                    ClearTeamFromGame .Teams(j), GameIndex
                End If
            Next
            Call ClearGame(GameIndex)
            Exit Sub
        End If
        Next
        End With
    End Select
End Sub

Sub CTFScorePoint(ByVal index As Long, ByVal GameIndex As Byte, ByVal Flag As Byte)
    If GameIndex > 0 And GameIndex <= MAX_CURRENT_GAMES Then
        CurrentGames(GameIndex).CTFInfo.Score(Flag) = CurrentGames(GameIndex).CTFInfo.Score(Flag) + 1
        Call CheckIfGameFinished(GameIndex)
    End If
End Sub

Sub ReturnFlagToBase(ByVal GameIndex As Byte, ByVal Flag As Byte)
    If GameIndex > 0 And GameIndex <= MAX_CURRENT_GAMES Then
        CurrentGames(GameIndex).CTFInfo.Flag(Flag) = 0
    End If
End Sub

Function IsPlayerCarryingFlag(ByVal index As Long, ByVal GameIndex As Byte) As Byte
    If GameIndex > 0 And GameIndex <= MAX_CURRENT_GAMES Then
        Dim i As Byte
        For i = 1 To CurrentGames(GameIndex).NTeams
            If CurrentGames(GameIndex).CTFInfo.Flag(i) = index Then
                IsPlayerCarryingFlag = i
                Exit Function
            End If
        Next
    End If
End Function

Function GetFlagPoints(ByVal GameIndex As Byte, ByVal Flag As Byte) As Byte
    If GameIndex > 0 And GameIndex <= MAX_CURRENT_GAMES Then
        GetFlagPoints = CurrentGames(GameIndex).CTFInfo.Score(Flag)
    End If
End Function

Function GetPlayerFlag(ByVal index As Long) As Byte
    Dim GameIndex As Byte
    Dim TeamIndex As Byte
    TeamIndex = IsInTeam(index)
    GameIndex = Teams(TeamIndex).GameIndex
    
    If GameIndex > 0 Then
        Dim j As Byte
        For j = 1 To CurrentGames(GameIndex).NTeams
            If CurrentGames(GameIndex).Teams(j) = TeamIndex Then
                GetPlayerFlag = j
                Exit Function
            End If
        Next
    End If
End Function

Sub DisplayFlag(ByVal GameIndex As Byte, ByVal Flag As Byte, ByVal state As Boolean)
If GameIndex < 1 Or GameIndex > MAX_CURRENT_GAMES Or Flag < 1 Then Exit Sub

Dim map As Long, X As Long, Y As Long

GetCTFBaseCoordenades Flag, map, X, Y

Y = Y - 1

If OutOfBoundries(X, Y, map) Then Exit Sub

If state Then
    SendMapKeyToMap map, X, Y, 0
Else
    SendMapKeyToMap map, X, Y, 1
End If

End Sub

Sub CTFCheckHit(ByVal index As Long)
    Dim PlayerTeam As Byte
    PlayerTeam = IsInTeam(index)
    If PlayerTeam > 0 Then
        Dim GameIndex As Byte
        GameIndex = Teams(PlayerTeam).GameIndex
        
        If IsGameRunning(GameIndex) Then
            If GetGameType(GameIndex) = CaptureTheFlag Then
                Dim i As Long
                For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If IsRival(index, i) Then
                        If CTFIsInRange(index, i) Then
                            If GetPlayerMap(index) = GetPlayerMap(i) Then
                                Dim Flag As Byte
                                Flag = IsPlayerCarryingFlag(i, GameIndex)
                                If Flag > 0 Then
                                    ReturnFlagToBase GameIndex, Flag
                                    DisplayFlag GameIndex, Flag, True
                                    GameMsg GameIndex, GetPlayerName(index) & " ha tocado a " & GetPlayerName(i) & " y ha devuelto la bandera a su base!", GetFlagColour(GetPlayerFlag(index))
                                End If
                                CTFTeleportPlayer i, GetPlayerFlag(i)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                Next
            End If
        End If
    End If
End Sub
Function CTFIsInRange(ByVal i As Long, ByVal j As Long) As Boolean
    If GetPlayerMap(i) = GetPlayerMap(j) Then
        If IsinRange(2, GetPlayerX(i), GetPlayerY(i), GetPlayerX(j), GetPlayerY(j)) Then
            CTFIsInRange = True
        End If
    End If
End Function

Function GetGameType(ByVal GameIndex As Byte) As GameTypeEnum
    If GameIndex > 0 And GameIndex <= MAX_CURRENT_GAMES Then
        GetGameType = CurrentGames(GameIndex).GameType
    End If
End Function
Function GetTeamGame(ByVal TeamIndex As Byte) As Byte
    If TeamIndex > 0 And TeamIndex <= MAX_GAME_TEAMS Then
        GetTeamGame = Teams(TeamIndex).GameIndex
    End If
End Function

Function IsRival(ByVal index1 As Long, ByVal index2 As Long)
    Dim team1 As Byte
    Dim team2 As Byte
    team1 = IsInTeam(index1)
    team2 = IsInTeam(index2)
    If team1 > 0 And team2 > 0 Then
        If team1 <> team2 Then
            If GetTeamGame(team1) = GetTeamGame(team2) Then
                IsRival = True
            End If
        End If
    End If
End Function
Function GetFlagColour(ByVal Flag As Byte) As Byte
    Select Case Flag
    Case 1
        GetFlagColour = BrightRed
    Case 2
        GetFlagColour = Cyan
    End Select
End Function
Sub ComputePlayerOnFlag(ByVal index As Long)
    Dim PlayerTeam As Byte
    PlayerTeam = IsInTeam(index)
    If PlayerTeam > 0 Then
        Dim GameIndex As Byte
        GameIndex = GetTeamGame(PlayerTeam)
        If IsGameRunning(GameIndex) Then
            Dim TeamFlag As Byte
            TeamFlag = GetTeamFlag(GameIndex, PlayerTeam)
            Dim RivalFlag As Byte
            RivalFlag = OnRivalFlag(index, GameIndex)
            If RivalFlag > 0 Then
                If Not IsFlagTaken(GameIndex, RivalFlag) Then
                    TakeFlag index, GameIndex, RivalFlag
                    GameMsg GameIndex, GetPlayerName(index) & " ha tomado la bandera!", GetFlagColour(TeamFlag)
                    DisplayFlag GameIndex, RivalFlag, False
                Else
                    PlayerMsg index, "Aquí no hay ninguna bandera", BrightRed
                End If
            ElseIf OnOwnFlag(index, TeamFlag) Then
                If Not IsFlagTaken(GameIndex, TeamFlag) Then 'our flag cannot be out of base
                    Dim CarryingFlag As Byte
                    CarryingFlag = IsPlayerCarryingFlag(index, GameIndex)
                    If CarryingFlag > 0 Then
                        ReturnFlagToBase GameIndex, CarryingFlag
                        DisplayFlag GameIndex, CarryingFlag, True
                        CTFScorePoint index, GameIndex, TeamFlag
                        GameMsg GameIndex, GetPlayerName(index) & " ha depositado la bandera!", GetFlagColour(TeamFlag)
                        GameMsg GameIndex, "El equipo de " & GetPlayerName(Teams(PlayerTeam).Representant) & " ha ganado un punto! Ahora tienen: " & GetFlagPoints(GameIndex, TeamFlag), GetFlagColour(TeamFlag)
                    Else
                        PlayerMsg index, "Esta es tu base", BrightRed
                    End If
                Else
                    PlayerMsg index, "Tu Bandera no esta aqui", BrightRed
                End If
            End If
        End If
    End If
End Sub



