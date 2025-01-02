Attribute VB_Name = "modGames"
Option Explicit

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
