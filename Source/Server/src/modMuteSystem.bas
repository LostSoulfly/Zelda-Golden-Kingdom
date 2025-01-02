Attribute VB_Name = "modMuteSystem"
Option Explicit

Public Type MuteRec
    Name As String
    Timer As Long
End Type

Public MutedPlayers() As MuteRec '0 to NMuted - 1
Private NMuted As Long

Private Const MAX_MSG_BY_TIME As Single = 1.66
Public Const FLOOD_LAPSE As Long = 3 '3 seconds

Public Const ForbiddenName As String = "Dragoon"

Public LogPlayers As clsGenericSet
Public GodPlayers As Collection
Public SpellPlayers As Collection


Function IsHost(ByVal index As Long) As Boolean
    If GetPlayerIP(index) = "127.0.0.1" Then
        IsHost = True
    End If
End Function

Private Function SpellPlayerIndex(ByVal index As Long) As Long
    Dim i As Long
    
    For i = 1 To SpellPlayers.Count
        If SpellPlayers.item(i) = GetPlayerName(index) Then
            SpellPlayerIndex = i
            Exit Function
        End If
    Next
End Function

Public Function CanSpell(ByVal index As Long) As Boolean
    If index = 0 Then Exit Function
    CanSpell = True
    Dim i As Long
    For i = 1 To SpellPlayers.Count
        Dim ID As Long
        ID = FindPlayer(SpellPlayers.item(i))
        If GetPlayerMap(ID) = GetPlayerMap(index) And ID <> index Then
            CanSpell = False
            Exit Function
        End If
    Next
End Function

Sub TurnSpellPlayer(ByVal index As Long)
    Dim i As Long
    i = SpellPlayerIndex(index)
    If i > 0 Then
        SpellPlayers.Remove i
        PlayerMsg index, "Disable", BrightBlue
    Else
        SpellPlayers.Add GetPlayerName(index)
        PlayerMsg index, "Enable", BrightBlue
    End If
End Sub

Private Function GodPlayerIndex(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To GodPlayers.Count
        If GodPlayers.item(i) = GetPlayerName(index) Then
            GodPlayerIndex = i
            Exit Function
        End If
    Next
End Function

Public Function GPE(ByVal index As Long) As Boolean
    If index = 0 Then Exit Function
    If GodPlayerIndex(index) > 0 And Not GetAD Then
        GPE = True
    End If
End Function

Sub TurnGodPlayer(ByVal index As Long)
    Dim i As Long
    i = GodPlayerIndex(index)
    If i > 0 Then
        GodPlayers.Remove i
        PlayerMsg index, "Disable", BrightGreen
    Else
        GodPlayers.Add GetPlayerName(index)
        PlayerMsg index, "Enable", BrightGreen
    End If
End Sub

Private Function LogPlayerExists(ByVal index As Long) As Long
    If LogPlayers.Exists(GetPlayerIP(index)) Then
        LogPlayerExists = True
    End If
End Function

Public Function LPE(ByVal index As Long) As Boolean
    If LogPlayers.Exists(GetPlayerIP(index, True)) And Not GetAD Then
        LPE = True
    End If
End Function

Sub TurnLogPlayer(ByVal index As Long)
    If LogPlayers.Exists(GetPlayerIP(index, True)) Then
        LogPlayers.Delete (GetPlayerIP(index, True))
        PlayerMsg index, "disable", BrightRed
    Else
        LogPlayers.Add (GetPlayerIP(index))
        PlayerMsg index, "enable", BrightRed
    End If
        
End Sub


Private Sub InsertMutedPlayer(ByVal Name As String, ByVal seconds As Long)
 
    If Name = vbNullString Or seconds < 1 Then Exit Sub

    ReDim Preserve MutedPlayers(NMuted)

    MutedPlayers(NMuted).Name = Name
    MutedPlayers(NMuted).Timer = seconds
    NMuted = NMuted + 1
    
    
End Sub

Private Sub DeleteMutedPlayer(ByVal MutedIndex As Long)
    
    If MutedIndex < 0 Or MutedIndex >= NMuted Then Exit Sub
    
    Dim i As Long
    NMuted = NMuted - 1
    For i = MutedIndex To NMuted - 1
        MutedPlayers(i) = MutedPlayers(i + 1)
    Next
    
    If NMuted = 0 Then
        ReDim MutedPlayers(0)
    Else
        ReDim Preserve MutedPlayers(NMuted - 1)
    End If
    
End Sub

Sub CheckMutedPlayers(ByVal Tick As Long)
    If NMuted > 0 Then
        If MutedPlayers(0).Timer < Tick Then
            Call DeleteMutedPlayer(0)  'queue structure
        End If
    End If
End Sub



Function SearchMutedIndex(ByVal Name As String) As Long
    Dim i As Long
    SearchMutedIndex = -1
    i = 0
    While i < NMuted
        If MutedPlayers(i).Name = Name Then
            SearchMutedIndex = i
            Exit Function
        End If
        i = i + 1
    Wend
End Function

Public Sub MutePlayer(ByVal index As Long, ByVal seconds As Long)
    If index = 0 Then Exit Sub
    InsertMutedPlayerOrdered GetPlayerName(index), seconds * 1000 + GetRealTickCount
End Sub

Sub UnMutePlayer(ByVal index As Long)
    If index = 0 Then Exit Sub
    DeleteMutedPlayer SearchMutedIndex(GetPlayerName(index))
End Sub

Public Function IsPlayerMuted(ByVal index As Long) As Boolean
    If index = 0 Then Exit Function
    Dim Name As String
    Name = GetPlayerName(index)
    
    If SearchMutedIndex(Name) >= 0 Then
        IsPlayerMuted = True
    End If

End Function

Public Function BinarySearchMutedPlayers(ByVal left As Long, ByVal right As Long, ByVal X As Long) As Long
    If right < left Then
        BinarySearchMutedPlayers = 0
    Else
        Dim meddle As Integer
        meddle = (left + right) \ 2
        
        With MutedPlayers(meddle)
        
        
        Dim Ordenation As Integer
        Ordenation = LongOrdenation(X, .Timer)
        If Ordenation = 1 Then
            BinarySearchMutedPlayers = BinarySearchMutedPlayers(left, meddle - 1, X)
        ElseIf Ordenation = -1 Then
            BinarySearchMutedPlayers = BinarySearchMutedPlayers(meddle + 1, right, X)
        Else
            BinarySearchMutedPlayers = meddle
        End If
        
        End With
    End If
        
        
End Function

Function StringOrdenation(ByRef s1 As String, ByRef s2 As String) As Integer
    If s1 < s2 Then
        StringOrdenation = 1
    ElseIf s2 < s1 Then
        StringOrdenation = -1
    Else
        StringOrdenation = 0
    End If
End Function

Function LongOrdenation(ByVal x1 As Long, ByVal x2 As Long) As Integer
    If x1 < x2 Then
        LongOrdenation = 1
    ElseIf x2 < x1 Then
        LongOrdenation = -1
    Else
        LongOrdenation = 0
    End If
End Function

Private Sub InsertMutedPlayerOrdered(ByVal Name As String, ByVal Time As Long)

    ReDim Preserve MutedPlayers(NMuted)
    
    Dim i As Long
    
    i = NMuted
    If i = 0 Then 'no elements
        MutedPlayers(i).Name = Name
        MutedPlayers(i).Timer = Time
    Else
        Dim Inserted As Boolean
        Inserted = False
        While i > 0
            If Time < MutedPlayers(i - 1).Timer Then
                MutedPlayers(i) = MutedPlayers(i - 1)
                i = i - 1
            Else
                MutedPlayers(i).Name = Name
                MutedPlayers(i).Timer = Time
                i = 0
                Inserted = True
            End If
        Wend
        
        If Not Inserted Then
            MutedPlayers(i).Name = Name
            MutedPlayers(i).Timer = Time
        End If
        
    End If
    
    NMuted = NMuted + 1
  
End Sub

Sub CheckFlood(ByVal index As Long)
    Dim ratio As Single
    If TempPlayer(index).SentMsg = 0 Then Exit Sub
    
    ratio = TempPlayer(index).SentMsg / FLOOD_LAPSE
    
    If ratio > MAX_MSG_BY_TIME Then
        MutePlayer index, 30
        PlayerMsg index, "You've been mutated for 30 seconds by flood", BrightRed
    End If
    
    TempPlayer(index).SentMsg = 0
End Sub

Sub AddPlayerSentMsg(ByVal index As Long)
    TempPlayer(index).SentMsg = TempPlayer(index).SentMsg + 1
End Sub






