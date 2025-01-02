Attribute VB_Name = "modSpell"
Option Explicit
Sub SpellStatBuffer(ByVal index As Long, ByVal spellnum As Long)
    If index < 1 Or spellnum < 1 Then Exit Sub
    If Spell(spellnum).Type <> SPELL_TYPE_BUFFER Then Exit Sub
    
    Dim BufferTime As Long
    Dim StatAdd As Integer
    Dim stat As Stats
    
    stat = Spell(spellnum).stat
    If stat < 1 Or stat >= Stat_Count Then Exit Sub
    
    BufferTime = Spell(spellnum).Duration * 1000
    StatAdd = Spell(spellnum).vital
    
    TempPlayer(index).StatsBuffer(stat).Timer = BufferTime + GetRealTickCount
    TempPlayer(index).StatsBuffer(stat).Value = StatAdd
    ComputePlayerStat index, stat
    SendStat index, stat
End Sub

Sub SpellProtect(ByVal index As Long, ByVal spellnum As Long)
    Dim i As Byte
    If Spell(spellnum).Type <> SPELL_TYPE_PROTECT Then Exit Sub
    For i = 1 To PlayerActions_Count - 1
        If Spell(spellnum).BlockActions(i) Then
            ProtectPlayerAction index, i, Spell(spellnum).StunDuration
        End If
    Next
    
    CheckPlayerActionsProtections index
End Sub

Function GetPlayerStatBuffer(ByVal index As Long, ByVal stat As Stats) As Integer
    If index < 1 Or stat < 1 Or stat >= Stat_Count Then Exit Function
    GetPlayerStatBuffer = TempPlayer(index).StatsBuffer(stat).Value
End Function
Sub ResetPlayerStatBuffer(ByVal index As Long, ByVal stat As Stats)
    If index < 1 Or stat < 1 Or stat >= Stat_Count Then Exit Sub
    
    TempPlayer(index).StatsBuffer(stat).Timer = 0
    TempPlayer(index).StatsBuffer(stat).Value = 0
End Sub

Sub CheckPlayerStatsBuffer(ByVal index As Long, ByVal Tick As Long)
    Dim i As Byte
    Dim Reset As Boolean
    i = 1
    While i < Stats.Stat_Count
        If TempPlayer(index).StatsBuffer(i).Timer > 0 And TempPlayer(index).StatsBuffer(i).Timer < Tick Then
            ResetPlayerStatBuffer index, i
            ComputePlayerStat index, i
            SendStat index, i
        End If
        i = i + 1
    Wend
    

End Sub




Sub CheckSpellStunts(ByVal Target As Long, ByVal spellnum As Long)
    Dim i As Byte
    For i = 1 To PlayerActions_Count - 1
        If Spell(spellnum).BlockActions(i) Then
            If Not IsActionProtected(Target, i) Then
                BlockPlayerAction Target, i, Spell(spellnum).StunDuration
            End If
        End If
    Next
    
End Sub

Function SpellExists(ByVal spellnum As Long) As Boolean
If LenB(Trim$(Spell(spellnum).Name)) > 0 And Asc(Spell(spellnum).Name) <> 0 Then
    SpellExists = True
End If
End Function

Function isSpellMPPercent(ByVal spellnum As Long) As Boolean
    If Spell(spellnum).UsePercent Then
        isSpellMPPercent = True
    End If
End Function

Function GetSpellMPCost(ByVal mapnum As Long, ByVal caster As Long, ByVal castertype As Byte, ByVal spellnum As Long) As Long
    If mapnum = 0 Or caster = 0 Or spellnum = 0 Then Exit Function
    If isSpellMPPercent(spellnum) Then 'spell uses %
        Dim castervital As Long
        Select Case castertype
        Case TARGET_TYPE_PLAYER
            castervital = GetPlayerMaxVital(caster, MP)
            GetSpellMPCost = castervital * (CDbl(Spell(spellnum).MPCost) / 100)
        Case TARGET_TYPE_NPC
            castervital = GetNpcMaxVital(mapnum, caster, MP)
            GetSpellMPCost = castervital * (CDbl(Spell(spellnum).MPCost) / 100)
        End Select
    Else
        GetSpellMPCost = Spell(spellnum).MPCost
    End If
    
End Function


Public Function GetSpellDamageStat(ByVal spellnum As Long) As Byte
GetSpellDamageStat = Spell(spellnum).StatDamage + 1
End Function

Public Function GetSpellDefenseStat(ByVal spellnum As Long) As Byte
GetSpellDefenseStat = Spell(spellnum).StatDefense + 1
End Function

Public Function GetSpellChangeState(ByVal spellnum As Long) As PlayerStateType
    GetSpellChangeState = Spell(spellnum).ChangeState
End Function

Public Sub SpellChangeState(ByVal index As Long, ByVal spellnum As Long)
    If Spell(spellnum).Type <> SPELL_TYPE_CHANGESTATE Then Exit Sub
    
    Dim state As Byte
    state = GetSpellChangeState(spellnum)
    Call CheckPlayerStateChange(index, state)
    
End Sub


Function GetLastUsedSpell(ByVal index As Long) As Long
    GetLastUsedSpell = TempPlayer(index).LastSpell
End Function
