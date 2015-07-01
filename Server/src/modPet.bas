Attribute VB_Name = "modPet"
Option Explicit


Public Type PlayerPetRec
    'Link to super class
    NumPet As Byte
    StatsAdd(1 To Stats.Stat_Count - 1) As Byte
    points As Integer
    Experience As Long
    level As Long
    
    'Server side
    CurVital(1 To Vitals.Vital_Count - 1) As Long
End Type

Public Type MapPetRec
    Owner As Long
End Type

Public Type PetRec
    Name As String * NAME_LENGTH
    npcnum As Long
    'Requeriments
    TamePoints As Integer
    'Progressions
    ExpProgression As Byte
    pointsprogression As Byte
    MaxLevel As Long
End Type

'Public Type PetCache
    'Pet(1 To MAX_MAP_NPCS) As Long
    'UpperBound As Long
'End Type

Public Type TempPlayerPetRec
    Mode As Byte
    TempPetSlot As Byte
    ActualPet As Byte
    PetSpawnWait As Long
    PetHasOwnTarget As Byte
    PetExpPercent As Byte
    PetState As PetStateEnum
End Type

Public Enum PetStateEnum
    Passive = 0
    Assist = 1
    Defensive = 2
End Enum


'makes the pet follow its owner
Public Sub PetFollowOwner(ByVal index As Long)
If index <= 0 Then Exit Sub
    If TempPlayer(index).TempPet.TempPetSlot < 1 Then Exit Sub

    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPet.TempPetSlot).TargetType = 1
    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPet.TempPetSlot).Target = index
    TempPlayer(index).TempPet.PetHasOwnTarget = 0
End Sub

'makes the pet wander around the map
Public Sub PetWander(ByVal index As Long)
If index <= 0 Then Exit Sub
    If TempPlayer(index).TempPet.TempPetSlot < 1 Then Exit Sub

    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPet.TempPetSlot).TargetType = TARGET_TYPE_NONE
    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPet.TempPetSlot).Target = 0
    TempPlayer(index).TempPet.PetHasOwnTarget = 0
    
End Sub

'Clear the npc from the map
Public Function PetDisband(ByVal index As Long, ByVal mapnum As Long, Optional ByVal WaitForSpawn As Boolean = True) As Boolean
    Dim i As Long
    Dim j As Long
    Dim mapnpcnum As Long
    PetDisband = True
    If index <= 0 Then Exit Function
    If TempPlayer(index).TempPet.TempPetSlot < 1 Then
        PetDisband = False
        Exit Function
    End If
    
    
    mapnpcnum = TempPlayer(index).TempPet.TempPetSlot
    
    For i = 1 To Vital_Count - 1
        If Not WaitForSpawn Then
            player(index).Pet(TempPlayer(index).TempPet.ActualPet).CurVital(i) = MapNpc(mapnum).NPC(mapnpcnum).vital(i)
        Else
            player(index).Pet(TempPlayer(index).TempPet.ActualPet).CurVital(i) = 0
        End If
    Next
    
    Call SendClearMapNpcToMap(mapnum, mapnpcnum)
    Call ClearSingleMapNpc(mapnpcnum, mapnum)
    
    
    'Reset slot
    TempPlayer(index).TempPet.TempPetSlot = 0
    
    'Check if pet must wait until next spawn
    If WaitForSpawn = True Then
        TempPlayer(index).TempPet.PetSpawnWait = GetRealTickCount
    End If
    
    'Reset Target
    TempPlayer(index).TempPet.PetHasOwnTarget = 0


End Function

Public Sub ChangePetMap(ByVal index As Long, ByVal OldMap As Long, ByVal mapnum As Long)
If index <= 0 Then Exit Sub
'remove pet from map oldmap
PetDisband index, OldMap, False

'spawn pet on new map
    Dim PlayerMap As Long
    Dim i As Integer
    Dim PetSlot As Byte
    Dim PlayerPet As Byte
   
    'Prevent multiple pets for the same owner
    If TempPlayer(index).TempPet.TempPetSlot > 1 Then Exit Sub
    
    'slot, 1 to MAX_PLAYER_PETS
    PlayerPet = TempPlayer(index).TempPet.ActualPet
    
    'Prevent player out of range slots
    If PlayerPet <= 0 Or PlayerPet > MAX_PLAYER_PETS Then Exit Sub
    
    'Prevent spawning inexistent pet
    If player(index).Pet(TempPlayer(index).TempPet.ActualPet).NumPet < 1 Or player(index).Pet(TempPlayer(index).TempPet.ActualPet).NumPet > MAX_PETS Then Exit Sub
    
    PlayerMap = mapnum
    PetSlot = 0
    
    'Prevent Boundries
    'Select Case player(index).dir
    'Case 0
    '    If player(index).Y = map(PlayerMap).MaxY Then Exit Sub
    'Case 1
    '    If player(index).Y = 0 Then Exit Sub
    'Case 2
    '    If player(index).X = map(PlayerMap).MaxX Then Exit Sub
    'Case 3
    '    If player(index).X = 0 Then Exit Sub
    'End Select
    
    If map(PlayerMap).moral = MAP_MORAL_SAFE Or map(PlayerMap).moral = MAP_MORAL_PACIFIC Then PetDisband index, OldMap, True: PlayerMsg index, "Your pet is not allowed on this map.", White, , False: Exit Sub
    
    For i = 1 To MAX_MAP_NPCS
        If map(PlayerMap).NPC(i) = 0 And MapNpc(PlayerMap).NPC(i).Num = 0 Then
            PetSlot = i
            Exit For
        End If
    Next
    
    If PetSlot = 0 Then
        Call PlayerMsg(index, "The map is too crowded for you to call on your pet!", Red)
        PetDisband index, OldMap, True
        Exit Sub
    End If

    'create the pet for the map
    MapNpc(PlayerMap).NPC(PetSlot).Num = Pet(player(index).Pet(PlayerPet).NumPet).npcnum  'pet npc number
    
    'set its Pet Data
    MapNpc(PlayerMap).NPC(PetSlot).PetData.Owner = index
    
    TempPlayer(index).TempPet.TempPetSlot = PetSlot

    Select Case GetPlayerDir(index)
        Case DIR_UP
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index), GetPlayerY(index) + 1, Pet(player(index).Pet(PlayerPet).NumPet).npcnum, True)
        Case DIR_DOWN
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index), GetPlayerY(index) - 1, Pet(player(index).Pet(PlayerPet).NumPet).npcnum, True)
        Case DIR_LEFT
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index) + 1, GetPlayerY(index), Pet(player(index).Pet(PlayerPet).NumPet).npcnum, True)
        Case DIR_RIGHT
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index) - 1, GetPlayerY(index), Pet(player(index).Pet(PlayerPet).NumPet).npcnum, True)
    End Select

'make the pet follow the player!
    PetFollowOwner index

End Sub

Public Sub SpawnPet(ByVal index As Long, ByVal mapnum As Long)
    Dim PlayerMap As Long
    Dim i As Integer
    Dim PetSlot As Byte
    Dim PlayerPet As Byte
    Dim UntilTime As Long
    If index <= 0 Then Exit Sub
    'Prevent multiple pets for the same owner
    If TempPlayer(index).TempPet.TempPetSlot > 0 Then
        If player(index).Pet(TempPlayer(index).TempPet.ActualPet).CurVital(Vitals.HP) = 0 Then
            Debug.Print
        End If
    Exit Sub
    End If
    
    'slot, 1 to MAX_PLAYER_PETS
    PlayerPet = TempPlayer(index).TempPet.ActualPet
    
    'Prevent player out of range slots
    If PlayerPet <= 0 Or PlayerPet > MAX_PLAYER_PETS Then Exit Sub
    
    
    UntilTime = GetRealTickCount - TempPlayer(index).TempPet.PetSpawnWait - GetPetSpawnTime(player(index).Pet(PlayerPet).NumPet)
    
    'Check if SpawnWait Finished
    If UntilTime <= 0 Then
        Call PlayerMsg(index, GetTranslation("Aún no puedes invocar a tu mascota!, faltan") & " " & Round(Abs(UntilTime) / 1000, 0) & " " & GetTranslation("segundos!"), BrightRed, , False)
        Exit Sub
    End If
    
    'Prevent spawning inexistent pet
    If player(index).Pet(TempPlayer(index).TempPet.ActualPet).NumPet < 1 Or player(index).Pet(TempPlayer(index).TempPet.ActualPet).NumPet > MAX_PETS Then Exit Sub
    
    PlayerMap = GetPlayerMap(index)
    PetSlot = 0
    
    'Prevent Boundries
    Select Case player(index).dir
    Case 0
        If player(index).Y = map(PlayerMap).MaxY Then Exit Sub
    Case 1
        If player(index).Y = 0 Then Exit Sub
    Case 2
        If player(index).X = map(PlayerMap).MaxX Then Exit Sub
    Case 3
        If player(index).X = 0 Then Exit Sub
    End Select
    
        If map(PlayerMap).moral = MAP_MORAL_SAFE Or map(PlayerMap).moral = MAP_MORAL_PACIFIC Then PetDisband index, PlayerMap, True: PlayerMsg index, "Your pet is not allowed on this map.", White, , False: Exit Sub
    
    For i = 1 To MAX_MAP_NPCS
        If map(PlayerMap).NPC(i) = 0 And MapNpc(PlayerMap).NPC(i).Num = 0 Then
            PetSlot = i
            Exit For
        End If
    Next
    
    If PetSlot = 0 Then
        Call PlayerMsg(index, "The map is too crowded for you to call on your pet!", Red)
        Exit Sub
    End If

    'create the pet for the map
    MapNpc(PlayerMap).NPC(PetSlot).Num = Pet(player(index).Pet(PlayerPet).NumPet).npcnum  'pet npc number
    
    'set its Pet Data
    MapNpc(PlayerMap).NPC(PetSlot).PetData.Owner = index
    
    
    TempPlayer(index).TempPet.TempPetSlot = PetSlot
       

    Select Case GetPlayerDir(index)
        Case DIR_UP
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index), GetPlayerY(index) + 1, Pet(player(index).Pet(PlayerPet).NumPet).npcnum, True)
        Case DIR_DOWN
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index), GetPlayerY(index) - 1, Pet(player(index).Pet(PlayerPet).NumPet).npcnum, True)
        Case DIR_LEFT
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index) + 1, GetPlayerY(index), Pet(player(index).Pet(PlayerPet).NumPet).npcnum, True)
        Case DIR_RIGHT
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index) - 1, GetPlayerY(index), Pet(player(index).Pet(PlayerPet).NumPet).npcnum, True)
    End Select
    
    PetFollowOwner index
    
End Sub

Public Function CheckFreePetSlots(ByVal index As Long) As Integer
'-1: There aren't free slots, >= 1: Index of first free slot
Dim i As Byte

CheckFreePetSlots = -1
For i = 1 To MAX_PLAYER_PETS

    If player(index).Pet(i).NumPet = 0 Then
        CheckFreePetSlots = i
        Exit Function
    End If
    
Next

End Function


Public Sub CheckPlayerTame(ByVal index As Long)

Dim slot As Integer
Dim PetIndex As Integer

slot = CheckFreePetSlots(index)
If index <= 0 Then Exit Sub
If Not (slot > 0) Then
    Call PlayerMsg(index, "No tienes slots libres", BrightRed)
    Exit Sub
End If

If TempPlayer(index).TargetType <> TARGET_TYPE_NPC Then
    Call PlayerMsg(index, "No tienes ningún NPC fijado", BrightRed)
    Exit Sub
End If

If IsMapNPCaPet(GetPlayerMap(index), TempPlayer(index).Target) Then
    Call PlayerMsg(index, "El objetivo es una mascota", BrightRed)
    Exit Sub
End If

PetIndex = IsNPCaPet(MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).Target).Num)

If PetIndex <= 0 Then
    Call PlayerMsg(index, "El NPC no es domable", BrightRed)
    Exit Sub
End If

If Not (GetPlayerTamePoints(index, Pet(PetIndex).npcnum) >= Pet(PetIndex).TamePoints) Then
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Call PlayerMsg(index, "You don't have enough tame points for this pet!", BrightRed, , False)
        Exit Sub
    End If
End If

'Agregar pet a slot libre
Call AddPlayerPet(index, CByte(slot), CByte(PetIndex))
Call PlayerMsg(index, Trim$(NPC(Pet(PetIndex).npcnum).TranslatedName) & " " & GetTranslation("se ha unido a tu equipo!"), BrightGreen, , False)
Call KillNpc(GetPlayerMap(index), TempPlayer(index).Target)
Call SendPetData(index, TempPlayer(index).TempPet.ActualPet)


End Sub

Public Function IsNPCaPet(ByVal npcnum As Long) As Integer
'-1: not a pet, >0 : pet index
Dim i As Byte

IsNPCaPet = -1

For i = 1 To MAX_PETS
    If Pet(i).npcnum = npcnum Then
        IsNPCaPet = i
        Exit Function
    End If
Next

End Function

Public Function GetPlayerPetStat(ByVal index As Long, ByVal stat As Byte) As Byte
    GetPlayerPetStat = player(index).Pet(TempPlayer(index).TempPet.ActualPet).StatsAdd(stat)
End Function

Public Function GetPlayerPetTotalStat(ByVal index As Long, ByVal stat As Byte) As Byte
Dim result As Integer
If player(index).Pet(TempPlayer(index).TempPet.ActualPet).NumPet > 0 Then

    result = player(index).Pet(TempPlayer(index).TempPet.ActualPet).StatsAdd(stat) + NPC(Pet(player(index).Pet(TempPlayer(index).TempPet.ActualPet).NumPet).npcnum).stat(stat)
    If result > 255 Then
        result = 255
    End If
    
    GetPlayerPetTotalStat = CByte(result)
End If
End Function

Public Sub SetPlayerPetStat(ByVal index As Long, ByVal stat As Byte, ByVal number As Byte)

If number > MAX_STAT Or number < 0 Then Exit Sub

player(index).Pet(TempPlayer(index).TempPet.ActualPet).StatsAdd(stat) = number

End Sub

Public Function GetPlayerPetLevel(ByVal index As Long) As Long
    GetPlayerPetLevel = player(index).Pet(TempPlayer(index).TempPet.ActualPet).level
End Function

Public Function SetPlayerPetLevel(ByVal index As Long, ByVal level As Long) As Boolean
    SetPlayerPetLevel = False
    If level > MAX_LEVELS Or level > Pet(player(index).Pet(TempPlayer(index).TempPet.ActualPet).NumPet).MaxLevel Then Exit Function
    player(index).Pet(TempPlayer(index).TempPet.ActualPet).level = level
    SetPlayerPetLevel = True
End Function

Public Function GetPlayerPetExp(ByVal index As Long) As Long
    GetPlayerPetExp = player(index).Pet(TempPlayer(index).TempPet.ActualPet).Experience
End Function

Public Sub SetPlayerPetExp(ByVal index As Long, ByVal exp As Long)
    player(index).Pet(TempPlayer(index).TempPet.ActualPet).Experience = exp
End Sub

Public Function GetPlayerPetNextLevel(ByVal index As Long) As Long
    Dim exp As Long
    exp = LevelExp(GetPlayerPetLevel(index))
    GetPlayerPetNextLevel = exp / 2 + exp / 15 * Pet(player(index).Pet(GetPlayerPetSlot(index)).NumPet).ExpProgression
End Function

Public Function GetPlayerPetExpByLevel(ByVal PetNum As Byte, ByVal level As Long) As Long
If PetNum = 0 Then Exit Function
Dim exp As Long
exp = LevelExp(level)
GetPlayerPetExpByLevel = exp / 2 + exp / 15 * Pet(PetNum).ExpProgression
End Function

Function GetPetExpPercent(ByVal index As Long) As Single
    Dim nextlvl As Long
    nextlvl = GetPlayerPetNextLevel(index)
    Dim exp As Long
    exp = GetPlayerPetExp(index)
    
    If exp = 0 Then Exit Function
    If nextlvl = 0 Then Exit Function
    
    GetPetExpPercent = 100 * CSng(exp) / CSng(nextlvl)
    
End Function

Public Function GetPlayerPetMaxLevel(ByVal index As Long) As Long
    GetPlayerPetMaxLevel = Pet(player(index).Pet(TempPlayer(index).TempPet.ActualPet).NumPet).MaxLevel
End Function
Public Function GetPlayerPetPOINTS(ByVal index As Long) As Integer
    GetPlayerPetPOINTS = player(index).Pet(TempPlayer(index).TempPet.ActualPet).points
End Function
Public Sub SetPlayerPetPOINTS(ByVal index As Long, ByVal points As Integer)
    player(index).Pet(TempPlayer(index).TempPet.ActualPet).points = points
End Sub
Public Function GetPetNextPOINTS(ByVal index As Long, ByVal PetNum As Byte)
    GetPetNextPOINTS = PointsPerLevel(PetNum, player(index).Pet(TempPlayer(index).TempPet.ActualPet).level)
End Function

Sub CheckPlayerPetLevelUp(ByVal index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    Dim PointsWon As Byte
    
    level_count = 0
    
    Do While GetPlayerPetExp(index) >= GetPlayerPetNextLevel(index)
        expRollover = GetPlayerPetExp(index) - GetPlayerPetNextLevel(index)
        
        ' can level up?
        If Not SetPlayerPetLevel(index, GetPlayerPetLevel(index) + 1) Then
            Exit Do
        End If
        
        PointsWon = GetPetNextPOINTS(index, player(index).Pet(TempPlayer(index).TempPet.ActualPet).NumPet)
        Call SetPlayerPetPOINTS(index, GetPlayerPetPOINTS(index) + PointsWon)
        'PlayerMsg index, PointsWon & " puntos ganados!", cyan
        Call SetPlayerPetExp(index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            PlayerMsg index, "¡Tu mascota ha subido " & level_count & " nivel!", Brown
        Else
            'plural
            PlayerMsg index, "¡Tu mascota ha subido " & level_count & " niveles!", Brown
        End If
        'SendEXP index
        SendPetData index, TempPlayer(index).TempPet.ActualPet
        'SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seLevelUp, 1
    End If
End Sub

Public Sub GivePetEXP(ByVal index As Long, ByVal exp As Long)
    ' give the exp
    If TempPlayer(index).TempPet.ActualPet < 1 Or TempPlayer(index).TempPet.ActualPet > MAX_PLAYER_PETS Then Exit Sub
    
    Call SetPlayerPetExp(index, exp + GetPlayerPetExp(index))
    'SendEXP index
    SendPetData index, TempPlayer(index).TempPet.ActualPet
    SendActionMsg GetPlayerMap(index), "+" & exp & " EXP", White, 1, (MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPet.TempPetSlot).X * 32), (MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPet.TempPetSlot).Y * 32)
    ' check if we've leveled
    CheckPlayerPetLevelUp index
End Sub

Public Function GetPetSpawnTime(ByVal PetNum As Byte) As Long
Dim BaseTime As Long

If PetNum < 1 Or PetNum > MAX_PETS Then Exit Function

'BaseTime = 10000 '10 seconds

'Process base time
BaseTime = 1000 * NPC(Pet(PetNum).npcnum).SpawnSecs
'* (Pet(PetNum).pointsprogression + 1) * 2

'erase this
'BaseTime = 3000

GetPetSpawnTime = BaseTime
End Function

Function PointsPerLevel(ByVal PetNum As Byte, ByVal level As Long) As Byte
Dim result As Double
Dim pntsPro As Byte
Dim logarithm As Double
Dim ProFactor As Double
    
pntsPro = Pet(PetNum).pointsprogression
    If pntsPro > 0 And pntsPro <= MAX_PET_POINTS_PERLVL Then
        ProFactor = GPN(pntsPro, Pet(PetNum).MaxLevel)
        logarithm = LogB((MAX_PET_POINTS_PERLVL - 1) / ProFactor, CDbl(Pet(PetNum).MaxLevel))
        result = ProFactor * (CDbl(level) ^ logarithm) + 1
        PointsPerLevel = CByte(Round(result))
    Else
        PointsPerLevel = 3
    End If
End Function

Public Function LogB(ByVal number As Double, ByVal base As Double) As Double
    LogB = Log(number) / Log(base)
End Function

Function GPN(ByVal pntsPro As Byte, ByVal MaxLvl As Byte) As Double
Dim CentralNumber As Double
Dim Dispersion As Double
Dim multiples As Double

If pntsPro > MaxLvl Or pntsPro < 0 Then
    GPN = 0
End If

CentralNumber = CDbl(MAX_PET_POINTS_PERLVL - 1) / CLng(MaxLvl)
'If points progression is above is medium, we have to return higher number than central, cause this,
'we'll do something inverse when the progression is under his medium

Dispersion = CDbl(pntsPro) - CDbl(MaxLvl) / 2

If Dispersion > 0 Then
    multiples = (MaxLvl - CentralNumber) / (MaxLvl / 2)
ElseIf Dispersion = 0 Then
    GPN = CentralNumber
    Exit Function
ElseIf Dispersion < 0 Then
    multiples = CentralNumber / (MaxLvl / 2)
End If

GPN = CentralNumber + multiples * Dispersion

'Negative dispersion -> return lower than central number, Positive dispersion -> return higher than central number

End Function

Public Function GetPlayerTamePoints(ByVal index As Long, ByVal npcnum As Long) As Integer

GetPlayerTamePoints = CInt(GetPlayerStat(index, willpower) + GetPlayerLevel(index) + (player(index).NPCKills / 500))

End Function

Public Sub AddPlayerPet(ByVal index As Long, ByVal PetSlot As Byte, ByVal PetNum As Byte)
'Dim i As Byte
If PetSlot <= 0 Or PetSlot > MAX_PLAYER_PETS Then
    Exit Sub
End If
If PetNum <= 0 Or PetNum > MAX_PETS Then
    Exit Sub
End If

Call ResetPlayerPetSlot(index, PetSlot)

'Player(index).Pet(PetSlot).Name = GetPlayerName(index) & "'s " & NPC(Pet(PetNum).NPCNum).Name
player(index).Pet(PetSlot).NumPet = PetNum
player(index).Pet(PetSlot).level = 1

End Sub

Public Sub InitPlayerPets(ByVal index As Long)
    TempPlayer(index).TempPet.ActualPet = 1
    TempPlayer(index).TempPet.TempPetSlot = 0
    TempPlayer(index).TempPet.PetExpPercent = 50
    
    'Player(index).Pet(TempPlayer(index).TempPet.ActualPet).SpawnWait = GetRealTickCount
End Sub

Public Sub ResetPlayerPetSlot(ByVal index As Long, ByVal PetSlot As Byte)
Dim i As Byte

player(index).Pet(PetSlot).Experience = 0
'Player(index).Pet(PetSlot).Name = ""
player(index).Pet(PetSlot).NumPet = 0
player(index).Pet(PetSlot).level = 0
player(index).Pet(PetSlot).points = 0
For i = 1 To Stats.Stat_Count - 1
    player(index).Pet(PetSlot).StatsAdd(i) = 0
Next
End Sub

Public Sub SharePetExp(ByVal index As Long, ByVal PetSlot As Byte, ByVal exp As Long, Optional ByVal Percent As Double = 50)
Dim PetExp As Long
If Not (exp > 0) Then Exit Sub

If Not (PetSlot > 0 And PetSlot <= MAX_PLAYER_PETS) Then Exit Sub

If Not (player(index).Pet(PetSlot).NumPet > 0) Then Exit Sub

    'Pet exp
    PetExp = Round(CLng(CDbl(exp) * (Percent / 100)))
    If PetExp > 0 Then
        Call GivePetEXP(index, PetExp)
    End If
    
    
    'Player Exp
    exp = exp - PetExp
    
    If Not (exp > 0) Then Exit Sub

    GivePlayerEXP index, exp

End Sub

Public Sub LeavePet(ByVal index As Long, ByVal PetSlot As Byte)
'Dim CumExp As Long
Dim i As Byte

'CumExp = 0

'For i = 1 To Player(index).Pet(PetSlot).level
    'CumExp = CumExp + GetPlayerPetExpByLevel(Player(index).Pet(PetSlot).NumPet, i)
    'Exit if level > MAX_LEVEL
    'If i > MAX_LEVELS Then
        'i = Player(index).Pet(PetSlot).level + 1
    'End If
'Next

If TempPlayer(index).TempPet.TempPetSlot > 0 Then
    If PetDisband(index, GetPlayerMap(index), False) = False Then Exit Sub
End If

Call PlayerMsg(index, GetTranslation("Abandonas a tu") & " " & Trim$(NPC(Pet(player(index).Pet(PetSlot).NumPet).npcnum).TranslatedName), Yellow, , False)
Call ResetPlayerPetSlot(index, PetSlot)
'Call GivePlayerEXP(index, CumExp / 2)
Call SendPetData(index, PetSlot)

End Sub



Public Function GetMapPetOwner(ByVal mapnum As Long, ByVal mapnpcnum As Long) As Long
Dim Owner As Long
GetMapPetOwner = 0

Owner = MapNpc(mapnum).NPC(mapnpcnum).PetData.Owner

If Owner > 0 Then
    GetMapPetOwner = Owner
End If

End Function

Public Function PlayerHasPetInMap(ByVal index As Long) As Long
Dim mapnum As Long
Dim mapnpcnum As Long
PlayerHasPetInMap = 0

If index > 0 And index <= MAX_PLAYERS Then
    Dim PetSlot As Long
    PetSlot = TempPlayer(index).TempPet.TempPetSlot
    If PetSlot > 0 And PetSlot < MAX_MAP_NPCS Then
        PlayerHasPetInMap = TempPlayer(index).TempPet.TempPetSlot
    End If
End If
            
End Function

Public Function ChoosePetSpellingMethod(ByVal index As Long, ByVal mapnpcnum As Long, ByVal SpellSlotNum As Long, ByVal spellnum As Long) As Boolean

Dim mapnum As Long

ChoosePetSpellingMethod = False

'If not autocast, exit, if not heal type, exit too
If Spell(spellnum).range <> 0 Or (Spell(spellnum).Type <> SPELL_TYPE_HEALHP And Spell(spellnum).Type <> SPELL_TYPE_HEALMP) Then Exit Function

If index = 0 Then Exit Function

'If PlayerHasPetInMap(index) = False Then Exit Function

mapnum = GetPlayerMap(index)

'subscript 9
If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Then Exit Function

'subscript 9
If Not (MapNpc(mapnum).NPC(mapnpcnum).Num > 0) Then Exit Function

'player hasn't pet?
If GetMapPetOwner(mapnum, mapnpcnum) <> index Then Exit Function

Select Case ComparePetAndOwnerVital(index, mapnpcnum, SpellSlotNum)
    Case 1
        'Player has better vital
        If RAND(0, 4) = 2 Then
            Call PetSpellItself(mapnum, mapnpcnum, SpellSlotNum)
            ChoosePetSpellingMethod = True
            Exit Function
        End If
        
    Case -1
        'Pet has better vital
        If RAND(0, 4) = 3 Then
        Call PetSpellOwner(mapnpcnum, index, SpellSlotNum)
        ChoosePetSpellingMethod = True
        End If
    Case 0
        'Choose Randomly
        If RAND(0, 4) = 0 Then
            Select Case RAND(1, 2)
            Case 1
                Call PetSpellItself(mapnum, mapnpcnum, SpellSlotNum)
            Case 2
                Call PetSpellOwner(mapnpcnum, index, SpellSlotNum)
            End Select
            ChoosePetSpellingMethod = True
        End If
    Case 2
        'max vital, do not heal
        ChoosePetSpellingMethod = True
        Exit Function
    Case Else
        'error or spell is not heal type
End Select

End Function

Public Function ComparePetAndOwnerVital(ByVal index As Long, ByVal mapnpcnum As Long, ByVal SpellSlotNum As Long) As Integer
Dim mapnum As Long

Dim RVital1 As Double
Dim RVital2 As Double
Dim vital As Vitals
Dim PlayerVital As Double
Dim PetVital As Double
Dim spellnum As Long

ComparePetAndOwnerVital = 3

' Check for subscript out of range
If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or IsPlaying(index) = False Then
    Exit Function
End If

' Check for subscript out of range
If MapNpc(GetPlayerMap(index)).NPC(mapnpcnum).Num <= 0 Then
    Exit Function
End If
   
If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Function
        
' The Variables
mapnum = GetPlayerMap(index)
spellnum = NPC(MapNpc(mapnum).NPC(mapnpcnum).Num).Spell(SpellSlotNum)

If Not (spellnum > 0 And spellnum <= MAX_SPELLS And mapnum > 0 And mapnum) <= MAX_MAPS Then Exit Function

Select Case Spell(spellnum).Type
Case SPELL_TYPE_HEALHP
    vital = HP
Case SPELL_TYPE_HEALMP
    vital = MP
Case Else
    Exit Function
End Select
PlayerVital = CDbl(GetPlayerVital(index, vital))
PetVital = CDbl(MapNpc(mapnum).NPC(mapnpcnum).vital(vital))

If PlayerVital <= 0 Then
    PlayerVital = 1
End If
If PetVital <= 0 Then
    PetVital = 1
End If

'1: Player has better vital, -1: npc has better vital, 0: equal, 2: equal but both max, 3: error
RVital1 = CDbl(GetPlayerMaxVital(index, vital)) / PlayerVital
RVital2 = CDbl(GetNpcMaxVital(mapnum, mapnpcnum, vital)) / PetVital

If RVital1 > 1 And vital = MP Then 'we don't want pet recupering mp
    ComparePetAndOwnerVital = -1 'fake the system
    Exit Function
ElseIf RVital1 = 1 And vital = MP Then
    ComparePetAndOwnerVital = 2
    Exit Function
End If

If RVital2 > RVital1 Then
    ComparePetAndOwnerVital = 1
ElseIf RVital2 < RVital1 Then
    ComparePetAndOwnerVital = -1
ElseIf RVital2 = RVital1 And RVital1 > 1 Then  'Same HP but not max
    ComparePetAndOwnerVital = 0
ElseIf RVital2 = RVital1 And RVital1 = 1 Then 'Same HP but max
    ComparePetAndOwnerVital = 2
End If
        
End Function

Sub PetAttack(ByVal index As Long)
    If TempPlayer(index).TempPet.TempPetSlot < 1 Then Exit Sub
    
    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPet.TempPetSlot).TargetType = TempPlayer(index).TargetType
    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPet.TempPetSlot).Target = TempPlayer(index).Target
    TempPlayer(index).TempPet.PetHasOwnTarget = TempPlayer(index).Target
End Sub

Sub ParsePetCommand(ByVal index As Long, ByVal PetCommand As PetCommandsType)
Select Case PetCommand
Case ePetSpawn
    Call SpawnPet(index, GetPlayerMap(index))
Case ePetAttack
    Call PetAttack(index)
Case ePetFollow
    Call PetFollowOwner(index)
Case ePetWander
    Call PetWander(index)
Case ePetDisband
    Call PetDisband(index, GetPlayerMap(index), False)
End Select
End Sub

Function IsPetTargetted(ByVal index As Long) As Boolean
    If TempPlayer(index).Target > 0 Then
        If TempPlayer(index).TargetType = TARGET_TYPE_NPC Then
            If TempPlayer(index).Target = PlayerHasPetInMap(index) Then
                IsPetTargetted = True
            End If
        End If
    End If
End Function

Public Function ResetPlayerPetPoints(ByVal index As Long, ByVal PetSlot As Byte) As Long
Dim i As Byte, sum As Long
ResetPlayerPetPoints = 0
'PlayerUnequip (index)
sum = 0

For i = 1 To Stats.Stat_Count - 1
    Do While player(index).Pet(PetSlot).StatsAdd(i) > 0
        player(index).Pet(PetSlot).StatsAdd(i) = player(index).Pet(PetSlot).StatsAdd(i) - 1
        sum = sum + 1
    Loop
Next

ResetPlayerPetPoints = sum
    
End Function

Public Function GetPlayerPetSlot(ByVal index As Long) As Byte
    GetPlayerPetSlot = TempPlayer(index).TempPet.ActualPet
End Function

Public Sub SetPetTarget(ByVal index As Long, ByVal PetSlot As Byte)
    Dim mapnum As Long
    mapnum = GetPlayerMap(index)
    If TempPlayer(index).TempPet.PetHasOwnTarget > 0 Then
    'I'm not sure what they intended on putting here..
    Debug.Print
    
    Else
        Dim i As Integer
        Dim Chosen As Integer
        Dim Chance As Single
        Chance = 10
        Chosen = 0
        For i = 1 To TempMap(mapnum).npc_highindex
            If MapNpc(mapnum).NPC(i).TargetType = TARGET_TYPE_PLAYER Then
                If MapNpc(mapnum).NPC(i).Target = index Then
                    If IsinRange(2, GetPlayerX(index), GetPlayerY(index), MapNpc(mapnum).NPC(i).X, MapNpc(mapnum).NPC(i).Y) Then
                        Dim auxchance As Single
                        auxchance = GetNPCSFightChance(mapnum, PlayerHasPetInMap(index), i)
                        If auxchance < Chance Then
                            Chance = auxchance
                            Chosen = i
                        End If
                    End If
                End If
            End If
        Next
    
    End If
    
    If Chosen > 0 Then
        TempPlayer(index).TempPet.PetHasOwnTarget = Chosen
        MapNpc(mapnum).NPC(PlayerHasPetInMap(index)).Target = Chosen
        MapNpc(mapnum).NPC(PlayerHasPetInMap(index)).TargetType = TARGET_TYPE_NPC
    End If

End Sub


Function PetExists(ByVal PetNum As Long) As Boolean
If LenB(Trim$(Pet(PetNum).Name)) > 0 And Asc(Pet(PetNum).Name) <> 0 Then
    PetExists = True
End If
End Function


Function IsMapNPCaPet(ByVal mapnum As Long, ByVal mapnpcnum As Long) As Boolean
    If mapnum = 0 Then Exit Function
    If mapnpcnum = 0 Then Exit Function
    
    If MapNpc(mapnum).NPC(mapnpcnum).PetData.Owner > 0 Then
        IsMapNPCaPet = True
    End If
End Function

Sub ResetPetTarget(ByVal index As Long)
    If Not (0 > index > MAX_PLAYERS) Then Exit Sub
    
    TempPlayer(index).TempPet.PetHasOwnTarget = 0
End Sub






