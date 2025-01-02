Attribute VB_Name = "modPet"


Public Type PlayerPetRec
    'SpriteNum As Byte
    Name As String * 50
    'Link to super class
    NumPet As Byte
    StatsAdd(1 To Stats.Stat_Count - 1) As Byte
    points As Integer
    Experience As Long
    Level As Long
    
    'Server side
    CurrentHP As Long
    
End Type

Public Type MapPetRec
    owner As Long
    Name As String
End Type

Public Type PetRec
    Name As String * NAME_LENGTH
    NPCNum As Long
    'Requeriments
    TamePoints As Integer
    'Progressions
    ExpProgression As Byte
    pointsprogression As Byte
    MaxLevel As Long
End Type

Public Type PetCache
    Pet(1 To MAX_MAP_NPCS) As Long
    UpperBound As Long
End Type


'makes the pet follow its owner
Public Sub PetFollowOwner(ByVal index As Long)
    If TempPlayer(index).TempPetSlot < 1 Then Exit Sub
    
    mapnpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).targetType = 1
    mapnpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).target = index
End Sub

'makes the pet wander around the map
Public Sub PetWander(ByVal index As Long)
    If TempPlayer(index).TempPetSlot < 1 Then Exit Sub

    mapnpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).targetType = TARGET_TYPE_NONE
    mapnpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).target = 0
End Sub

'Clear the npc from the map
Public Function PetDisband(ByVal index As Long, ByVal mapnum As Long, Optional ByVal WaitForSpawn As Boolean = True) As Boolean
    Dim i As Long
    Dim j As Long
    PetDisband = True

    If TempPlayer(index).TempPetSlot < 1 Then
        PetDisband = False
        Exit Function
    End If
    
    'Cache the Pets for players logging on [Remove Number from array]
    'THIS IS KINDA SLOW (EVEN WITHOUT TESTING, LOL), MAY HAVE TO CONVERT TO LINKED LIST FOR SPEED
    For i = 1 To PetMapCache(mapnum).UpperBound
        If PetMapCache(mapnum).Pet(i) = TempPlayer(index).TempPetSlot Then
            If PetMapCache(mapnum).UpperBound > 1 Then
                For j = PetMapCache(mapnum).UpperBound To i + 1 Step -1
                    PetMapCache(mapnum).Pet(j - 1) = PetMapCache(mapnum).Pet(j)
                Next
            Else
                PetMapCache(mapnum).Pet(1) = 0
            End If
            
            PetMapCache(mapnum).UpperBound = PetMapCache(mapnum).UpperBound - 1
            Exit For
        End If
    Next
    
    If mapnpc(mapnum).NPC(TempPlayer(index).TempPetSlot).Vital(Vitals.HP) > 0 Then
        Player(index).Pet(TempPlayer(index).ActualPet).CurrentHP = mapnpc(mapnum).NPC(TempPlayer(index).TempPetSlot).Vital(Vitals.HP)
    Else
        Player(index).Pet(TempPlayer(index).ActualPet).CurrentHP = 0
    End If
    
    Call ClearSingleMapNpc(TempPlayer(index).TempPetSlot, mapnum)
    Map(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot) = 0
    TempPlayer(index).TempPetSlot = 0
    
    'Check if pet must wait until next spawn
    If WaitForSpawn = True Then
        TempPlayer(index).PetSpawnWait = GetTickCount
    End If
    
    'Reset Target
    TempPlayer(index).PetHasOwnTarget = 0
    

    're-warp the players on the map
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                'Call PlayerWarp(i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i))
                'SendPlayerData index
                'SendMap index, GetPlayerMap(index)
                'SendMapNpcsTo index, GetPlayerMap(index)
                Call RefreshMapNPCS(i)
            End If
        End If
    Next
End Function

Public Sub SpawnPet(ByVal index As Long, ByVal mapnum As Long)
    Dim PlayerMap As Long
    Dim i As Integer
    Dim PetSlot As Byte
    Dim PlayerPet As Byte
    Dim UntilTime As Long
    
    'Prevent multiple pets for the same owner
    If TempPlayer(index).TempPetSlot > 0 Then Exit Sub
    
    'slot, 1 to MAX_PLAYER_PETS
    PlayerPet = TempPlayer(index).ActualPet
    
    'Prevent player out of range slots
    If PlayerPet <= 0 Or PlayerPet > MAX_PLAYER_PETS Then Exit Sub
    
    UntilTime = GetTickCount - TempPlayer(index).PetSpawnWait - GetPetSpawnTime(Player(index).Pet(PlayerPet).NumPet)
    'Check if SpawnWait Finished
    'If Not (GetTickCount > TempPlayer(index).PetSpawnWait + GetPetSpawnTime(Player(index).Pet(PlayerPet).NumPet)) Then
    If UntilTime <= 0 Then
        Call PlayerMsg(index, "You can't summon your pet for " & Round(Abs(UntilTime) / 1000, 0) & " seconds!", BrightRed)
        Exit Sub
    End If
    
    'Prevent spawning inexistent pet
    If Player(index).Pet(TempPlayer(index).ActualPet).NumPet < 1 Or Player(index).Pet(TempPlayer(index).ActualPet).NumPet > MAX_PETS Then Exit Sub
    
    
    
    PlayerMap = GetPlayerMap(index)
    PetSlot = 0
    
    'Prevent Boundries
    Select Case Player(index).Dir
    Case 0
        If Player(index).y = Map(PlayerMap).MaxY Then Exit Sub
    Case 1
        If Player(index).y = 0 Then Exit Sub
    Case 2
        If Player(index).x = Map(PlayerMap).MaxX Then Exit Sub
    Case 3
        If Player(index).x = 0 Then Exit Sub
    End Select
    
    For i = 1 To MAX_MAP_NPCS
        'If Map(PlayerMap).Npc(i) = 0 Then
         If mapnpc(PlayerMap).NPC(i).SpawnWait = 0 And mapnpc(PlayerMap).NPC(i).Num = 0 Then
            PetSlot = i
            Exit For
         End If
    Next
    
    If PetSlot = 0 Then
        Call PlayerMsg(index, "The map is too crowded for you to call on your pet!", Red)
        Exit Sub
    End If

    'create the pet for the map
    Map(PlayerMap).NPC(PetSlot) = Pet(Player(index).Pet(PlayerPet).NumPet).NPCNum 'pet npc number
    mapnpc(PlayerMap).NPC(PetSlot).Num = Pet(Player(index).Pet(PlayerPet).NumPet).NPCNum  'pet npc number
    'set its Pet Data
    mapnpc(PlayerMap).NPC(PetSlot).IsPet = YES
    'mapnpc(PlayerMap).NPC(PetSlot).PetData.Name = GetPlayerName(index) & "'s " & NPC(200).Name
    mapnpc(PlayerMap).NPC(PetSlot).PetData.owner = index
    
    'If Pet doesn't exist with player, link it to the player
    'If Player(index).Pet.SpriteNum <> NPC(200).Sprite Then
        'Player(index).Pet.SpriteNum = NPC(200).Sprite
        'Player(index).Pet.Name = GetPlayerName(index) & "'s " & NPC(200).Name
    'End If
    
    TempPlayer(index).TempPetSlot = PetSlot
       
    'cache the map for sending
    Call MapCache_Create(PlayerMap)

    'Cache the Pets for players logging on [Add new Number to array]
    PetMapCache(PlayerMap).UpperBound = PetMapCache(PlayerMap).UpperBound + 1
    PetMapCache(PlayerMap).Pet(PetMapCache(PlayerMap).UpperBound) = PetSlot
    
    If PetMapCache(Player(index).Map).UpperBound > 0 Then
        For i = 1 To PetMapCache(Player(index).Map).UpperBound
            Call NPCCache_Create(index, Player(index).Map, PetMapCache(Player(index).Map).Pet(i))
        Next
    End If

    Select Case GetPlayerDir(index)
        Case DIR_UP
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index), GetPlayerY(index) + 1)
        Case DIR_DOWN
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index), GetPlayerY(index) - 1)
        Case DIR_LEFT
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index) + 1, GetPlayerY(index))
        Case DIR_RIGHT
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index) - 1, GetPlayerY(index))
    End Select
    
    're-warp the players on the map
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(index) Then
                Call RefreshMapNPCS(i, True)
                'Call PlayerWarp(i, PlayerMap, GetPlayerX(i), GetPlayerY(i))
            End If
        End If
    Next
    
End Sub

Public Function CheckFreePetSlots(ByVal index As Long) As Integer
'-1: There aren't free slots, >= 1: Index of first free slot
Dim i As Byte

CheckFreePetSlots = -1
For i = 1 To MAX_PLAYER_PETS

    If Player(index).Pet(i).Name = "" Or Player(index).Pet(i).NumPet = 0 Then
        CheckFreePetSlots = i
        Exit Function
    End If
    
Next

End Function


Public Sub CheckPlayerTame(ByVal index As Long)

Dim Slot As Integer
Dim PetIndex As Integer

Slot = CheckFreePetSlots(index)

If Not (Slot > 0) Then
    Call PlayerMsg(index, "You don't have any free slots", BrightRed)
    Exit Sub
End If

If TempPlayer(index).targetType <> TARGET_TYPE_NPC Then
    Call PlayerMsg(index, "You don't have any NPC fixed", BrightRed)
    Exit Sub
End If

If mapnpc(GetPlayerMap(index)).NPC(TempPlayer(index).target).IsPet Then
    Call PlayerMsg(index, "The goal is a pet", BrightRed)
    Exit Sub
End If

PetIndex = IsNPCaPet(mapnpc(GetPlayerMap(index)).NPC(TempPlayer(index).target).Num)

If PetIndex <= 0 Then
    Call PlayerMsg(index, "The NPC is not tame", BrightRed)
    Exit Sub
End If

If Not (GetPlayerTamePoints(index, Pet(PetIndex).NPCNum) >= Pet(PetIndex).TamePoints) Then
    Call PlayerMsg(index, "You don't have enough taming points", BrightRed)
    Exit Sub
End If

'Agregar pet a slot libre
Call AddPlayerPet(index, CByte(Slot), CByte(PetIndex))
Call PlayerMsg(index, Trim$(NPC(Pet(PetIndex).NPCNum).Name) & " he has joined your team!", BrightGreen)
Call KillNpc(GetPlayerMap(index), TempPlayer(index).target)
Call SendPetData(index, TempPlayer(index).ActualPet)


End Sub

Public Function IsNPCaPet(ByVal NPCNum As Long) As Integer
'-1: not a pet, >0 : pet index
Dim i As Byte

IsNPCaPet = -1

For i = 1 To MAX_PETS
    If Pet(i).NPCNum = NPCNum Then
        IsNPCaPet = i
        Exit Function
    End If
Next

End Function

Public Function GetPlayerPetStat(ByVal index As Long, ByVal stat As Byte) As Byte
    GetPlayerPetStat = Player(index).Pet(TempPlayer(index).ActualPet).StatsAdd(stat)
End Function

Public Function GetPlayerPetTotalStat(ByVal index As Long, ByVal stat As Byte) As Byte
Dim result As Integer
If Player(index).Pet(TempPlayer(index).ActualPet).NumPet > 0 Then

    result = Player(index).Pet(TempPlayer(index).ActualPet).StatsAdd(stat) + NPC(Pet(Player(index).Pet(TempPlayer(index).ActualPet).NumPet).NPCNum).stat(stat)
    If result > 255 Then
        result = 255
    End If
    
    GetPetTotalStat = CByte(result)
End If
End Function

Public Sub SetPlayerPetStat(ByVal index As Long, ByVal stat As Byte, ByVal number As Byte)

If number > 255 Or number < 0 Then Exit Sub

Player(index).Pet(TempPlayer(index).ActualPet).StatsAdd(stat) = number

End Sub

Public Function GetPlayerPetLevel(ByVal index As Long) As Long
    GetPlayerPetLevel = Player(index).Pet(TempPlayer(index).ActualPet).Level
End Function

Public Function SetPlayerPetLevel(ByVal index As Long, ByVal Level As Long) As Boolean
    SetPlayerPetLevel = False
    If Level > MAX_LEVELS Or Level > Pet(Player(index).Pet(TempPlayer(index).ActualPet).NumPet).MaxLevel Then Exit Function
    Player(index).Pet(TempPlayer(index).ActualPet).Level = Level
    SetPlayerPetLevel = True
End Function

Public Function GetPlayerPetExp(ByVal index As Long) As Long
    GetPlayerPetExp = Player(index).Pet(TempPlayer(index).ActualPet).Experience
End Function

Public Sub SetPlayerPetExp(ByVal index As Long, ByVal Exp As Long)
    Player(index).Pet(TempPlayer(index).ActualPet).Experience = Exp
End Sub

Public Function GetPlayerPetNextLevel(ByVal index As Long) As Long
    GetPlayerPetNextLevel = (50 / 6) * ((GetPlayerPetLevel(index) + Pet(Player(index).Pet(TempPlayer(index).ActualPet).NumPet).ExpProgression) ^ 3 - (6 * (GetPlayerPetLevel(index) + Pet(Player(index).Pet(TempPlayer(index).ActualPet).NumPet).ExpProgression) ^ 2) + 17 * (GetPlayerPetLevel(index) + Pet(Player(index).Pet(TempPlayer(index).ActualPet).NumPet).ExpProgression) - 12) + 50
End Function

Public Function GetPlayerPetExpByLevel(ByVal PetNum As Byte, ByVal Level As Long)

GetPlayerPetExpByLevel = (50 / 6) * ((Level + Pet(PetNum).ExpProgression) ^ 3 - (6 * (Level + Pet(PetNum).ExpProgression) ^ 2) + 17 * (Level) + Pet(PetNum).ExpProgression - 12)
End Function

Public Function GetPlayerPetMaxLevel(ByVal index As Long) As Long
    GetPlayerPetMaxLevel = Pet(Player(index).Pet(TempPlayer(index).ActualPet).NumPet).MaxLevel
End Function
Public Function GetPlayerPetPOINTS(ByVal index As Long) As Integer
    GetPlayerPetPOINTS = Player(index).Pet(TempPlayer(index).ActualPet).points
End Function
Public Sub SetPlayerPetPOINTS(ByVal index As Long, ByVal points As Integer)
    Player(index).Pet(TempPlayer(index).ActualPet).points = points
End Sub
Public Function GetPetNextPOINTS(ByVal index As Long, ByVal PetNum As Byte)
    GetPetNextPOINTS = PointsPerLevel(PetNum, Player(index).Pet(TempPlayer(index).ActualPet).Level)
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
            Exit Sub
        End If
        
        PointsWon = GetPetNextPOINTS(index, Player(index).Pet(TempPlayer(index).ActualPet).NumPet)
        Call SetPlayerPetPOINTS(index, GetPlayerPetPOINTS(index) + PointsWon)
        PlayerMsg index, PointsWon & "points won!", Blue
        Call SetPlayerPetExp(index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            PlayerMsg index, "Your pet has come up" & level_count & "level!", Brown
        Else
            'plural
            PlayerMsg index, "Your pet has come up" & level_count & " levels!", Brown
        End If
        'SendEXP index
        SendPetData index, TempPlayer(index).ActualPet
        'SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seLevelUp, 1
    End If
End Sub

Public Sub GivePetEXP(ByVal index As Long, ByVal Exp As Long)
    ' give the exp
    If TempPlayer(index).ActualPet < 1 Or TempPlayer(index).ActualPet > MAX_PLAYER_PETS Then Exit Sub
    
    Call SetPlayerPetExp(index, Exp + GetPlayerPetExp(index))
    'SendEXP index
    SendPetData index, TempPlayer(index).ActualPet
    SendActionMsg GetPlayerMap(index), "+" & Exp & " EXP", White, 1, (mapnpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).x * 32), (mapnpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).y * 32)
    ' check if we've leveled
    CheckPlayerPetLevelUp index
End Sub

Public Function GetPetSpawnTime(ByVal PetNum As Byte) As Long
Dim BaseTime As Long

If PetNum < 1 Or PetNum > MAX_PETS Then Exit Function

'BaseTime = 10000 '10 seconds

'Process base time
BaseTime = 1000 * NPC(Pet(PetNum).NPCNum).SpawnSecs + 200 * Pet(PetNum).pointsprogression
'* (Pet(PetNum).pointsprogression + 1) * 2

'erase this
'BaseTime = 3000

GetPetSpawnTime = BaseTime
End Function

Function PointsPerLevel(ByVal PetNum As Byte, ByVal Level As Long) As Byte
Dim result As Double
Dim pntsPro As Byte
Dim logarithm As Double
Dim ProFactor As Double
    
pntsPro = Pet(PetNum).pointsprogression
    If pntsPro > 0 And pntsPro <= MAX_PET_POINTS_PERLVL Then
        ProFactor = GPN(pntsPro, Pet(PetNum).MaxLevel)
        logarithm = LogB((MAX_PET_POINTS_PERLVL - 1) / ProFactor, CDbl(Pet(PetNum).MaxLevel))
        result = ProFactor * (CDbl(Level) ^ logarithm) + 1
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

Public Function GetPlayerTamePoints(ByVal index As Long, ByVal NPCNum As Long) As Integer

GetPlayerTamePoints = CInt(GetPlayerStat(index, willpower) + GetPlayerLevel(index) + (Player(index).NPC(NPCNum).Kills / 100))

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

Player(index).Pet(PetSlot).Name = GetPlayerName(index) & "'s " & NPC(Pet(PetNum).NPCNum).Name
Player(index).Pet(PetSlot).NumPet = PetNum
Player(index).Pet(PetSlot).Level = 1

End Sub

Public Sub InitPlayerPets(ByVal index As Long)
    TempPlayer(index).ActualPet = 1
    TempPlayer(index).TempPetSlot = 0
    TempPlayer(index).PetExpPercent = 50
    
    'Player(index).Pet(TempPlayer(index).ActualPet).SpawnWait = GetTickCount
End Sub

Public Sub ResetPlayerPetSlot(ByVal index As Long, ByVal PetSlot As Byte)
Dim i As Byte

Player(index).Pet(PetSlot).Experience = 0
Player(index).Pet(PetSlot).Name = ""
Player(index).Pet(PetSlot).NumPet = 0
Player(index).Pet(PetSlot).Level = 0
Player(index).Pet(PetSlot).points = 0
For i = 1 To Stats.Stat_Count - 1
    Player(index).Pet(PetSlot).StatsAdd(i) = 0
Next
End Sub

Public Sub SharePetExp(ByVal index As Long, ByVal PetSlot As Byte, ByVal Exp As Long, Optional ByVal Percent As Double = 50, Optional ByVal Share As Boolean = True)
Dim PetExp As Long
If Not (Exp > 0) Then Exit Sub

If Not (PetSlot > 0 And PetSlot <= MAX_PLAYER_PETS) Then Exit Sub

If Not (Player(index).Pet(PetSlot).NumPet > 0) Then Exit Sub

    'Pet exp
    PetExp = Round(CLng(CDbl(Exp) * (Percent / 100)))
    If PetExp > 0 Then
        Call GivePetEXP(index, PetExp)
    End If
    
    
    'Player Exp
    Exp = Exp - PetExp
    
    If Not (Exp > 0) Then Exit Sub
    
    If TempPlayer(index).inParty > 0 And Share Then
        ' pass through party sharing function
        Party_ShareExp TempPlayer(index).inParty, Exp, index
    Else
        ' no party - keep exp for self
        GivePlayerEXP index, Exp
    End If

End Sub

Public Sub LeavePet(ByVal index As Long, ByVal PetSlot As Byte)
Dim CumExp As Long
Dim i As Byte

CumExp = 0

For i = 1 To Player(index).Pet(PetSlot).Level
    CumExp = CumExp + GetPlayerPetExpByLevel(Player(index).Pet(PetSlot).NumPet, i)
    'Exit if level > MAX_LEVEL
    If i > MAX_LEVELS Then
        i = Player(index).Pet(PetSlot).Level + 1
    End If
Next

If TempPlayer(index).TempPetSlot > 0 Then
    If PetDisband(index, GetPlayerMap(index), False) = False Then Exit Sub
End If

Call PlayerMsg(index, "You abandon your" & Trim$(NPC(Pet(Player(index).Pet(PetSlot).NumPet).NPCNum).Name) & " and you get half of his Exp:" & CumExp / 2 & " points!", Yellow)
Call ResetPlayerPetSlot(index, PetSlot)
Call GivePlayerEXP(index, CumExp / 2)
Call SendPetData(index, PetSlot)

End Sub



Public Function GetMapPetOwner(ByVal mapnum As Long, ByVal MapNPCNum As Long) As Long
Dim auxOwner As Long
GetMapPetOwner = 0

If mapnpc(mapnum).NPC(MapNPCNum).IsPet = YES Then
    auxOwner = mapnpc(mapnum).NPC(MapNPCNum).PetData.owner
    If auxOwner > 0 Then
        If IsPlaying(auxOwner) Then
            GetMapPetOwner = auxOwner
        End If
    End If
End If

End Function

Public Function PlayerHasPetInMap(ByVal index As Long) As Boolean
Dim mapnum As Long
Dim MapNPCNum As Long
PlayerHasPetInMap = False

    If index > 0 Then
        If IsPlaying(index) Then
            mapnum = GetPlayerMap(index)
            MapNPCNum = TempPlayer(index).TempPetSlot
            If mapnum < 1 Or mapnum > MAX_MAPS Or MapNPCNum < 1 Or MapNPCNum > MAX_MAP_NPCS Then Exit Function
            
            If index = mapnpc(mapnum).NPC(MapNPCNum).PetData.owner Then
                PlayerHasPetInMap = True
                Exit Function
            End If
        End If
    End If
            
End Function

Public Function ChoosePetSpellingMethod(ByVal index As Long, ByVal MapNPCNum As Long, ByVal SpellSlotNum As Long, ByVal SpellNum As Long) As Boolean

Dim mapnum As Long

ChoosePetSpellingMethod = False

'If not autocast, exit, if not heal type, exit too
If Spell(SpellNum).Range <> 0 Or (Spell(SpellNum).Type <> SPELL_TYPE_HEALHP And Spell(SpellNum).Type <> SPELL_TYPE_HEALMP) Then Exit Function

If index = 0 Then Exit Function

If PlayerHasPetInMap(index) = False Then Exit Function

mapnum = GetPlayerMap(index)

'subscript 9
If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Then Exit Function

'subscript 9
If Not (mapnpc(mapnum).NPC(MapNPCNum).Num > 0) Then Exit Function

'player hasn't pet?
If GetMapPetOwner(mapnum, MapNPCNum) <> index Then Exit Function


Select Case ComparePetAndOwnerVital(index, MapNPCNum, SpellSlotNum)
    Case 1
        'Player has better vital
        Call PetSpellItself(mapnum, MapNPCNum, SpellSlotNum)
        ChoosePetSpellingMethod = True
    Case -1
        'Pet has better vital
        Call PetSpellOwner(MapNPCNum, index, SpellSlotNum)
        ChoosePetSpellingMethod = True
    Case 0
        'Choose Randomly
        Select Case RAND(1, 2)
        Case 1
            Call PetSpellItself(mapnum, MapNPCNum, SpellSlotNum)
        Case 2
            Call PetSpellOwner(MapNPCNum, index, SpellSlotNum)
        End Select
        ChoosePetSpellingMethod = True
    Case 2
        'max hp, do not heal
        ChoosePetSpellingMethod = True
        Exit Function
    Case Else
        'error or spell is not heal type
End Select

End Function

Public Function ComparePetAndOwnerVital(ByVal index As Long, ByVal MapNPCNum As Long, ByVal SpellSlotNum As Long) As Integer
Dim mapnum As Long

Dim RVital1 As Double
Dim RVital2 As Double
Dim Vital As Vitals
Dim PlayerVital As Double
Dim PetVital As Double

ComparePetAndOwnerVital = 3

' Check for subscript out of range
If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or IsPlaying(index) = False Then
    Exit Function
End If

' Check for subscript out of range
If mapnpc(GetPlayerMap(index)).NPC(MapNPCNum).Num <= 0 Then
    Exit Function
End If
   
If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Function
        
' The Variables
mapnum = GetPlayerMap(index)
SpellNum = NPC(mapnpc(mapnum).NPC(MapNPCNum).Num).Spell(SpellSlotNum)

If Not (SpellNum > 0 And SpellNum <= MAX_SPELLS And mapnum > 0 And mapnum) <= MAX_MAPS Then Exit Function

Select Case Spell(SpellNum).Type
Case SPELL_TYPE_HEALHP
    Vital = HP
Case SPELL_TYPE_HEALMP
    Vital = mp
Case Else
    Exit Function
End Select
PlayerVital = CDbl(GetPlayerVital(index, Vital))
PetVital = CDbl(mapnpc(mapnum).NPC(MapNPCNum).Vital(Vital))

If PlayerVital <= 0 Then
    PlayerVital = 1
End If
If PetVital <= 0 Then
    PetVital = 1
End If

'1: Player has better vital, -1: npc has better vital, 0: equal, 2: equal but both max, 3: error
RVital1 = CDbl(GetPlayerMaxVital(index, Vital)) / PlayerVital
RVital2 = CDbl(GetNpcMaxVital(mapnpc(mapnum).NPC(MapNPCNum).Num, Vital, index)) / PetVital

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









