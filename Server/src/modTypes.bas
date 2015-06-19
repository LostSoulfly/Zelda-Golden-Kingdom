Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures

Public map(1 To MAX_MAPS) As ServerMapRec
Public MapCache(1 To MAX_MAPS) As Cache
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public player(1 To MAX_PLAYERS) As PlayerRec
Public Bank(1 To MAX_PLAYERS) As BankRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Class() As ClassRec
Public item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS) As MapDataRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Movements(1 To MAX_MOVEMENTS) As MovementRec
Public Actions(1 To MAX_ACTIONS) As ActionRec
'Public PetMapCache(1 To MAX_MAPS) As PetCache
Public Pet(1 To MAX_PETS) As PetRec



Public AuxPlayer As PlayerRec
Public AuxBank As BankRec
Public AuxGuild As GuildRec

Public Options As OptionsRec


Private Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    Website As String
    DisableAdmins As String
    Update As String
    Instructions As String
    ExpMultiplier As Long
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
End Type

Private Type Cache
    Data() As Byte
End Type

Private Type BankRec
    item(1 To MAX_BANK) As PlayerInvRec
End Type

Public Type HotbarRec
    slot As Long
    sType As Byte
End Type

Public Type ProjectileRec
    TravelTime As Long
    Direction As Long
    X As Long
    Y As Long
    Pic As Long
    range As Long
    Damage As Long
    Speed As Long
    Depth As Byte
End Type

Public Type PlayerDoorRec
    state As Byte
End Type



Public Type PlayerRec
    ' Account
    login As String * ACCOUNT_LENGTH
    password As String * NAME_LENGTH
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    level As Byte
    exp As Long
    Access As Byte
    'Codification : 0: Normal Player, 1: PK, 2: HERO
    PK As Byte
    
    GuildFileId As Long
    GuildMemberId As Long
    
    ' Vitals
    vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    stat(1 To Stats.Stat_Count - 1) As Byte
    points As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    map As Long
    X As Byte
    Y As Byte
    dir As Byte
    
    'ALATAR
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec
    '/ALATAR
    
    'PlayerDoors(1 To MAX_DOORS) As PlayerDoorRec
    Visible As Long
    
    'Pet system
    Pet(1 To MAX_PLAYER_PETS) As PlayerPetRec
    
    'Npc kills info
    NPCKills As Long
    
    'Triforce
    triforce(1 To TriforceType.TriforceType_Count - 1) As Boolean
    
    HeroPoints As Long
    PKPoints As Long
    
    'Ice System
    onIce As Boolean
    IceDir As Byte
    
    'Rupee System
    RupeeBags As Byte
    
    'Custom Sprite
    CustomSprite As Byte
    
    'Max inventory weight
    MaxWeight As Long
    
    SafeMode As Boolean
    
    PKKillPoints As Single
    HeroKillPoints As Single

    NeutralEnabled As Boolean
    
    BonusPoints As Long
    
    state As Byte
    
    'Kill Counter
    Kill As Long
    Dead As Long
    NpcKill As Long
    NpcDead As Long
    EnviroDead As Long

End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    Target As Long
    tType As Byte
End Type

Public Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Long
    caster As Long
    StartTime As Long
End Type

Public Type StatBufferRec
    Value As Integer
    Timer As Long
End Type

Public Type SwitchRec
    Value As Boolean
    Timer As Long
End Type



Public Type TempPlayerRec
    ' Non saved local vars
    Buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    TargetType As Byte
    Target As Long
    GettingMap As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    InBank As Boolean
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerInvRec
    AcceptTrade As Boolean
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell buffer
    spellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' party
    inParty As Long
    partyInvite As Long
    
    ProjecTile(1 To MAX_PLAYER_PROJECTILES) As ProjectileRec
    FreeAction As Boolean
    tmpGuildSlot As Long
    tmpGuildInviteSlot As Long
    tmpGuildInviteTimer As Long
    tmpGuildInviteId As Long
    IsLoading As Boolean
    'Pet
    TempPet As TempPlayerPetRec
    '/Pet
    'Safe Mode
    SafeMode As Byte
    
    'Ping
    Req As Boolean
    PingStart As Long
    Ping As Long
    
    'Inv Current Weight
    weight As Long
    
    'Team index
    TeamIndex As Byte
    
    InactiveTime As Long
    
    Stats(1 To Stats.Stat_Count - 1) As Integer
    StatsBuffer(1 To Stats.Stat_Count - 1) As StatBufferRec
    
    BlockedActions(1 To PlayerActionsType.PlayerActions_Count - 1) As SwitchRec
    ProtectedActions(1 To PlayerActionsType.PlayerActions_Count - 1) As SwitchRec
    BlockedDirections(DIR_UP To DIR_RIGHT) As Boolean
    
    SentMsg As Integer
    
    SpeedHackChecker As Long
    RunSpeed As Long
    WalkSpeed As Long
    
    MovementsStack As clsStack
    
    LastSpell As Long
    
End Type

Private Type TileDataRec
    X As Long
    Y As Long
    Tileset As Long
End Type

Private Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    DirBlock As Byte
End Type

Public Type SingularMovementRec
    'Number of tiles is a property of custom by tile movement, won't be used in other movement types
    Direction As Byte
    NumberOfTiles As Byte
End Type

Public Type VectDataRec
    Data As SingularMovementRec
End Type

Public Type MovementsListRec
    Actual As Byte
    nelem As Byte
    vect() As VectDataRec
End Type

Public Type MovementRec
    'Basic information about the movement, name, his type and table containing movements
    Name As String * NAME_LENGTH
    Type As MovementType
    'Table size: Only Directional: 1 value, Custom: multiple value
    MovementsTable As MovementsListRec
    Repeat As Boolean ' if true, movement starts at the beginning of the list when it ends
End Type

Public Type MapNPCPropertiesRec
    'contains index's of the Movements/Actions table
    Movement As Byte
    Action As Byte
End Type

Public Type MapRec
    Name As String * NAME_LENGTH
    
    Music As String * NAME_LENGTH
    
    Revision As Long
    moral As Byte
    
    Up As Long
    Down As Long
    left As Long
    right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    NPC(1 To MAX_MAP_NPCS) As Long
    
    Weather As Long ' 0 = None, 1 = Rain 2, = Snow, 3 = Sandstorm
    
    NPCSProperties(1 To MAX_MAP_NPCS) As MapNPCPropertiesRec
    
    AllowedStates(1 To Max_States - 1) As Boolean
    
    TranslatedName As String * NAME_LENGTH
End Type

Public Type ServerTileRec
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    DirBlock As Byte
End Type

Public Type ServerMapRec
    Name As String * NAME_LENGTH
    
    Revision As Long
    moral As Byte
    
    Up As Long
    Down As Long
    left As Long
    right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As ServerTileRec
    NPC(1 To MAX_MAP_NPCS) As Long
    
    Weather As Long ' 0 = None, 1 = Rain 2, = Snow, 3 = Sandstorm
    
    NPCSProperties(1 To MAX_MAP_NPCS) As MapNPCPropertiesRec
    
    AllowedStates(1 To Max_States - 1) As Boolean
    
    TranslatedName As String * NAME_LENGTH
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    
    stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    Face As Long
    
    startItemCount As Long
    StartItem() As Long
    StartValue() As Long
    
    startSpellCount As Long
    StartSpell() As Long
    StartMap As Long
    StartMapX As Long
    StartMapY As Long
    
    TranslatedName As String * NAME_LENGTH
End Type

Private Type ContainerRec
    ItemNum As Long
    Value As Long
End Type

Private Type ImpactarRec
    Spell As Long
    Auto As Long
End Type




Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    Speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Long
    
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    ProjecTile As ProjectileRec
    ammo As Long
    ammoreq As Long
    ConsumeItem As Long
    istwohander As Boolean
    Container(MAX_ITEM_CONTAINERS) As ContainerRec
    AddBags As Byte
    weight As Long

    AddHPPercent As Boolean
    AddMPPercent As Boolean

    ArmyType_Req As Byte
    ArmyRange_Req As Byte
    
    Impactar As ImpactarRec
    ExtraHP As Long
    
    TranslatedName As String * NAME_LENGTH
    
End Type

Private Type MapItemRec
    Num As Long
    Value As Long
    X As Byte
    Y As Byte
    ' ownership + despawn
    playerName As String
    playerTimer As Long
    isDrop As Boolean
    Timer As Long
End Type

' Type for npc's drops
Private Type NpcDropRec
    DropChance As Long
    DropItem As Long
    DropItemValue As Long
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    range As Byte
    'DropChance As Long
    'DropItem As Long
    'DropItemValue As Long
    stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    exp As Long
    Animation As Long
    Damage As Long
    level As Long
    ' Npc Spells
    Spell(1 To MAX_NPC_SPELLS) As Long
    'ALATAR
    Quest As Byte
    questnum As Long
    '/ALATAR
    'Need npc converter in order to start multiple drops
    Drops(1 To MAX_NPC_DROPS) As NpcDropRec
    Speed As Long
    
    TranslatedName As String * NAME_LENGTH
End Type

Public Type MapNpcRec
    Num As Long
    Target As Long
    TargetType As Byte
    vital(1 To Vitals.Vital_Count - 1) As Long
    X As Byte
    Y As Byte
    dir As Byte
    mapnpcnum As Long 'references original map npc num
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' Npc spells
    SpellTimer(1 To MAX_NPC_SPELLS) As Long
    'Npc Movement List
    Inverse As Boolean
    Count As Byte
    Actual As Byte
    'Pet Data
    'IsPet As Byte
    PetData As MapPetRec
    'Temp npc
    IsTempNPC As Boolean
    'Walk timer
    MoveTimer As Long
    LastDir As Byte
End Type



Private Type SpellRec
    Name As String * NAME_LENGTH
    
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    map As Long
    X As Long
    Y As Long
    dir As Byte
    vital As Long
    Duration As Long
    Interval As Long
    range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    stat As Byte
    BlockActions(1 To PlayerActions_Count - 1) As Boolean
    UsePercent As Boolean
    StatDamage As Byte
    StatDefense As Byte
    ChangeState As Byte
    
    TranslatedName As String * NAME_LENGTH
End Type

Private Type TileDoorRec
    doornum As Long
    DoorTimer As Long
    X As Long
    Y As Long
    state As Boolean
End Type

Private Type TempTileRec
    Door() As TileDoorRec
    NumDoors As Integer
    NPCSpawnSite() As Point
    NumSpawnSites As Integer
End Type



Private Type MapDataRec
    NPC() As MapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    X As Long
    Y As Long
    cur_health As Long
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
End Type

Private Type ResourceRewardRec
    Reward As Long 'used for encoding option
    Chance As Byte
    RewardType As Byte
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    ResourceType As Integer
    ResourceImage As Long
    ExhaustedImage As Long
    ToolRequired As Long
    health As Long
    RespawnTime As Long
    Walkthrough As Boolean
    Animation As Long
    
    WalkableNormal As Boolean
    WalkableExhausted As Boolean
    
    Rewards(1 To MAX_RESOURCE_REWARDS) As ResourceRewardRec
    
    ' True = say item name when reward, False = Say specificated caption
    ItemSuccessMessage As Boolean
    
    TranslatedName As String * NAME_LENGTH
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
    
    TranslatedName As String * NAME_LENGTH
End Type

'Used for calculate Resource's rewards probability
Public Type AuxiliarResourceRewardrec
    CummulativeProb As Byte
    index As Byte
End Type

Public Type ActionRec
    Name As String * NAME_LENGTH
    
    Type As Byte
    Moment As MomentType
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    range As Byte
    
    TranslatedName As String * NAME_LENGTH
End Type

Public Type WaitingItemsRec
    X As Long
    Y As Long
    Active As Boolean
    Timer As Long
End Type




