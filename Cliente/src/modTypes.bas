Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Doors(1 To MAX_DOORS) As DoorRec
Public map As MapRec
Public Bank As BankRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Movements(1 To MAX_MOVEMENTS) As MovementRec
Public Actions(1 To MAX_ACTIONS) As ActionRec
Public Pet(1 To MAX_PETS) As PetRec
Public Chat(1 To 20) As ChatRec

' client-side stuff
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public MenuButton(1 To MAX_MENUBUTTONS) As ButtonRec
Public MainButton(1 To MAX_MAINBUTTONS) As ButtonRec
Public Party As PartyRec
Public RainDrop(0 To 200) As RainDropRec
Public ChatRooms(1 To 6) As list

' options
Public Options As OptionsRec

' Type recs
Private Type OptionsRec
    Game_Name As String
    SavePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    ip As String
    port As Long
    MenuMusic As String
    Music As Byte
    Sound As Byte
    Debug As Byte
    Names As Byte
    Level As Byte
    WASD As Byte
    Chat As Byte
    SafeMode As Byte
    DefaultVolume As Byte
    ActivatedChats(1 To 6) As Boolean
    MiniMap As Byte
    MappingMode As Byte
    ChatToScreen As Byte
    DllRegistered As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    num As Long
    value As Long
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type SpellAnim
    spellnum As Long
    Timer As Long
    FramePointer As Long
End Type

Public Type ProjectileRec
    TravelTime As Long
    direction As Long
    X As Long
    y As Long
    Pic As Long
    range As Long
    Damage As Long
    Speed As Long
    Depth As Byte
End Type

Public Type DoorRec
    Name As String * NAME_LENGTH
    
    DoorType As Long
    
    WarpMap As Long
    WarpX As Long
    WarpY As Long
    
    UnlockType As Long
    key As Long
    Switch As Long
    
    Time As Long
    
    InitialState As Boolean
    
    TranslatedName As String * NAME_LENGTH
End Type

Public Type PlayerPetRec
    'Link to super class
    NumPet As Byte
    'stadistics
    StatsAdd(1 To Stats.Stat_Count - 1) As Byte
    points As Integer
    Experience As Long
    Level As Long
    CurrentHP As Long
End Type

Public Enum PetState
    Passive = 0
    Assist = 1
    Defensive = 2
End Enum

Public Type MapPetRec
    Owner As Long
    Name As String * NAME_LENGTH
End Type

Private Type PlayerRec
    ' General
    Name As String
    Class As Long
    sprite As Long
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    ' Vitals
    vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    stat(1 To Stats.Stat_Count - 1) As Byte
    points As Long
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    ' Position
    map As Long
    X As Byte
    y As Byte
    dir As Byte
    'Pet As PlayerPetRec
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    step As Byte
    GuildName As String
    GuildMemberId As Long
    onIce As Boolean
    IceDir As Byte
    ' projectiles
    ProjecTile(1 To MAX_PLAYER_PROJECTILES) As ProjectileRec
    Visible As Long
    'ALATAR
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec
    '/ALATAR
    PlayerDoors(1 To MAX_DOORS) As DoorRec
    'Pet system
    Pet(1 To MAX_PLAYER_PETS) As PlayerPetRec
    ActualPet As Byte
    PetState As Byte
    'Triforce
    triforce(1 To TriforceType.TriforceType_Count - 1) As Boolean
    
    'Rupee System
    RupeeBags As Byte
    
    'Sprite System
    CustomSprite As Byte '0: normal sprite, > 0: custom sprite
    
    'weight system
    MaxWeight As Long
    
    BonusPoints As Long
    
    WalkSpeed As Long
    RunSpeed As Long
    
    MovementSprite As Boolean
    PreviousSprite As Long
    
    State As Byte
    
    RideInfo As PlayerRideRec
    
    LagDirections As clsQueue
    LagMovements As clsQueue
    automatizedmove As Boolean
    Started As Boolean
    'Kill Counter
    Kill As Long
    Dead As Long
    NpcKill As Long
    NpcDead As Long
    EnviroDead As Long
End Type

Public Type TileDataRec
    X As Long
    y As Long
    Tileset As Long
End Type

Public Type TileRec
    layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    DirBlock As Byte
End Type

Public Type SingularMovementRec
    'Number of tiles is a property of custom by tile movement, won't be used in other movement types
    direction As Byte
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
    movement As Byte
    Action As Byte
    Inverse As Boolean
    Count As Byte
    Actual As Byte
End Type

Public Type MapRec
    Name As String * NAME_LENGTH
    
    Music As String * NAME_LENGTH
    
    Revision As Long
    moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    NPC(1 To MAX_MAP_NPCS) As Long
    
    Weather As Long ' 0 = None 1 = Rain, 2 = Snow, 3 = Sandstorm
    
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
    ' For client use
    vital(1 To Vitals.Vital_Count - 1) As Long
    
    TranslatedName As String * NAME_LENGTH
End Type

Private Type ContainerRec
    ItemNum As Long
    value As Long
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
    Weight As Long
    
    AddHPPercent As Boolean
    AddMPPercent As Boolean
    
    ArmyType_Req As Byte
    ArmyRange_Req As Byte
    
    Impactar As ImpactarRec
    ExtraHP As Long
    
    TranslatedName As String * NAME_LENGTH
End Type

Private Type MapItemRec
    PlayerName As String
    num As Long
    value As Long
    Frame As Byte
    X As Byte
    y As Byte
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
    
    sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    range As Byte
    stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    Exp As Long
    Animation As Long
    Damage As Long
    Level As Long
    ' Npc Spells
    Spell(1 To MAX_NPC_SPELLS) As Long
    'ALATAR
    Quest As Byte
    QuestNum As Long
    '/ALATAR
    Drops(1 To MAX_NPC_DROPS) As NpcDropRec
    Speed As Long
    
    TranslatedName As String * NAME_LENGTH
End Type



Private Type MapNpcRec
    num As Long
    target As Long
    targetType As Byte
    vital(1 To Vitals.Vital_Count - 1) As Long
    map As Long
    X As Byte
    y As Byte
    dir As Byte
    ' Client use only
    XOffset As Long
    YOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    step As Byte
    'Pet Data
    'IsPet As Byte
    petData As MapPetRec

End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
    PriceType As Byte 'references shop prices type
    
    TranslatedName As String * NAME_LENGTH
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
    y As Long
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

Private Type TempTileRec
    DoorOpen As Byte
    DoorFrame As Byte
    DoorTimer As Long
    DoorAnimate As Byte ' 0 = nothing| 1 = opening | 2 = closing
End Type

Public Type MapResourceRec
    X As Long
    y As Long
    ResourceState As Byte
End Type

Private Type ResourceRewardRec
    Reward As Long
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

Private Type ActionMsgRec
    message As String
    Created As Long
    Type As Long
    Color As Long
    Scroll As Long
    X As Long
    y As Long
    Timer As Long
End Type

Private Type BloodRec
    sprite As Long
    Timer As Long
    X As Long
    y As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    
    Sound As String * NAME_LENGTH
    
    sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
    
    TranslatedName As String * NAME_LENGTH
End Type

Private Type AnimInstanceRec
    Animation As Long
    X As Long
    y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    ' timing
    Timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    FrameIndex(0 To 1) As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type ButtonRec
    Filename As String
    State As Byte
End Type

Public Type ChatBubbleRec
    msg As String
    colour As Long
    target As Long
    targetType As Byte
    Timer As Long
    active As Boolean
End Type

Public Type RainDropRec
X As Long
y As Long
InMotion As Long
End Type

Private Type TempPlayerRec
    IsLoading As Boolean
End Type



'Used for drops information on editor
Public Type NPCDropInfoRec
    Number As Long
    Chances As Long
    value As Long
End Type

Public Type ResourceRewardInfoRec
    Reward As Long
    Chance As Byte
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

Public Type PetRec
    Name As String * NAME_LENGTH
    NPCNum As Long
    'Requeriments
    TamePoints As Integer
    'Progressions
    ExpProgression As Byte
    PointsProgression As Byte
    MaxLevel As Long
End Type




Public Type ServerPlayerRec
    ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    sprite As Long
    Level As Byte
    Exp As Long
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
    y As Byte
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
    
    State As Byte
End Type


Public Type ServerBankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type


Private Type ServerGuildRanksRec
    'General variables
    Used As Boolean
    Name As String
    
    'Rank Variables
    RankPermission(1 To MAX_GUILD_RANKS_PERMISSION) As Byte
End Type

Private Type ServerGuildMemberRec
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

Public Type ServerGuildRec
    In_Use As Boolean
    
    Guild_Name As String
    
    'Guild file number for saving
    Guild_Fileid As Long
    
    Guild_Members(1 To MAX_GUILD_MEMBERS) As ServerGuildMemberRec
    Guild_Ranks(1 To MAX_GUILD_RANKS) As ServerGuildRanksRec
    
    'Message of the day
    Guild_MOTD As String * 100
    
    'The rank recruits start at
    Guild_RecruitRank As Integer
    'Color of guild name
    Guild_Color As Integer

End Type

Public Type SwitchRec
    value As Boolean
    Timer As Long
End Type

Private Type ChatRec
    Text As String
    colour As Long
End Type

