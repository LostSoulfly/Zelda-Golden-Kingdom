Attribute VB_Name = "modConstants"
Option Explicit

' API Declares
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByRef msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "USER32" () As Long

' animated buttons
Public Const MAX_MENUBUTTONS As Long = 4
Public Const MAX_MAINBUTTONS As Long = 9 'Alatar v1.2 replace 6 to 7
Public Const MENUBUTTON_PATH As String = "\data\graphics\gui\menu\buttons\"
Public Const MAINBUTTON_PATH As String = "\data\graphics\gui\main\buttons\"

'Currency Name
Public Const CURRENCY_NAME As String = "Tingles"

' Hotbar
Public Const HotbarTop As Long = 2
Public Const HotbarLeft As Long = 2
Public Const HotbarOffsetX As Long = 8

' Inventory constants
Public Const InvTop As Long = 24
Public Const InvLeft As Long = 12
Public Const InvOffsetY As Long = 3
Public Const InvOffsetX As Long = 3
Public Const InvColumns As Long = 5

' Bank constants
Public Const BankTop As Long = 38
Public Const BankLeft As Long = 42
Public Const BankOffsetY As Long = 3
Public Const BankOffsetX As Long = 4
Public Const BankColumns As Long = 11

' spells constants
Public Const SpellTop As Long = 24
Public Const SpellLeft As Long = 12
Public Const SpellOffsetY As Long = 3
Public Const SpellOffsetX As Long = 3
Public Const SpellColumns As Long = 5

' shop constants
Public Const ShopTop As Long = 6
Public Const ShopLeft As Long = 8
Public Const ShopOffsetY As Long = 2
Public Const ShopOffsetX As Long = 4
Public Const ShopColumns As Long = 5

' Character consts
Public Const EqTop As Long = 224
Public Const EqLeft As Long = 18
Public Const EqOffsetX As Long = 10
Public Const EqColumns As Long = 4

' values
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' path constants
Public Const SOUND_PATH As String = "\data\sound\"
Public Const MUSIC_PATH As String = "\data\music\"

' Font variables
Public Const FONT_NAME As String = "Georgia"
Public Const FONT_SIZE As Byte = 14

' Log Path and variables
Public Const LOG_DEBUG As String = "debug.txt"
Public Const LOG_PATH As String = "\data\logs\"

' Map Path and variables
Public Const MAP_PATH As String = "\data\maps\"
Public Const MAP_EXT As String = ".map"

' Gfx Path and variables
Public Const GFX_PATH As String = "\data\graphics\"
Public Const GFX_EXT As String = ".bmp"

' Key constants
Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11

' Menu states
Public Const MENU_STATE_NEWACCOUNT As Byte = 0
Public Const MENU_STATE_DELACCOUNT As Byte = 1
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_GETCHARS As Byte = 3
Public Const MENU_STATE_NEWCHAR As Byte = 4
Public Const MENU_STATE_ADDCHAR As Byte = 5
Public Const MENU_STATE_DELCHAR As Byte = 6
Public Const MENU_STATE_USECHAR As Byte = 7
Public Const MENU_STATE_INIT As Byte = 8

' Speed moving vars
Public Const WALK_SPEED As Byte = 8
Public Const RUN_SPEED As Byte = 12

' Tile size constants
Public Const PIC_X As Long = 32
Public Const PIC_Y As Long = 32

' Sprite, item, spell size constants
Public Const SIZE_X As Long = 32
Public Const SIZE_Y As Long = 32

' ********************************************************
' * The values below must match with the server's values *
' ********************************************************

' General constants
Public Const MAX_DOORS As Long = 400
Public Const MAX_PLAYERS As Long = 70
Public Const MAX_ITEMS As Long = 900
Public Const MAX_NPCS As Long = 600
Public Const MAX_ANIMATIONS As Long = 255
Public Const MAX_NPC_SPELLS As Long = 10
Public Const MAX_INV As Long = 35
Public Const MAX_MAP_ITEMS As Long = 255
Public Const MAX_MAP_NPCS As Long = 30
Public Const MAX_SHOPS As Long = 100
Public Const MAX_PLAYER_SPELLS As Long = 35
Public Const MAX_SPELLS As Long = 255
Public Const MAX_TRADES As Long = 30
Public Const MAX_RESOURCES As Long = 200
Public Const MAX_LEVELS As Long = 80
Public Const MAX_BANK As Long = 99
Public Const MAX_HOTBAR As Long = 12
Public Const MAX_PARTYS As Long = 35
Public Const MAX_PARTY_MEMBERS As Long = 4
Public Const MAX_PLAYER_PROJECTILES As Long = 5
Public Const MAX_NPC_DROPS As Long = 5
Public Const MAX_RESOURCE_REWARDS As Byte = 5
Public Const MAX_MOVEMENTS As Byte = 100
Public Const MAX_MOVEMENT_MOVEMENTS As Byte = 255
Public Const MAX_ACTIONS As Byte = 100
Public Const MAX_PETS As Byte = 50
Public Const MAX_PLAYER_PETS As Byte = 6
Public Const MAX_ITEM_CONTAINERS As Byte = 5
'Pet System
Public Const MAX_PET_POINTS_PERLVL As Byte = 10

' Website
Public Const GAME_WEBSITE As String = "http://trollparty.org/Zelda/"

' text color constants
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = BrightBlue
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = BrightGreen
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = BrightBlue
Public Const WhoColor As Byte = BrightBlue
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = BrightRed
Public Const NewMapColor As Byte = BrightBlue
Public Const CourageColor As Byte = BrightBlue
Public Const WisdomColor As Byte = BrightGreen
Public Const PowerColor As Byte = Red

Public Const PKColor As Long = BrightRed
Public Const HeroColor As Long = 65535
Public Const NoneColor As Long = 43775

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const ACCOUNT_LENGTH As Byte = 12

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
Public Const MAX_MAPS As Long = 900
Public Const MAX_MAPX As Byte = 24
Public Const MAX_MAPY As Byte = 18
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_ARENA As Byte = 2
Public Const MAP_MORAL_PK_SAFE As Byte = 3
Public Const MAP_MORAL_PACIFIC As Byte = 4

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_RESOURCE As Byte = 7
Public Const TILE_TYPE_DOOR As Byte = 8
Public Const TILE_TYPE_NPCSPAWN As Byte = 9
Public Const TILE_TYPE_SHOP As Byte = 10
Public Const TILE_TYPE_BANK As Byte = 11
Public Const TILE_TYPE_HEAL As Byte = 12
Public Const TILE_TYPE_TRAP As Byte = 13
Public Const TILE_TYPE_SLIDE As Byte = 14
Public Const TILE_TYPE_SCRIPT As Byte = 15
Public Const TILE_TYPE_ICE As Byte = 16


'Rupees Value
Public Const GREEN_RUPEE As Long = 1
Public Const BLUE_RUPEE As Long = 5
Public Const YELLOW_RUPEE As Long = 10
Public Const RED_RUPEE As Long = 20
Public Const PURPLE_RUPEE As Long = 50

'Rupees Capacity Constants
Public Const MAX_BANK_RUPEES As Long = 9999
Public Const MAX_PLAYER_INITIAL_RUPEES As Long = 100
Public Const MAX_RUPEE_BAGS As Byte = 10
Public Const BAG_CAPACITY As Long = 100


' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' Constants for player movement: Tiles per Second
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Admin constants
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_DEVELOPER As Byte = 2
Public Const ADMIN_MAPPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOUR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOUR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOUR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOUR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOUR_GUARD As Byte = 4
Public Const NPC_BEHAVIOUR_BLADE As Byte = 5
Public Const NPC_BEHAVIOUR_SLIDE As Byte = 6

' Spell constants
Public Const SPELL_TYPE_DAMAGEHP As Byte = 0
Public Const SPELL_TYPE_DAMAGEMP As Byte = 1
Public Const SPELL_TYPE_HEALHP As Byte = 2
Public Const SPELL_TYPE_HEALMP As Byte = 3
Public Const SPELL_TYPE_WARP As Byte = 4
Public Const SPELL_TYPE_BUFFER As Byte = 5
Public Const SPELL_TYPE_PROTECT As Byte = 6
Public Const SPELL_TYPE_CHANGESTATE As Byte = 7


' Game editor constants
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4
Public Const EDITOR_RESOURCE As Byte = 5
Public Const EDITOR_ANIMATION As Byte = 6
Public Const EDITOR_DOORS As Byte = 7
Public Const EDITOR_MOVEMENTS As Byte = 8
Public Const EDITOR_ACTIONS As Byte = 9
Public Const EDITOR_PETS As Byte = 10
Public Const EDITOR_CUSTOMSPRITES As Byte = 10

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

' Dialogue box constants
Public Const DIALOGUE_TYPE_NONE As Byte = 0
Public Const DIALOGUE_TYPE_TRADE As Byte = 1
Public Const DIALOGUE_TYPE_FORGET As Byte = 2
Public Const DIALOGUE_TYPE_PARTY As Byte = 3
Public Const DIALOGUE_TYPE_QUESTION As Byte = 4

'Action Constants
Public Const ACTION_TYPE_SUBVITAL As Byte = 0
Public Const ACTION_TYPE_WARP As Byte = 1

' Do Events
Public Const nLng As Long = (&H80 Or &H1 Or &H4 Or &H20) + (&H8 Or &H40)

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2

'item weight constants
Public Const ONE_VITAL_WEIGHT As Byte = 1   'unity
Public Const ONE_STAT_ADD_WEIGHT As Long = 50

'Resource Reward Enum
Public Const REWARD_ITEM As Byte = 1
Public Const REWARD_SPAWN_NPC As Byte = 2

'Max Chat Lines
Public Const MAX_CHAT_LINES As Byte = 200

' stuffs
Public Const HalfX As Integer = ((MAX_MAPX + 1) / 2) * PIC_X
Public Const HalfY As Integer = ((MAX_MAPY + 1) / 2) * PIC_Y
Public Const ScreenX As Integer = (MAX_MAPX + 1) * PIC_X
Public Const ScreenY As Integer = (MAX_MAPY + 1) * PIC_Y
Public Const StartXValue As Integer = ((MAX_MAPX + 1) / 2)
Public Const StartYValue As Integer = ((MAX_MAPY + 1) / 2)
Public Const EndXValue As Integer = (MAX_MAPX + 1) + 1
Public Const EndYValue As Integer = (MAX_MAPY + 1) + 1
Public Const Half_PIC_X As Integer = PIC_X / 2
Public Const Half_PIC_Y As Integer = PIC_Y / 2

Public Const LEVEL_SOUND As String = "LevelUp.wav"
Public Const Switch As String = "Switch.wav"
Public Const SwitchFloor As String = "SwitchFloor.wav"
Public Const Sandstorm As String = "Sandstorm.wav"
Public Const Slide As String = "Slide.wav"
Public Const Reset As String = "Error.wav"
Public Const Error As String = "Error.wav"

Public Const ChatBubbleWidth As Long = 200
Public Const Font_Default As String = "Default"
