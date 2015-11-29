Attribute VB_Name = "modConstants"
Option Explicit

' API
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' path constants
Public Const ADMIN_LOG As String = "admin.log"
Public Const PLAYER_LOG As String = "player.log"
Public Const GUILD_LOG As String = "guild.log"

' Version constants
Public Const CLIENT_MAJOR As Byte = 0
Public Const CLIENT_MINOR As Byte = 5
Public Const CLIENT_REVISION As Byte = 20
Public Const MAX_LINES As Long = 500 ' Used for frmServer.txtText

' ********************************************************
' * The values below must match with the client's values *
' ********************************************************
' General constants
Public Const MAX_DOORS As Long = 400
Public Const MAX_PLAYERS As Long = 40
Public Const MAX_ITEMS As Long = 900
Public Const MAX_NPCS As Long = 600
Public Const MAX_ANIMATIONS As Long = 255
Public Const MAX_INV As Long = 35
Public Const MAX_MAP_ITEMS As Long = 30
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
Public Const MAX_PLAYER_PROJECTILES As Long = 5
Public Const MAX_PARTY_MEMBERS As Long = 4
Public Const MAX_NPC_DROPS As Long = 5
Public Const MAX_RESOURCE_REWARDS As Byte = 5
Public Const MAX_MOVEMENTS As Byte = 100
Public Const MAX_MOVEMENT_MOVEMENTS As Byte = 255
Public Const MAX_ACTIONS As Byte = 100
Public Const MAX_PETS As Byte = 100
Public Const MAX_PLAYER_PETS As Byte = 6
Public Const MAX_ITEM_CONTAINERS As Byte = 5
Public Const MAX_SPELL_EFECTS As Byte = 3

'pet system
Public Const MAX_PET_POINTS_PERLVL As Byte = 6

Public Const POINTS_PERLVL As Byte = 3
Public Const MAX_STAT As Byte = 130
Public Const MAX_PET_STAT As Byte = 100


' NPC Spells
Public Const MAX_NPC_SPELLS As Long = 10

' server-side stuff
Public Const ITEM_SPAWN_TIME As Long = 1000 ' 1 seconds
Public Const ITEM_DESPAWN_TIME As Long = 120000 ' 2 minutes
Public Const MAX_DOTS As Long = 30

'Minims
Public Const MIN_LEVEL_TO_RESET As Long = 80

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
Public Const WhoColor As Byte = Grey
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = BrightRed
Public Const NewMapColor As Byte = BrightBlue
Public Const PKColor As Long = BrightRed
Public Const HeroColor As Long = BrightBlue

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
Public Const MAX_MAPX As Byte = 25
Public Const MAX_MAPY As Byte = 19
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_ARENA As Byte = 2
Public Const MAP_MORAL_PK_SAFE As Byte = 3
Public Const MAP_MORAL_PACIFIC As Byte = 4

' Tile consants
'Public Const TILE_TYPE_WALKABLE As Byte = 0
'Public Const TILE_TYPE_BLOCKED As Byte = 1
'Public Const TILE_TYPE_WARP As Byte = 2
'Public Const TILE_TYPE_ITEM As Byte = 3
'Public Const TILE_TYPE_NPCAVOID As Byte = 4
'Public Const TILE_TYPE_KEY As Byte = 5
'Public Const TILE_TYPE_KEYOPEN As Byte = 6
'Public Const TILE_TYPE_RESOURCE As Byte = 7
'Public Const TILE_TYPE_DOOR As Byte = 8
'Public Const TILE_TYPE_NPCSPAWN As Byte = 9
'Public Const TILE_TYPE_SHOP As Byte = 10
'Public Const TILE_TYPE_BANK As Byte = 11
'Public Const TILE_TYPE_HEAL As Byte = 12
'Public Const TILE_TYPE_TRAP As Byte = 13
'Public Const TILE_TYPE_SLIDE As Byte = 14
'Public Const TILE_TYPE_SCRIPT As Byte = 15
'Public Const TILE_TYPE_ICE As Byte = 16

' Item constants


' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3



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

'Door Types
Public Const DOOR_TYPE_DOOR As Byte = 0
Public Const DOOR_TYPE_SWITCH As Byte = 1
Public Const DOOR_TYPE_WEIGHTSWITCH As Byte = 2

' Game editor constants
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4
Public Const EDITOR_RESOURCE As Byte = 5
Public Const EDITOR_ANIMATION As Byte = 6
Public Const EDITOR_DOORS As Byte = 7

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

'Action Constants
Public Const ACTION_TYPE_SUBVITAL As Byte = 0
Public Const ACTION_TYPE_WARP As Byte = 1

'Justice Status Constants
Public Const NONE_PLAYER As Byte = 0
Public Const PK_PLAYER As Byte = 1
Public Const HERO_PLAYER As Byte = 2

'Rupees Value
Public Const GREEN_RUPEE As Long = 1
Public Const BLUE_RUPEE As Long = 5
Public Const YELLOW_RUPEE As Long = 10
Public Const RED_RUPEE As Long = 20
Public Const PURPLE_RUPEE As Long = 50

'Rupees Capacity Constants
Public Const MAX_BANK_RUPEES As Long = 9999
Public Const MAX_RUPEE_BAGS As Byte = 10
Public Const BAG_CAPACITY As Long = 100
Public Const INITIAL_BAGS As Byte = 1

'Temporal Constants for Random Spawning system
Public Const NPC_SKULLTULA As Long = 442
'Public Const SKULLTULAS As Long = 1
Public Const SKULLTULAS As Long = 100

'Resource Reward Enum
Public Const REWARD_ITEM As Byte = 1
Public Const REWARD_SPAWN_NPC As Byte = 2

' ********************************************
' Default starting location [Server Only]
Public Const START_MAP As Long = 1
Public Const START_X As Long = 11
Public Const START_Y As Long = 3

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2

'Sleep time constants
Public Const ONE_PLAYER_WAIT_TIME As Long = 10
Public Const NO_PLAYERS_WAIT_TIME As Long = 25

' values
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647



'Public Const EULER As Double = 2.71828182845905

' Do Events
Public Const nLng As Long = (&H80 Or &H1 Or &H4 Or &H20) + (&H8 Or &H40)
