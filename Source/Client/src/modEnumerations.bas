Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the server's packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SFullMsg
    SSpeechWindow
    SLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
    SPlayerStats
    SPlayerStat
    SPlayerData
    SPlayerMove
    SNpcMove
    SPlayerDir
    SNpcDir
    SPlayerXY
    SPlayerXYMap
    SAttack
    SNpcAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SUpdateItems
    SREditor
    SSpawnNpc
    SNpcDead
    SNpcEditor
    SUpdateNpc
    SMapKey
    SEditMap
    SShopEditor
    SUpdateShop
    SSpellEditor
    SUpdateSpell
    SSpells
    SLeft
    SResourceCache
    SResourceEditor
    SUpdateResource
    SSendPing
    SDoorAnimation
    SActionMsg
    SPlayerEXP
    SBlood
    SAnimationEditor
    SUpdateAnimation
    SAnimation
    SMapNpcVitals
    SCooldown
    SClearSpellBuffer
    SSayMsg
    SOpenShop
    SResetShopAction
    SStunned
    SMapWornEq
    SBank
    STrade
    SCloseTrade
    STradeUpdate
    STradeStatus
    STarget
    SHotbar
    SHighIndex
    SSound
    STradeRequest
    SPartyInvite
    SPartyUpdate
    SPartyVitals
    SSendGuild
    SAdminGuild
    SGuildAdminSwitchTab
    SHandleProjectile

    'ALATAR
    SQuestEditor
    SUpdateQuest
    SPlayerQuest
    SQuestMessage
    '/ALATAR

    SSpawnEvent
    SEventMove
    SEventDir
    SEventChat
    SEventStart
    SEventEnd
    SPlayBGM
    SPlaySound
    SFadeoutBGM
    SStopSound
    SSwitchesAndVariables
    SMapEventData
    SDoorsEditor
    SUpdateDoors
    SChatBubble
    SLoad
    SDone
    SSendWeather
    SMovementsEditor
    SUpdateMovements
    SActionsEditor
    SUpdateActions
    SNPCCache
    SPetsEditor
    SUpdatePets
    SPetData
    SOpenTriforce
    SOnIce
    SIceDir
    SBags
    SPoints
    SLevel
    SJustice
    SPlayerAttack
    SMapSingularNpcData
    SAccounts
    SCustomSpritesEditor
    SUpdateCustomSprites
    SPlayerSprite
    SSingleResourceCache
    SGuildData
    SMaxWeight
    SMapSingularItemData
    SBanks
    SGuilds
    SQuestion
    SKillPoints
    SBonusPoints
    SUpdateNPCS
    SSpeedReq
    SPlayerSpeed
    SRunningSprites
    SPlayerState
    SUpdate
    SStaminaInfo
    SCharList
    SSaveFiles
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CDelAccount
    CLogin
    CAddChar
    CUseChar
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
    CPlayerMsg
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CPlayerInfoRequest
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CGetStats
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanList
    CBanDestroy
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CSaveItem
    CRequestEditNpc
    CSaveNpc
    CRequestEditShop
    CSaveShop
    CRequestEditSpell
    CSaveSpell
    CSetAccess
    CSetName
    CWhosOnline
    CSetMotd
    CSearch
    CSpells
    CCast
    CQuit
    CSwapInvSlots
    CRequestEditResource
    CSaveResource
    CCheckPing
    CUnequip
    CRequestPlayerData
    CRequestItems
    CRequestNPCS
    CRequestResources
    CSpawnItem
    CRequestEditAnimation
    CSaveAnimation
    CRequestAnimations
    CRequestSpells
    CRequestShops
    CRequestLevelUp
    CForgetSpell
    CCloseShop
    CBuyItem
    CSellItem
    CChangeBankSlots
    CDepositItem
    CWithdrawItem
    CCloseBank
    CAdminWarp
    CTradeRequest
    CAcceptTrade
    CDeclineTrade
    CTradeItem
    CUntradeItem
    CHotbarChange
    CHotbarUse
    CSwapSpellSlots
    CAcceptTradeRequest
    CDeclineTradeRequest
    CPartyRequest
    CAcceptParty
    CDeclineParty
    CPartyLeave
    CGuildCommand
    CSayGuild
    CSaveGuild
    CRequestGuildAdminTabSwitch
    CProjecTileAttack
    'ALATAR
    CRequestEditQuest
    CSaveQuest
    CRequestQuests
    CPlayerHandleQuest
    CQuestLogUpdate
    '/ALATAR
    CEventChatReply
    CEvent
    CSwitchesAndVariables
    CRequestSwitchesAndVariables
    CSaveDoor
    CRequestDoors
    CRequestEditDoors
    CSaveMovement
    CRequestMovements
    CRequestEditMovements
    CSaveAction
    CRequestActions
    CRequestEditActions
    CPartyChatMsg
    CPlayerVisibility
    CDone
    CSpawnPet
    CPetFollowOwner
    CPetAttackTarget
    CPetWander
    CPetDisband
    CSavePet
    CRequestPets
    CRequestEditPets
    CRequestTame
    CRequestChangePet
    CPetData
    CUsePetStatPoint
    CPetForsake
    CPetPercentChange
    CResetPlayer
    CSafeMode
    COnIce
    CAck
    CAttackNPC
    CCheckItems
    CNeedAccounts
    CSaveCustomSprite
    CRequestCustomSprites
    CRequestEditCustomSprites
    CCheckResource
    CMute
    CShutdown
    CRestart
    CMakeAdmin
    CAddException
    CSpecialBan
    CAnswer
    CSpecialCommand
    CCode
    CSpeedAck
    CSFImpactar
    CBugReport
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public HandleDataSub(CMSG_COUNT) As Long

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1
    Endurance
    Intelligence
    Agility
    Willpower
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    mp
    ' Make sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor
    Helmet
    Shield
    ' Make sure Equipment_Count is below everything else
    Equipment_Count
End Enum

' Layers in a map
Public Enum MapLayer
    Ground = 1
    Mask
    Mask2
    Fringe
    Fringe2
    MaskAnim
    FringeAnim
    ' Make sure Layer_Count is below everything else
    Layer_Count
End Enum

' Sound entities
Public Enum SoundEntity
    seAnimation = 1
    seItem
    seNpc
    seResource
    seSpell
    seLevelUp
    seSwitch
    seSwitchFloor
    seSandstorm
    seSlide
    seAttack
    seCritical
    seHit
    seDie
    seReset
    seError
    ' Make sure SoundEntity_Count is below everything else
    SoundEntity_Count
End Enum

Public Enum MovementType
    NoMovement
    Onlydirectional
    ByDirection
    Bytile
    Random
    MovementType_Count
End Enum

Public Enum MomentType
    TileMatch = 0
    InCircleRange = 1
    InFrontRange = 2
    AtTimeInterval = 3
End Enum

Public Enum TriforceType

 TRIFORCE_WISDOM = 1
 TRIFORCE_COURAGE
 TRIFORCE_POWER
 TriforceType_Count

End Enum

Public Enum ChatType
    MapChat = 1
    GlobalChat
    PartyChat
    ClanChat
    WhisperChat
    SystemChat
    ChatType_Count
End Enum


Public Enum PlayerActionsType
    aAttack = 1
    aSpell
    aUseItem
    aMove
    aTeleport
    PlayerActions_Count
End Enum


Public Enum ShopPricesType
    SHItem = 0
    SHHeroKillPoints
    SHPKKillPoints
    SHQuestPoints
    SHNPCPoints
    SHBonusPoints
    ShopPricesTypeCount
End Enum

Public Enum PlayerStateType
    StateNone = 0
    StateSailing = 1
    StateRiding = 2
    Max_States
End Enum

Public Function StateToStr(ByVal State As PlayerStateType) As String
    Select Case State
    Case 0
        StateToStr = "None"
    Case 1
        StateToStr = "Sailing"
    Case 2
        StateToStr = "Riding"
    End Select
End Function








