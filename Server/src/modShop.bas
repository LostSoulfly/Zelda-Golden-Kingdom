Attribute VB_Name = "modShop"
Option Explicit


Public Enum ShopPricesType
    SHItem = 0
    SHHeroKillPoints
    SHPKKillPoints
    SHQuestPoints
    SHNPCPoints
    SHBonusPoints
    ShopPricesTypeCount
End Enum



Private Type TradeItemRec
    item As Long
    Itemvalue As Long
    CostItem As Long
    Costvalue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
    PriceType As Byte
    TranslatedName As String * NAME_LENGTH
End Type


Public Shop(1 To MAX_SHOPS) As ShopRec


Sub BuyItem(ByVal index As Long, ByVal shopnum As Long, ByVal shopslot As Long)

    With Shop(shopnum).TradeItem(shopslot)
        ' check trade exists
        If .item < 1 Then Exit Sub
            
        If ProcessPlayerCostItem(index, shopnum, shopslot) Then
        
            Dim GivenValue As Long
            Dim i As Long
            i = CanGiveItem(index, .item, .Itemvalue, GivenValue)
            If i > 0 Then
                GiveInvSlot index, i, .item, GivenValue
                PlayerMsg index, "Compra realizada.", BrightGreen ' send confirmation message & reset their shop action
            'Else
                'GiveInvItem index, .CostItem, .Costvalue 'return the money
            End If
        End If
    End With
    
    
End Sub


Function ProcessPlayerCostItem(ByVal index As Long, ByVal shopnum As Long, ByVal shopslot As Long) As Boolean
    With Shop(shopnum).TradeItem(shopslot)
    
    If Not CanPlayerHoldWeight(index, GetItemValWeight(.item, .Itemvalue)) Then
        PlayerMsg index, "No Puedes soportar el peso del objeto!", BrightRed
        Exit Function
    End If
    
    Dim CostAmount As Long
    CostAmount = .Costvalue
    
    Dim points As Long
    
    Select Case Shop(shopnum).PriceType
    Case SHItem
        Dim CostItem As Long
        CostItem = .CostItem
        
        Dim ItemAmount As Long
        ItemAmount = HasItem(index, CostItem)
        If ItemAmount > 0 And ItemAmount >= CostAmount Then
            TakeInvItem index, CostItem, CostAmount
            ProcessPlayerCostItem = True
        Else
            PlayerMsg index, "No posees suficiente dinero para adquirir éste objeto.", BrightRed
        End If
        
    Case SHPKKillPoints
        points = GetPlayerKillPoints(index, PK_PLAYER)
        
        If points > 0 And points >= CostAmount Then
            SetPlayerKillPoints index, points - CostAmount, PK_PLAYER
            ProcessPlayerCostItem = True
        Else
            PlayerMsg index, "No posees suficientes puntos!", BrightRed
        End If
        
    Case SHHeroKillPoints
        points = GetPlayerKillPoints(index, HERO_PLAYER)
        
        If points > 0 And points >= CostAmount Then
            SetPlayerKillPoints index, points - CostAmount, HERO_PLAYER
            ProcessPlayerCostItem = True
        Else
            PlayerMsg index, "No posees suficientes puntos!", BrightRed
        End If
        
    Case SHQuestPoints
    
    Case SHNPCPoints
    
    Case SHBonusPoints
        points = GetPlayerBonusPoints(index)
        
        If points > 0 And points >= CostAmount Then
            SetPlayerBonusPoints index, points - CostAmount
            SendPlayerBonusPoints index
            ProcessPlayerCostItem = True
        Else
            PlayerMsg index, "No posees suficientes puntos!", BrightRed
        End If
    End Select
    
    End With
End Function


