Attribute VB_Name = "modItem"
Option Explicit

Public Const INITIAL_MAX_WEIGHT As Long = 100000
Public Const ONE_VITAL_WEIGHT As Byte = 1   'unity
Public Const ONE_STAT_ADD_WEIGHT As Long = 10

Public Enum ItemTypes
    ITEM_TYPE_NONE = 0
    ITEM_TYPE_WEAPON
    ITEM_TYPE_ARMOR
    ITEM_TYPE_HELMET
    ITEM_TYPE_SHIELD
    ITEM_TYPE_CONSUME
    ITEM_TYPE_KEY
    ITEM_TYPE_CURRENCY
    ITEM_TYPE_SPELL
    ITEM_TYPE_RESET_POINTS
    ITEM_TYPE_TRIFORCE
    ITEM_TYPE_REDEMPTION
    ITEM_TYPE_CONTAINER
    ITEM_TYPE_BAG
    ITEM_TYPE_ADDWEIGHT
    ITEM_TYPE_RESIGN
    MAX_ITEM_TYPES 'this must be below
End Enum


Public Function CalculateItemWeight(ByVal ItemNum As Long) As Long
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function
    
    If IsItemUnitWeight(ItemNum) Then
        CalculateItemWeight = 1
        Exit Function
    End If
    
    Dim weight As Long
    weight = 0
    With item(ItemNum)
        If item(ItemNum).Type = ITEM_TYPE_CONSUME Then
            weight = weight + .AddEXP
            weight = weight + .AddHP
            weight = weight + .AddMP
            weight = weight + CalculateItemWeight(.ConsumeItem)
        Else 'sword, shield...
            weight = weight + GetEquipableItemStatsSum(ItemNum) * ONE_STAT_ADD_WEIGHT
        End If
    End With
            
End Function

Public Function GetEquipableItemStatsSum(ByVal ItemNum As Long) As Long
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function
    
    Dim ret As Long
    ret = 0
    With item(ItemNum)
    Dim i As Byte
    For i = 1 To Stats.Stat_Count - 1
        ret = ret + .Add_Stat(i)
    Next
    
    ret = ret + .Data2
    
    GetEquipableItemStatsSum = ret
    End With
End Function

Public Function GetPlayerWeight(ByVal index As Long) As Long
If index > 0 And index < MAX_PLAYERS Then
    GetPlayerWeight = TempPlayer(index).weight
End If
End Function
Public Function GetPlayerInvMaxWeight(ByVal index As Long) As Long
If index > 0 And index < MAX_PLAYERS Then
    GetPlayerInvMaxWeight = player(index).MaxWeight
End If
End Function
Public Function CanPlayerHoldWeight(ByVal index As Long, ByVal weight As Long) As Boolean
    If GetPlayerWeight(index) + weight <= GetPlayerMaxWeight(index) Then
        CanPlayerHoldWeight = True
    Else
        CanPlayerHoldWeight = False
    End If
End Function
Public Function GetItemValWeight(ByVal ItemNum As Long, Optional ByVal Itemvalue As Long = 1) As Long
    If isItemStackable(ItemNum) Then
        GetItemValWeight = GetItemWeight(ItemNum) * Itemvalue
    Else
        GetItemValWeight = GetItemWeight(ItemNum)
    End If
End Function
Public Sub SetPlayerWeight(ByVal index As Long, ByVal weight As Long)
    TempPlayer(index).weight = weight
End Sub

Public Function GetItemWeight(ByVal ItemNum As Long) As Long
If ItemNum > 0 And ItemNum < MAX_ITEMS Then
    GetItemWeight = item(ItemNum).weight
End If
End Function

Public Function IsItemUnitWeight(ByVal ItemNum As Long) As Boolean
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function
    
    Dim ItemType As Long
    ItemType = item(ItemNum).Type
    
    If ItemType = ITEM_TYPE_KEY Or ItemType = ITEM_TYPE_CURRENCY Or ItemType = ITEM_TYPE_SPELL Or ItemType = ITEM_TYPE_RESET_POINTS Or ItemType = ITEM_TYPE_TRIFORCE Or ItemType = ITEM_TYPE_REDEMPTION Or ItemType = ITEM_TYPE_CONTAINER Or ItemType = ITEM_TYPE_BAG Or ItemType = ITEM_TYPE_NONE Then
        IsItemUnitWeight = True
    End If
End Function

Public Function GetPlayerInvWeight(ByVal index As Long) As Long
If index > 0 And index < MAX_PLAYERS Then
    Dim i As Byte
    Dim res As Currency
    res = 0
    For i = 1 To MAX_INV
        Dim ItemNum As Long
        ItemNum = GetPlayerInvItemNum(index, i)
        If isItemStackable(ItemNum) Then
            If res + CCur(GetItemWeight(ItemNum)) * CCur(GetPlayerInvItemValue(index, i)) > MAX_LONG Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                SendInventoryUpdate index, i
            Else
                res = res + GetItemWeight(ItemNum) * GetPlayerInvItemValue(index, i)
            End If
        Else
            res = res + GetItemWeight(ItemNum)
        End If
    Next

    If res > MAX_LONG Then
        res = MAX_LONG
    End If
    
    GetPlayerInvWeight = CLng(res)
End If
End Function
Public Function GetPlayerInvSlotWeight(ByVal index As Long, ByVal slot As Byte) As Long
If index < 1 Or index > MAX_PLAYERS Or slot < 1 Or slot > MAX_INV Then Exit Function

GetPlayerInvSlotWeight = GetItemWeight(GetPlayerInvItemNum(index, slot))

End Function
Public Function CalculatePlayerWeight(ByVal index As Long) As Long
    Dim InvWeight As Long
    Dim EquipWeight As Long
    EquipWeight = GetPlayerEquipmentWeight(index)
    InvWeight = GetPlayerInvWeight(index)
    Dim ResultWeight As Currency
    'ResultWeight = InvWeight + EquipWeight
    If EquipWeight >= MAX_LONG Or InvWeight >= MAX_LONG Then
        ResultWeight = MAX_LONG
    Else
        ResultWeight = EquipWeight + InvWeight
    End If
    
    If ResultWeight > MAX_LONG Then
        ResultWeight = MAX_LONG
    End If
    CalculatePlayerWeight = CLng(ResultWeight)
End Function
Public Function GetPlayerEquipmentWeight(ByVal index As Long) As Long
If index > 0 And index < MAX_PLAYERS Then
    Dim i As Byte
    Dim res As Long
    For i = 1 To Equipment.Equipment_Count - 1
        res = res + GetItemWeight(GetPlayerEquipment(index, i))
    Next
    
    GetPlayerEquipmentWeight = res
End If
End Function
Public Function GetPlayerEquipmentSlotWeight(ByVal index As Long, ByVal slot As Byte) As Long
If index < 1 Or index > MAX_PLAYERS Or slot < 1 Or slot > Equipment.Equipment_Count Then Exit Function

    GetPlayerEquipmentSlotWeight = GetItemWeight(GetPlayerEquipment(index, slot))
End Function
Public Function GetPlayerMaxWeight(ByVal index As Long) As Long
If index > 0 And index < MAX_PLAYERS Then
    GetPlayerMaxWeight = player(index).MaxWeight
End If
End Function

Public Sub SetPlayerMaxWeight(ByVal index As Long, ByVal weight As Long)
    player(index).MaxWeight = weight
End Sub

Public Function IsPlayerOverWeight(ByVal index As Long) As Boolean
    IsPlayerOverWeight = False
    If GetPlayerWeight(index) > GetPlayerMaxWeight(index) Then
        IsPlayerOverWeight = True
    End If
End Function


Public Function IsItemMapped(ByVal mapnum As Long, ByVal index As Long) As Boolean
    If mapnum < 1 Or mapnum > MAX_MAPS Or index < 1 Or index > MAX_MAP_ITEMS Then Exit Function
    
    With MapItem(mapnum, index)
    
    If OutOfBoundries(.X, .Y, mapnum) Then Exit Function
    
    If map(mapnum).Tile(.X, .Y).Type = TILE_TYPE_ITEM Then
        If map(mapnum).Tile(.X, .Y).Data1 = .Num Then
            If map(mapnum).Tile(.X, .Y).Data2 = .Value Then
                IsItemMapped = True
            End If
        End If
    End If
    
    End With
End Function

Public Sub SetMapItemHighIndex(ByVal mapnum As Long, Optional ByVal StartIndex As Long = MAX_MAP_ITEMS)

Dim i As Long
Dim IndexChanged As Boolean
IndexChanged = False
For i = StartIndex To 1 Step -1
    If MapItem(mapnum, i).Num > 0 Then
        TempMap(mapnum).Item_highindex = i
        IndexChanged = True
        Exit For
    End If
Next

If Not IndexChanged Then
    'no items found
    TempMap(mapnum).Item_highindex = 0
End If

End Sub


Public Sub CheckMapItemHighIndex(ByVal mapnum As Long, ByVal ItemIndex As Long, ByVal spawn As Boolean)
    If ItemIndex <= 0 And ItemIndex > MAX_MAP_ITEMS Then Exit Sub
    
    Select Case spawn
    Case True
        If ItemIndex > TempMap(mapnum).Item_highindex Then
            TempMap(mapnum).Item_highindex = ItemIndex
        End If
    Case False
        If ItemIndex >= TempMap(mapnum).Item_highindex Then
            SetMapItemHighIndex mapnum, ItemIndex
        End If
    End Select
End Sub

Public Sub AddMapWaitingItem(ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    Dim i As Long
    
    If Not TempMap(mapnum).HasItems Then Exit Sub
    With TempMap(mapnum)
    For i = 1 To UBound(.WaitingForSpawnItems)
        If .WaitingForSpawnItems(i).Active = False Then
            .WaitingForSpawnItems(i).Active = True
            .WaitingForSpawnItems(i).Timer = GetRealTickCount + ITEM_DESPAWN_TIME
            .WaitingForSpawnItems(i).X = X
            .WaitingForSpawnItems(i).Y = Y
            Exit For
        End If
    Next
    End With
End Sub


Public Sub CheckMapWaitingItem(ByVal mapnum As Long)
    Dim i As Long
    If Not TempMap(mapnum).HasItems Then Exit Sub
    
    For i = 1 To UBound(TempMap(mapnum).WaitingForSpawnItems)
        With TempMap(mapnum).WaitingForSpawnItems(i)
        If .Active Then
            If Not OutOfBoundries(.X, .Y, mapnum) Then
                If .Timer < GetRealTickCount Then
                    Call SpawnItem(map(mapnum).Tile(.X, .Y).Data1, map(mapnum).Tile(.X, .Y).Data2, mapnum, .X, .Y, , False)
                    .Active = False
                    .Timer = 0
                    .X = 0
                    .Y = 0
                End If
            End If
        End If
        End With
    Next
End Sub

Function GetItemName(ByVal ItemNum As Long) As String
If ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function
GetItemName = Trim$(item(ItemNum).TranslatedName)
End Function

Function ItemExists(ByVal ItemNum As Long) As Boolean
If LenB(Trim$(item(ItemNum).Name)) > 0 And Asc(item(ItemNum).Name) <> 0 Then
    ItemExists = True
End If
End Function


Function GetItemSpeed(ByVal ItemNum As Long) As Long
    GetItemSpeed = item(ItemNum).Speed
End Function

