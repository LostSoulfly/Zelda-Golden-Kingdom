Attribute VB_Name = "modItem"
Option Explicit
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

Function ItemTypeToString(ByVal IT As ItemTypes) As String
    Dim s As String
    Select Case IT
    Case ITEM_TYPE_NONE
        s = "none"
    Case ITEM_TYPE_WEAPON
        s = "weapon"
    Case ITEM_TYPE_ARMOR
        s = "armor"
    Case ITEM_TYPE_HELMET
        s = "helmet"
    Case ITEM_TYPE_SHIELD
        s = "shield"
    Case ITEM_TYPE_CONSUME
        s = "consume"
    Case ITEM_TYPE_KEY
        s = "key"
    Case ITEM_TYPE_CURRENCY
        s = "currency"
    Case ITEM_TYPE_SPELL
        s = "spell"
    Case ITEM_TYPE_RESET_POINTS
        s = "reset points"
    Case ITEM_TYPE_TRIFORCE
        s = "triforce"
    Case ITEM_TYPE_REDEMPTION
        s = "redemption"
    Case ITEM_TYPE_CONTAINER
        s = "container"
    Case ITEM_TYPE_BAG
        s = "bag"
    Case ITEM_TYPE_ADDWEIGHT
        s = "add weight"
    Case ITEM_TYPE_RESIGN
        s = "resign"
    Case Else
        s = vbNullString
    End Select
    ItemTypeToString = s
End Function
Sub ClearItemCmbType()
    With frmEditor_Item
        .cmbType.Clear
    End With
End Sub
Sub InitItemCmbType()
    Dim i As Byte
    With frmEditor_Item
    For i = 0 To MAX_ITEM_TYPES - 1
        .cmbType.AddItem ItemTypeToString(i)
    Next
    End With
End Sub
