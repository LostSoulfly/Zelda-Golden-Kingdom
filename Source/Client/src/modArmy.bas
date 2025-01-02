Attribute VB_Name = "modArmy"
Option Explicit

Public Enum HeroRangesType
    Soldado = 1
    Escolta
    Teniente
    Capitan
    Protector
    Caballero
    HeroRangesTypeCount
End Enum

Public Enum PKRangesType
    Mercenario = 1
    Aniquilador
    Devastador
    Asolador
    Comandante
    Elite
    PkRangesTypeCount
End Enum


'Justice Status Constants
Public Const NONE_PLAYER As Byte = 0
Public Const PK_PLAYER As Byte = 1
Public Const HERO_PLAYER As Byte = 2


Function RangeToStr(ByVal range As Byte, ByVal army As Byte) As String
Select Case army
Case HERO_PLAYER
    Select Case range
    Case Soldado
        RangeToStr = "Soldier"
    Case Escolta
        RangeToStr = "Escort"
    Case Teniente
        RangeToStr = "Lieutenant"
    Case Capitan
        RangeToStr = "Capitan"
    Case Protector
        RangeToStr = "Protector"
    Case Caballero
        RangeToStr = "Knight"
    Case Else
        RangeToStr = "None"
    End Select
Case PK_PLAYER
    Select Case range
    Case Mercenario
        RangeToStr = "Mercenary"
    Case Aniquilador
        RangeToStr = "Annihilator"
    Case Devastador
        RangeToStr = "Devastating"
    Case Asolador
        RangeToStr = "Ravager"
    Case Comandante
        RangeToStr = "Commander"
    Case Elite
        RangeToStr = "Elite"
    Case Else
        RangeToStr = "None"
    End Select
Case NONE_PLAYER
    RangeToStr = "None"
End Select
End Function

Function JusticeToStr(ByVal Justice As Byte) As String
    If Justice = HERO_PLAYER Then
        JusticeToStr = "Hero"
    ElseIf Justice = PK_PLAYER Then
        JusticeToStr = "Killer"
    Else
        JusticeToStr = "Neutral"
    End If
End Function


Sub ItemEditorInitRanges()
    Dim ItemArmyTypeReq As Byte
    Dim ItemArmyRangeReq As Byte
    ItemArmyTypeReq = Item(EditorIndex).ArmyType_Req
    ItemArmyRangeReq = Item(EditorIndex).ArmyRange_Req
    
    With frmEditor_Item
    .scrlArmyTypeReq.Min = NONE_PLAYER
    .scrlArmyTypeReq.Max = HERO_PLAYER
    
    .scrlArmyRangeReq.Min = 0
    
    
    
    If ItemArmyTypeReq >= 0 And ItemArmyTypeReq <= HERO_PLAYER Then
        .scrlArmyTypeReq = ItemArmyTypeReq
        If ItemArmyTypeReq = PK_PLAYER Then
             .scrlArmyRangeReq.Max = PkRangesTypeCount - 1
        ElseIf ItemArmyTypeReq = HERO_PLAYER Then
            .scrlArmyRangeReq.Max = HeroRangesTypeCount - 1
        ElseIf ItemArmyTypeReq = NONE_PLAYER Then
            .scrlArmyRangeReq.Max = 0
        End If
        
        .scrlArmyRangeReq.value = ItemArmyRangeReq
    End If
    
    End With
    
   
End Sub

Sub CheckItemEditorRangeScrolls()
    With frmEditor_Item
    .scrlArmyRangeReq.Min = 0
    Dim ItemArmyTypeReq As Byte
    ItemArmyTypeReq = .scrlArmyTypeReq.value
    If ItemArmyTypeReq = PK_PLAYER Then
         .scrlArmyRangeReq.Max = PkRangesTypeCount - 1
    ElseIf ItemArmyTypeReq = HERO_PLAYER Then
        .scrlArmyRangeReq.Max = HeroRangesTypeCount - 1
    ElseIf ItemArmyTypeReq = NONE_PLAYER Then
        .scrlArmyRangeReq.Max = 0
    End If
    
    .scrlArmyRangeReq.value = .scrlArmyRangeReq.value
    
    End With
    
End Sub

