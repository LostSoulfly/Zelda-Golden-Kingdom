Attribute VB_Name = "modUseFulData"
Option Explicit


Sub SetSpellUsefulData(ByVal spellnum As Long, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    
    With Spell(spellnum)
    .Name = buffer.ReadString
    .TranslatedName = buffer.ReadString
    .Desc = buffer.ReadString
    .Sound = buffer.ReadString
    
    .MPCost = buffer.ReadLong
    .CDTime = buffer.ReadLong
    .Icon = buffer.ReadLong
    .CastTime = buffer.ReadLong
    End With
    
    Set buffer = Nothing
End Sub

Sub SetNPCUsefulData(ByVal NPCNum As Long, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    
    With NPC(NPCNum)
    .Name = buffer.ReadString
    .TranslatedName = buffer.ReadString
    .Sound = buffer.ReadString
    .sprite = buffer.ReadLong
    
    .Behaviour = buffer.ReadByte
    Dim i As Byte
    For i = 1 To Stats.Stat_Count - 1
        .stat(i) = buffer.ReadByte
    Next
    .HP = buffer.ReadLong
    .QuestNum = buffer.ReadLong
    .Speed = buffer.ReadLong
    .Level = buffer.ReadLong
    End With
    
    Set buffer = Nothing
End Sub

Sub SetItemUsefulData(ByVal ItemNum As Long, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    
    With Item(ItemNum)
    .Name = buffer.ReadString
    .TranslatedName = buffer.ReadString
    .Desc = buffer.ReadString
    .Sound = buffer.ReadString
    
    .Pic = buffer.ReadLong
    .Type = buffer.ReadByte
    .Price = buffer.ReadLong
    .Rarity = buffer.ReadByte
    .Speed = buffer.ReadLong
    .Paperdoll = buffer.ReadLong
    .ProjecTile.Pic = buffer.ReadLong
    .Weight = buffer.ReadLong

    End With
    
    Set buffer = Nothing
End Sub

Sub SetResourceUsefulData(ByVal ResourceNum As Long, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    
    With Resource(ResourceNum)
    .Name = ""
    .ResourceImage = buffer.ReadLong
    .ExhaustedImage = buffer.ReadLong
    .WalkableNormal = buffer.ReadByte
    .WalkableExhausted = buffer.ReadByte
    End With
    
    Set buffer = Nothing
End Sub

Sub SetAnimationUseFulData(ByVal AnimNum As Long, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    
    With Animation(AnimNum)
    .Sound = buffer.ReadString
    Dim i As Byte
    For i = 0 To 1
        .sprite(i) = buffer.ReadLong
        .Frames(i) = buffer.ReadLong
        .LoopCount(i) = buffer.ReadLong
        .looptime(i) = buffer.ReadLong
    Next
    End With
    
    Set buffer = Nothing
End Sub
