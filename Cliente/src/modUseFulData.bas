Attribute VB_Name = "modUseFulData"
Option Explicit


Sub SetSpellUsefulData(ByVal spellnum As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    
    With Spell(spellnum)
    .Name = Buffer.ReadString
    .TranslatedName = Buffer.ReadString
    .Desc = Buffer.ReadString
    .Sound = Buffer.ReadString
    
    .MPCost = Buffer.ReadLong
    .CDTime = Buffer.ReadLong
    .Icon = Buffer.ReadLong
    .CastTime = Buffer.ReadLong
    End With
    
    Set Buffer = Nothing
End Sub

Sub SetNPCUsefulData(ByVal NPCNum As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    
    With NPC(NPCNum)
    .Name = Buffer.ReadString
    .TranslatedName = Buffer.ReadString
    .Sound = Buffer.ReadString
    .sprite = Buffer.ReadLong
    
    .Behaviour = Buffer.ReadByte
    Dim i As Byte
    For i = 1 To Stats.Stat_Count - 1
        .stat(i) = Buffer.ReadByte
    Next
    .HP = Buffer.ReadLong
    .QuestNum = Buffer.ReadLong
    .Speed = Buffer.ReadLong
    .Level = Buffer.ReadLong
    End With
    
    Set Buffer = Nothing
End Sub

Sub SetItemUsefulData(ByVal ItemNum As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    
    With Item(ItemNum)
    .Name = Buffer.ReadString
    .TranslatedName = Buffer.ReadString
    .Desc = Buffer.ReadString
    .Sound = Buffer.ReadString
    
    .Pic = Buffer.ReadLong
    .Type = Buffer.ReadByte
    .Price = Buffer.ReadLong
    .Rarity = Buffer.ReadByte
    .Speed = Buffer.ReadLong
    .Paperdoll = Buffer.ReadLong
    .ProjecTile.Pic = Buffer.ReadLong
    .Weight = Buffer.ReadLong

    End With
    
    Set Buffer = Nothing
End Sub

Sub SetResourceUsefulData(ByVal ResourceNum As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    
    With Resource(ResourceNum)
    .Name = ""
    .ResourceImage = Buffer.ReadLong
    .ExhaustedImage = Buffer.ReadLong
    .WalkableNormal = Buffer.ReadByte
    .WalkableExhausted = Buffer.ReadByte
    End With
    
    Set Buffer = Nothing
End Sub

Sub SetAnimationUseFulData(ByVal AnimNum As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    
    With Animation(AnimNum)
    .Sound = Buffer.ReadString
    Dim i As Byte
    For i = 0 To 1
        .sprite(i) = Buffer.ReadLong
        .Frames(i) = Buffer.ReadLong
        .LoopCount(i) = Buffer.ReadLong
        .looptime(i) = Buffer.ReadLong
    Next
    End With
    
    Set Buffer = Nothing
End Sub
