Attribute VB_Name = "modCustomSprite"
Option Explicit
Public Const MAX_DIRECTIONS As Byte = 4
Public Const MAX_SPRITE_ANIMS As Byte = 4
Public Const MAX_SPRITE_LAYERS As Byte = 5
Public Const MAX_CUSTOM_SPRITES As Byte = 15

Public CustomSprites(1 To MAX_CUSTOM_SPRITES) As CustomSpriteRec

Public Type FixedAnimRec
    EnabledAnims(MAX_DIRECTIONS - 1, MAX_SPRITE_ANIMS - 1) As Byte
    'this only has sense if |enabled anims| == 1
End Type

Public Type Point
    X As Integer
    Y As Integer
End Type

Public Type SpriteLayer
    Sprite As Long
    UseCenterPosition As Boolean
    UsePlayerSprite As Boolean
    fixed As FixedAnimRec
    CentersPositions(MAX_DIRECTIONS - 1) As Point 'from 0 to MAxDir - 1
End Type

Public Type CustomSpriteRec
    Name As String * NAME_LENGTH
    NLayers As Byte
    Layers() As SpriteLayer 'Numered from 1 to NLayers
End Type


Public Function GetCustomSpriteData(ByVal CustomSprite As Byte) As Byte()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    With CustomSprites(CustomSprite)
        Buffer.WriteString .Name
        Buffer.WriteByte .NLayers
        Dim i As Byte
        For i = 1 To .NLayers
            Buffer.WriteLong .Layers(i).Sprite
            Buffer.WriteByte .Layers(i).UseCenterPosition
            Buffer.WriteByte .Layers(i).UsePlayerSprite
            Dim j As Byte, k As Byte
            For j = 0 To MAX_DIRECTIONS - 1
                For k = 0 To MAX_SPRITE_ANIMS - 1
                    Buffer.WriteByte .Layers(i).fixed.EnabledAnims(j, k)
                Next
            Next
            For j = 0 To MAX_DIRECTIONS - 1
                 Buffer.WriteInteger .Layers(i).CentersPositions(j).X
                 Buffer.WriteInteger .Layers(i).CentersPositions(j).Y
            Next
        Next
                            
    End With
    
    GetCustomSpriteData = Buffer.ToArray
    Set Buffer = Nothing
End Function

Public Sub SetCustomSpriteData(ByVal CustomSprite As Byte, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    With CustomSprites(CustomSprite)
        .Name = Buffer.ReadString
        .NLayers = Buffer.ReadByte
        If .NLayers <> 0 Then
            ReDim .Layers(1 To .NLayers)
        End If
        Dim i As Byte
        For i = 1 To .NLayers
            .Layers(i).Sprite = Buffer.ReadLong
            .Layers(i).UseCenterPosition = Buffer.ReadByte
            .Layers(i).UsePlayerSprite = Buffer.ReadByte
            Dim j As Byte, k As Byte
            For j = 0 To MAX_DIRECTIONS - 1
                For k = 0 To MAX_SPRITE_ANIMS - 1
                    .Layers(i).fixed.EnabledAnims(j, k) = Buffer.ReadByte
                Next
            Next
            For j = 0 To MAX_DIRECTIONS - 1
                .Layers(i).CentersPositions(j).X = Buffer.ReadInteger
                .Layers(i).CentersPositions(j).Y = Buffer.ReadInteger
            Next
        Next
                            
    End With
    
    
    Set Buffer = Nothing
End Sub


