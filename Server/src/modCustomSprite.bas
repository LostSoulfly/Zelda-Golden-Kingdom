Attribute VB_Name = "modCustomSprite"
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
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    With CustomSprites(CustomSprite)
        buffer.WriteString .Name
        buffer.WriteByte .NLayers
        Dim i As Byte
        For i = 1 To .NLayers
            buffer.WriteLong .Layers(i).Sprite
            buffer.WriteByte .Layers(i).UseCenterPosition
            buffer.WriteByte .Layers(i).UsePlayerSprite
            Dim j As Byte, k As Byte
            For j = 0 To MAX_DIRECTIONS - 1
                For k = 0 To MAX_SPRITE_ANIMS - 1
                    buffer.WriteByte .Layers(i).fixed.EnabledAnims(j, k)
                Next
            Next
            For j = 0 To MAX_DIRECTIONS - 1
                 buffer.WriteInteger .Layers(i).CentersPositions(j).X
                 buffer.WriteInteger .Layers(i).CentersPositions(j).Y
            Next
        Next
                            
    End With
    
    GetCustomSpriteData = buffer.ToArray
    Set buffer = Nothing
End Function

Public Sub SetCustomSpriteData(ByVal CustomSprite As Byte, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    With CustomSprites(CustomSprite)
        .Name = buffer.ReadString
        .NLayers = buffer.ReadByte
        If .NLayers <> 0 Then
            ReDim .Layers(1 To .NLayers)
        End If
        Dim i As Byte
        For i = 1 To .NLayers
            .Layers(i).Sprite = buffer.ReadLong
            .Layers(i).UseCenterPosition = buffer.ReadByte
            .Layers(i).UsePlayerSprite = buffer.ReadByte
            Dim j As Byte, k As Byte
            For j = 0 To MAX_DIRECTIONS - 1
                For k = 0 To MAX_SPRITE_ANIMS - 1
                    .Layers(i).fixed.EnabledAnims(j, k) = buffer.ReadByte
                Next
            Next
            For j = 0 To MAX_DIRECTIONS - 1
                .Layers(i).CentersPositions(j).X = buffer.ReadInteger
                .Layers(i).CentersPositions(j).Y = buffer.ReadInteger
            Next
        Next
                            
    End With
    
    
    Set buffer = Nothing
End Sub


