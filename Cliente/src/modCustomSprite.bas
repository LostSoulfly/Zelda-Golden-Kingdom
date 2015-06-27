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
    sprite As Long
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
            Buffer.WriteLong .Layers(i).sprite
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
            .Layers(i).sprite = Buffer.ReadLong
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


Public Function GetCustomSpriteLayer(ByRef Csprite As CustomSpriteRec, ByVal layer As Byte) As SpriteLayer
    If layer > Csprite.NLayers - 1 Then Exit Function
    
    GetCustomSpriteLayer = Csprite.Layers(layer)
End Function

Public Function GetCustomSpriteNLayers(ByRef Csprite As CustomSpriteRec) As Byte
    GetCustomSpriteNLayers = Csprite.NLayers
End Function


Public Function IsLayerUsingPlayerSprite(ByRef layer As SpriteLayer) As Boolean
    IsLayerUsingPlayerSprite = layer.UsePlayerSprite
End Function

Public Sub SetLayerUsePlayerSprite(ByRef layer As SpriteLayer, ByVal use As Boolean)
    layer.UsePlayerSprite = use
End Sub

'Public Function IsRideAnim(ByRef fixed As FixedAnimRec) As Boolean
    'IsRideAnim = fixed.RideEffect
'End Function

Public Function GetSpriteLayerFixed(ByRef layer As SpriteLayer) As FixedAnimRec
    GetSpriteLayerFixed = layer.fixed
End Function

Public Function IsLayerUsingCenter(ByRef layer As SpriteLayer) As Boolean
    IsLayerUsingCenter = layer.UseCenterPosition
End Function

Public Function GetLayerSprite(ByRef layer As SpriteLayer) As Long
    GetLayerSprite = layer.sprite
End Function

Public Function GetLayerCenterX(ByRef layer As SpriteLayer, ByVal dir As Byte) As Integer
    If dir >= MAX_DIRECTIONS Then Exit Function
    GetLayerCenterX = layer.CentersPositions(dir).X
End Function

Public Function GetLayerCenterY(ByRef layer As SpriteLayer, ByVal dir As Byte) As Integer
    If dir >= MAX_DIRECTIONS Then Exit Function
    GetLayerCenterY = layer.CentersPositions(dir).Y
End Function

Public Function GetAnimFromCurrentAnim(ByRef fixed As FixedAnimRec, ByVal dir As Byte, ByVal AnimNum As Byte) As Byte
    If AnimNum >= MAX_SPRITE_ANIMS Then Exit Function
    GetAnimFromCurrentAnim = fixed.EnabledAnims(dir, AnimNum)
End Function

Public Function GetClosestAnimFromOne(ByRef fixed As FixedAnimRec, ByVal AnimNum As Byte) As Byte

End Function

Public Sub AddEmptyLayer(ByRef Csprite As CustomSpriteRec)
    If Csprite.NLayers >= MAX_SPRITE_LAYERS Then Exit Sub 'can't add new layer
    
    Csprite.NLayers = Csprite.NLayers + 1
    ReDim Preserve Csprite.Layers(1 To Csprite.NLayers)
End Sub

Public Sub AddLayer(ByRef Csprite As CustomSpriteRec, ByVal Actual As Byte)
    If Csprite.NLayers >= MAX_SPRITE_LAYERS Then Exit Sub
    
    Csprite.NLayers = Csprite.NLayers + 1
    ReDim Preserve Csprite.Layers(1 To Csprite.NLayers)
    
    If Actual > 0 Then 'so we had at least 1 element
        Dim i As Byte
        i = Csprite.NLayers
        While (i > Actual + 1)
            Csprite.Layers(i) = Csprite.Layers(i - 1)
            i = i - 1
        Wend
        
        'erase the new layer, that is a provisonaly copy
        Call ZeroMemory(ByVal VarPtr(Csprite.Layers(Actual)), LenB(Csprite.Layers(Actual)))
    End If
            
        
    
End Sub

Public Sub DeleteLayer(ByRef Csprite As CustomSpriteRec, ByVal layer As Byte)
    If layer < 1 Or layer > Csprite.NLayers Then Exit Sub 'prevent errors
    
    Dim i As Byte
    For i = layer To Csprite.NLayers - 1
        Csprite.Layers(i) = Csprite.Layers(i + 1)
    Next
    
    Csprite.NLayers = Csprite.NLayers - 1
    If Csprite.NLayers > 0 Then
        ReDim Preserve Csprite.Layers(1 To Csprite.NLayers)
    End If
End Sub

Public Sub SetLayerSprite(ByRef layer As SpriteLayer, ByVal sprite As Long)
    If sprite < 1 Or sprite > NumCharacters Then Exit Sub
    
    layer.sprite = sprite
  
End Sub

Public Sub SetLayerFixedAnims(ByRef fixed As FixedAnimRec, ByVal AnimNum As Byte, ByVal enabled As Boolean)
    If AnimNum < MAX_SPRITE_ANIMS Then Exit Sub
    
    'fixed.EnabledAnims(AnimNum) = enabled

End Sub

Public Sub SetLayerCenterPosition(ByRef layer As SpriteLayer, ByVal enabled As Boolean)
    layer.UseCenterPosition = enabled
End Sub

Public Sub SetLayerCenterPositions(ByRef layer As SpriteLayer, ByVal dir As Byte, ByVal X As Integer, ByVal Y As Integer)
    If X < 0 Or Y < 0 Or dir >= MAX_DIRECTIONS Then Exit Sub
    
    layer.CentersPositions(dir).X = X
    layer.CentersPositions(dir).Y = Y

End Sub




