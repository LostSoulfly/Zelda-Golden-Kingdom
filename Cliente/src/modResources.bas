Attribute VB_Name = "modResources"


Public MapResources() As Long

Public Sub InitializeMapResources()
    With map
        ReDim MapResources(0 To .MaxX, 0 To .MaxY)
    End With
    
    Dim i As Long
    For i = 0 To Resource_Index
        MapResources(MapResource(i).x, MapResource(i).y) = i
    Next
End Sub

Public Function GetResourceIndex(ByVal x As Long, ByVal y As Long) As Long
    If x < 0 Or x > map.MaxX Or y < 0 Or y > map.MaxY Then
        GetResourceIndex = -1
    Else
        GetResourceIndex = MapResources(x, y)
    End If
    
End Function
