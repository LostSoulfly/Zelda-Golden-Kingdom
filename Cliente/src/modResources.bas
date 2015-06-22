Attribute VB_Name = "modResources"


Public MapResources() As Long

Public Sub InitializeMapResources()
    With map
        ReDim MapResources(0 To .MaxX, 0 To .MaxY)
    End With
    DoEvents
    Dim i As Long
    For i = 0 To Resource_Index
    If MapResource(i).X > map.MaxX Then
        Debug.Print "dangit"
        ReDim Preserve MapResources(map.MaxX, map.MaxY)
        GoTo cont
    End If
    If MapResource(i).y > map.MaxY Then
        Debug.Print "flux capacitor, man."
        ReDim Preserve MapResources(map.MaxX, map.MaxY)
        GoTo cont
    End If
        MapResources(MapResource(i).X, MapResource(i).y) = i
cont:
    Next
End Sub

Public Function GetResourceIndex(ByVal X As Long, ByVal y As Long) As Long
    If X < 0 Or X > map.MaxX Or y < 0 Or y > map.MaxY Then
        GetResourceIndex = -1
    Else
        GetResourceIndex = MapResources(X, y)
    End If
    
End Function
