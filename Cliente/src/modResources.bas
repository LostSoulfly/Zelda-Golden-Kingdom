Attribute VB_Name = "modResources"
Option Explicit


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
        AddText "There was an issue loading this map. Please leave this map and try again!", vbRed
        'ReDim Preserve MapResources(map.MaxX, map.MaxY)
        GoTo cont
    End If
    If MapResource(i).Y > map.MaxY Then
        Debug.Print "flux capacitor, man."
        AddText "There was an issue loading this map. Please leave this map and try again!", vbRed
        'ReDim Preserve MapResources(map.MaxX, map.MaxY)
        GoTo cont
    End If
        MapResources(MapResource(i).X, MapResource(i).Y) = i
cont:
    Next
    
    
End Sub

Public Function GetResourceIndex(ByVal X As Long, ByVal Y As Long) As Long
On Error GoTo oops
    If X < 0 Or X > map.MaxX Or Y < 0 Or Y > map.MaxY Then
        GetResourceIndex = -1
    Else
        GetResourceIndex = MapResources(X, Y)
    End If
Exit Function
oops:
    InitializeMapResources
End Function
