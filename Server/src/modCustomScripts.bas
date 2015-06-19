Attribute VB_Name = "modCustomScripts"
Public Sub CustomScript(index As Long, caseID As Long)
    Select Case caseID
        Case Else
            PlayerMsg index, "You just activated custom script " & caseID & ". This script is not yet programmed.", BrightRed
    End Select
End Sub
