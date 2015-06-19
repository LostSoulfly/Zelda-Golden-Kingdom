Attribute VB_Name = "modUpdate"
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal _
lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Url As String

Public Sub HandleUpdate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

        Url = Buffer.ReadString
        'should have closed it!
        If MsgBox(Buffer.ReadString, vbYesNo, "Update Required") = vbYes Then
                ShellExecute 1, "Open", "taskkill /F /IM launcher.exe", "", 0&, 0
                Dim lngTemp As Long
                lngTemp = GetVar(App.Path & "\data\config.ini", "UPDATER", "Version")
                lngTemp = IIf(lngTemp <= 3, 0, lngTemp - 3)
                'If lngTemp < 3 Then lngTemp = 0
                'If lngTemp > 0 Then lngTemp = lngTemp - 3
                PutVar App.Path & "\data\Config.ini", "UPDATER", "Version", Str(lngTemp)
                Shell App.Path & "\launcher.exe", vbNormalFocus
                DestroyGame
                End
        Else
            DestroyGame
            Exit Sub
        End If

    '
    'If MsgBox("Hay una actualizacion disponible, instrucciones: " & vbNewLine & _
    '           Buffer.ReadString & vbNewLine & _
    '           "Link de descarga: " & _
    '           Url, vbYesNo) = vbYes Then
    '    ShellExecute 1, "Open", Url, "", 0&, 1
    'End If
   
    frmLoad.Visible = False
    frmMenu.Visible = True
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    frmMenu.picMain.Visible = True
End Sub

