Attribute VB_Name = "modUpdate"
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal _
lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Url As String

Public Sub HandleUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    Url = Buffer.ReadString
    If MsgBox("Hay una actualizacion disponible, instrucciones: " & vbNewLine & _
               Buffer.ReadString & vbNewLine & _
               "Link de descarga: " & _
               Url, vbYesNo) = vbYes Then
        ShellExecute 1, "Open", Url, "", 0&, 1
    End If
   
    frmLoad.Visible = False
    frmMenu.Visible = True
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    frmMenu.picMain.Visible = True
End Sub

