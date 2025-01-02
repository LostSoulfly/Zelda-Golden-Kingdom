Attribute VB_Name = "modRegister"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public regasm As String, GtDLL As String, GtTLB As String

Public Sub registerDLL()
On Error Resume Next

If FileExist(regasm) Then

    FileCopy GtDLL, Environ("windir") & "\System32\GTranslateDLL.dll"
    FileCopy GtTLB, Environ("windir") & "\System32\GTranslateDLL.tlb"

    'FileCopy GtDLL, Environ("windir") & "\SysWOW64\GTranslateDLL.dll"
    'FileCopy GtTLB, Environ("windir") & "\SysWOW64\GTranslateDLL.tlb"
    
    If Not FileExist(Environ("windir") & "\System32\GTranslateDLL.dll") Then MsgBox "Please run this program as an Administrator!": Exit Sub
    If Not FileExist(Environ("windir") & "\System32\GTranslateDLL.tlb") Then MsgBox "Please run this program as an Administrator!": Exit Sub
    
    Label1.Caption = "Registering GTranslateDLL.."
    'Shell "cmd.exe /k " & regasm & " /tlb:" & Chr(34) & App.Path & "\GTranslateDLL.tlb" & Chr(34) & " " & Chr(34) & App.Path & "\GTranslateDLL.dll" & Chr(34), vbHide
        DoEvents
        Sleep 500
        
    Shell regasm & " /register " & "GTranslateDLL.dll", vbNormalFocus
        Label1.Caption = "GTranslateDLL codebase running.."
        DoEvents
        Sleep 500

    Shell regasm & " /codebase " & "GTranslateDLL.dll", vbNormalFocus
        Sleep 500
Else

MsgBox "regasm not found! Please install .net Framework 4.0"
End If

    'FileCopy GtDLL, Environ("windir") & "\SysWOW64\GTranslateDLL.dll"
    'FileCopy GtTLB, Environ("windir") & "\SysWOW64\GTranslateDLL.tlb"
    'FileCopy GtDLL, Environ("windir") & "\System32\GTranslateDLL.dll"
    'FileCopy GtTLB, Environ("windir") & "\System32\GTranslateDLL.dll"



End Sub
