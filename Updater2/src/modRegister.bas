Attribute VB_Name = "modRegister"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public regasm As String, GtDLL As String, GtTLB As String

Public Sub registerDLL()
On Error Resume Next


If FileExist(Environ("windir") & "\System32\GTranslate.dll") Then Exit Sub

If FileExist(regasm) Then

    FileCopy GtDLL, Environ("windir") & "\System32\GTranslate.dll"
    FileCopy GtTLB, Environ("windir") & "\System32\GTranslate.tlb"

    'FileCopy GtDLL, Environ("windir") & "\SysWOW64\GTranslate.dll"
    'FileCopy GtTLB, Environ("windir") & "\SysWOW64\GTranslate.tlb"
    
    If Not FileExist(Environ("windir") & "\System32\GTranslate.dll") Then MsgBox "Please run this program as an Administrator!": Exit Sub
    If Not FileExist(Environ("windir") & "\System32\GTranslate.tlb") Then MsgBox "Please run this program as an Administrator!": Exit Sub
    
    Label1.Caption = "Registering GTranslate.."
    'Shell "cmd.exe /k " & regasm & " /tlb:" & Chr(34) & App.Path & "\GTranslate.tlb" & Chr(34) & " " & Chr(34) & App.Path & "\GTranslate.dll" & Chr(34), vbHide
        DoEvents
        Sleep 500
        
    Shell regasm & " /register " & "GTranslate.dll", vbNormalFocus
        Label1.Caption = "GTranslate codebase running.."
        DoEvents
        Sleep 500

    Shell regasm & " /codebase " & "GTranslate.dll", vbNormalFocus
        Sleep 500
Else

MsgBox "regasm not found! Please install .net Framework 4.0"
End If

    'FileCopy GtDLL, Environ("windir") & "\SysWOW64\GTranslate.dll"
    'FileCopy GtTLB, Environ("windir") & "\SysWOW64\GTranslate.tlb"
    'FileCopy GtDLL, Environ("windir") & "\System32\GTranslate.dll"
    'FileCopy GtTLB, Environ("windir") & "\System32\GTranslate.dll"



End Sub
