Sub Demo_HiddenShell()
    # ruleid: vba-hidden-window
    Shell "calc.exe", 0   ' 0 = hidden window
End Sub

Sub Demo_HiddenVariable()
    Dim cmd As String
    cmd = "notepad.exe"
    # ruleid: vba-hidden-window
    Shell cmd, 0
End Sub

Sub Demo_ObfuscatedHidden()
    Dim s As String
    s = StrReverse("exe.clac")   ' "calc.exe"
    # ruleid: vba-hidden-window
    Shell s, 0
End Sub
