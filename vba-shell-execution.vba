Sub Demo_Shell_Notepad()
    ' Opens Notepad - harmless example
    #ruleid: vba-shell-execution
    Shell "notepad.exe", vbNormalFocus
End Sub

Sub Demo_WScriptShell_Run()
    Dim sh As Object
    #ruleid: vba-shell-execution
    Set sh = CreateObject("WScript.Shell").Run("calc.exe")
End Sub

Sub Demo_WScriptShell_Run()
    Dim sh As Object
    #ruleid: vba-shell-execution
    Set sh = CreateObject("WScript.Shell")

    ' Harmless: opens Calculator
    sh.Run "calc.exe", 1, False
End Sub

Sub Demo_WScriptShell_Exec()
    Dim sh As Object
    Dim proc As Object

    #ruleid: vba-shell-execution
    Set sh = CreateObject("WScript.Shell")
    Set proc = sh.Exec("cmd /c echo Hello from VBA")

    MsgBox proc.StdOut.ReadAll
End Sub


