Sub Demo_Environ_Paths()
    Dim profile As String, tmp As String
    # ruleid: vba-environment-discovery
    profile = Environ("USERPROFILE")
    # ruleid: vba-environment-discovery
    tmp = Environ("TEMP")

    MsgBox "Profile: " & profile & vbCrLf & "Temp: " & tmp
End Sub
