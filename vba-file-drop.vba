Sub Demo_FSO_CreateTextFile()
    Dim fs As Object
    Dim f As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
    # ruleid: vba-file-drop
    Set f = fs.CreateTextFile("C:\Temp\hello.txt", True)

    f.WriteLine "Hello from VBA"
    f.Close
End Sub

Sub Demo_Open_For_Output()
    Dim filePath As String
    filePath = "C:\Temp\log.txt"

    # ruleid: vba-file-drop
    Open filePath For Output As #1
    Print #1, "Log line from VBA"
    Close #1
End Sub
