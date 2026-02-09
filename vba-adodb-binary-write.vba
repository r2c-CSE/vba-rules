Sub Demo_Stream_TextFile()
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2              ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText "Hello from ADODB.Stream!"
    # ruleid: vba-adodb-binary-write
    stm.SaveToFile "C:\Temp\demo.txt", 2   ' 2 = overwrite
    stm.Close
End Sub
