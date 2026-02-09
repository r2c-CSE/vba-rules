Sub Demo_MSXML2_XMLHTTP()
    Dim http As Object
    # ruleid: vba-downloader-objects
    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", "https://example.com", False
    http.Send

    MsgBox "Status: " & http.Status
End Sub

Sub Demo_Microsoft_XMLHTTP()
    Dim http As Object
    # ruleid: vba-downloader-objects
    Set http = CreateObject("Microsoft.XMLHTTP")

    http.Open "GET", "https://example.com", False
    http.Send

    MsgBox "Response length: " & Len(http.ResponseText)
End Sub

Sub Demo_WinHttpRequest()
    Dim req As Object
    # ruleid: vba-downloader-objects
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")

    req.Open "GET", "https://example.com", False
    req.Send

    MsgBox "Status: " & req.Status
End Sub
