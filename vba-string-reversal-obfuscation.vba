Sub Demo_ReversedUrl()
    Dim hidden As String
    hidden = "moc.elpmaxe.www//:sptth"   ' reversed "https://www.example.com"

    Dim realUrl As String
    # ruleid: vba-string-reversal-obfuscation
    realUrl = StrReverse(hidden)

    MsgBox "Decoded URL: " & realUrl
End Sub