Sub Demo_StoreVariable()
    # ruleid: vba-doc-var-hiding
    ActiveDocument.Variables.Add Name:="HelperKey", Value:="HelloWorld"
End Sub


Sub Demo_FragmentStore()
    # ruleid: vba-doc-var-hiding
    ActiveDocument.Variables.Add Name:="PartA", Value:="dlroW"
    
    # ruleid: vba-doc-var-hiding
    ActiveDocument.Variables.Add Name:="PartB", Value:=" olleH"
End Sub