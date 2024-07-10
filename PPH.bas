Attribute VB_Name = "PPH"
Sub PPHH1Text()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim para As paragraph
    For Each para In doc.Paragraphs
        If para.style = "H1" Then
            para.range.text = UCase(para.range.text)
        End If
    Next para
End Sub
