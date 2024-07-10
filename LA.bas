Attribute VB_Name = "LA"
Sub LAH1()
    Dim doc As Document
    Dim para As paragraph
    
    Set doc = ActiveDocument
    
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph style is 'H1' or 'EH'
        If para.style = "H1" Or para.style = "EH" Then
            ' Set the font style to bold
            para.range.Font.Bold = True
        End If
    Next para
    
    Set doc = Nothing
    Set para = Nothing
End Sub
Sub LAH2()
    Dim doc As Document
    Dim para As paragraph
    
    Set doc = ActiveDocument
    
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph style is 'H1' or 'EH'
        If para.style = "H2" Or para.style = "H3" Then
            ' Set the font style to bold
            para.range.Font.Italic = True
        End If
    Next para
    
    Set doc = Nothing
    Set para = Nothing
End Sub
Sub LAheadlevel()
'
' LAheadlevel Macro
'
'
Call LAH1

Call LAH2

End Sub

Sub LAetal()
    Dim searchTerm As String
    Dim rng As range
    
    ' Set the search term with wildcard
    searchTerm = "et al."
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight the search term with the specified style
    With rng.Find
        .text = searchTerm
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        
        ' Loop through each found item
        Do While .Execute
            rng.HighlightColorIndex = wdRed
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub
