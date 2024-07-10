Attribute VB_Name = "NewMacros1"
Sub Journalcheck()
    Dim doc As Document
    Dim rng As range
    Dim found As Boolean
    Dim styleName As String
    Dim paraRange As range
    Dim volFound As Boolean
    Dim firstPageFound As Boolean
    Dim lastPageFound As Boolean
    Dim authorText As String
    Dim Year As String
    Dim p As Variant ' Define p as Variant type
    Dim leftQuote As String
    Dim rightQuote As String
    
    ' Unicode values for left and right curly quotes
    leftQuote = ChrW(8220)
    rightQuote = ChrW(8221)
    
    ' Set reference to the active document
    Set doc = ActiveDocument
    
    ' Set the style name to search for
    styleName = "stl" ' Replace "YourStyleName" with the actual style name
    
    ' Set the range to search in the entire document
    Set rng = doc.Content
    
    ' Perform the find operation
    found = True
    Do While found
        rng.Find.ClearFormatting
        rng.Find.style = doc.Styles(styleName)
        
        If rng.Find.Execute Then
            ' Expand the range to cover the entire paragraph
            Set paraRange = rng.Paragraphs(1).range
            
            ' Exclude the paragraph mark
            paraRange.MoveEnd wdCharacter, -1
            
            ' Select the entire paragraph
            paraRange.Select
            
            ' Check for "vol" style within the selected paragraph
            volFound = False
            For Each p In paraRange.words
                If p.style = doc.Styles("vol") Then
                    volFound = True
                    Exit For
                End If
            Next p
            
            ' Check for "firstpage" style within the selected paragraph
            firstPageFound = False
            For Each p In paraRange.words
                If p.style = doc.Styles("first-page") Then
                    firstPageFound = True
                    Exit For
                End If
            Next p
            
            ' Check for "lastpage" style within the selected paragraph
            lastPageFound = False
            For Each p In paraRange.words
                If p.style = doc.Styles("last-page") Then
                    lastPageFound = True
                    Exit For
                End If
            Next p
            
            ' Check for "author" style within the selected paragraph
            If Not volFound Or Not firstPageFound Or Not lastPageFound Then
                Dim remainingWords As String
                Dim wordCount As Integer
                remainingWords = ""
                For Each p In paraRange.words
                    If p.style = doc.Styles("author") Then
                        ' Remove all-capitalized words
                        If Not IsAllCaps(p.text) Then
                            remainingWords = remainingWords & p.text & " "
                        End If
                    End If
                Next p
                Year = ""
                For Each p In paraRange.words
                    If p.style = doc.Styles("adate") Then
                            Year = ", " & p.text & "."
                    End If
                Next p
                
                ' Count the number of remaining words
                wordCount = UBound(Split(Trim(remainingWords))) + 1
                
                ' Manipulate based on word count
                If wordCount = 2 Then
                    Dim wordsArray() As String
                    wordsArray = Split(Trim(remainingWords))
                    authorText = wordsArray(0) & " & " & wordsArray(1)
                    rng.Collapse Direction:=wdCollapseEnd
                ElseIf wordCount > 2 Then
                    Dim firstWord As String
                    firstWord = Split(Trim(remainingWords))(0)
                    authorText = firstWord & " et al."
                End If
                
                ' Manipulate paragraphs based on found styles
                If Not volFound Then
                paraRange.Comments.Add range:=paraRange, text:="[AQ: Please provide volume number and page range for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                ElseIf volFound And Not firstPageFound Then
                    ' Insert "no firstpage" if "firstpage" is not found
                    paraRange.Comments.Add range:=paraRange, text:="[AQ: Please provide complete page range for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
                ElseIf volFound And firstPageFound And Not lastPageFound Then
                    ' Insert "no lastpage" if "lastpage" is not found
                    paraRange.Comments.Add range:=paraRange, text:="[AQ: Please provide last page for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
                End If
            End If
        Else
            ' Exit the loop if no more instances are found
            found = False
        End If
    Loop
End Sub
Sub Bookcheck()
    Dim doc As Document
    Dim rng As range
    Dim found As Boolean
    Dim styleName As String
    Dim paraRange As range
    Dim volFound As Boolean
    Dim firstPageFound As Boolean
    Dim lastPageFound As Boolean
    Dim authorText As String
    Dim Year As String
    Dim p As Variant ' Define p as Variant type
    Dim leftQuote As String
    Dim rightQuote As String
    
    ' Unicode values for left and right curly quotes
    leftQuote = ChrW(8220)
    rightQuote = ChrW(8221)
    
    ' Set reference to the active document
    Set doc = ActiveDocument
    
    ' Set the style name to search for
    styleName = "btl" ' Replace "YourStyleName" with the actual style name
    
    ' Set the range to search in the entire document
    Set rng = doc.Content
    
    ' Perform the find operation
    found = True
    Do While found
        rng.Find.ClearFormatting
        rng.Find.style = doc.Styles(styleName)
        
        If rng.Find.Execute Then
            ' Expand the range to cover the entire paragraph
            Set paraRange = rng.Paragraphs(1).range
            
            ' Exclude the paragraph mark
            paraRange.MoveEnd wdCharacter, -1
            
            ' Select the entire paragraph
            paraRange.Select
            
            ' Check for "vol" style within the selected paragraph
            pubFound = False
            For Each p In paraRange.words
                If p.style = doc.Styles("pub") Then
                    pubFound = True
                    Exit For
                End If
            Next p
           
            
            ' Check for "author" style within the selected paragraph
            If Not pubFound Then
                Dim remainingWords As String
                Dim wordCount As Integer
                remainingWords = ""
                For Each p In paraRange.words
                    If p.style = doc.Styles("author") Then
                        ' Remove all-capitalized words
                        If Not IsAllCaps(p.text) Then
                            remainingWords = remainingWords & p.text & " "
                        End If
                    End If
                Next p
                Year = ""
                For Each p In paraRange.words
                    If p.style = doc.Styles("adate") Then
                            Year = ", " & p.text & "."
                    End If
                Next p
                
                ' Count the number of remaining words
                wordCount = UBound(Split(Trim(remainingWords))) + 1
                
                ' Manipulate based on word count
                If wordCount = 2 Then
                    Dim wordsArray() As String
                    wordsArray = Split(Trim(remainingWords))
                    authorText = wordsArray(0) & " & " & wordsArray(1)
                    rng.Collapse Direction:=wdCollapseEnd
                ElseIf wordCount > 2 Then
                    Dim firstWord As String
                    firstWord = Split(Trim(remainingWords))(0)
                    authorText = firstWord & " et al."
                End If
                
                ' Manipulate paragraphs based on found styles
                If Not pubFound Then
                    paraRange.Comments.Add range:=paraRange, text:="[AQ: Please provide publisher details for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                End If
            End If
        Else
            ' Exit the loop if no more instances are found
            found = False
        End If
    Loop
End Sub
Sub CountWordsInsideQuotes()
    Dim doc As Document
    Dim rng As range
    Dim quoteStart As Long
    Dim quoteEnd As Long
    Dim quoteText As String
    Dim wordCount As Integer
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Set the initial range to the beginning of the document
    Set rng = doc.range
    rng.Start = 0
    
    ' Loop through the document
    Do While rng.Find.Execute(findText:="“*”", MatchWildcards:=True) = True
        ' Get the start and end positions of the quotes
        quoteStart = rng.Start
        quoteEnd = rng.End
        
        ' Extract the text within the quotes
        quoteText = Mid(rng.text, 2, Len(rng.text) - 2)
        
        ' Count words inside the quotes
        wordCount = Len(Trim(quoteText) & " ") - Len(Replace(Trim(quoteText), " ", "")) + 1
        
        ' Display the found text and word count
        MsgBox "Text: " & quoteText & vbNewLine & "Word Count: " & wordCount
        rng.HighlightColorIndex = wdBrightGreen
        
        ' Move the range to continue searching from the end of the current quote
        rng.Start = quoteEnd
        rng.End = doc.range.End
    Loop
End Sub
Sub CountWordsInsideQuotesHightlight()
    Dim doc As Document
    Dim rng As range
    Dim quoteStart As Long
    Dim quoteEnd As Long
    Dim quoteText As String
    Dim wordCount As Integer
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Set the initial range to the beginning of the document
    Set rng = doc.range
    rng.Start = 0
    
    ' Loop through the document
    Do While rng.Find.Execute(findText:="“*”", MatchWildcards:=True) = True
        ' Get the start and end positions of the quotes
        quoteStart = rng.Start
        quoteEnd = rng.End
        
        ' Check if the range has text
        If rng.text <> "" Then
            ' Extract the text within the quotes
            quoteText = Mid(rng.text, 2, Len(rng.text) - 2)
            
            ' Count words inside the quotes
            wordCount = Len(Trim(quoteText) & " ") - Len(Replace(Trim(quoteText), " ", "")) + 1
            
            If wordCount > 40 Then
                ' Display the found text and word count
                'MsgBox "Text: " & quoteText & vbNewLine & "Word Count: " & wordCount
                rng.HighlightColorIndex = wdBrightGreen
            End If
        End If
        
        ' Move the range to continue searching from the end of the current quote
        rng.Start = quoteEnd
        rng.End = doc.range.End
    Loop
End Sub

