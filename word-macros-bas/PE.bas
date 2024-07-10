Attribute VB_Name = "PE"
Function IsAllCaps(ByVal text As String) As Boolean
    IsAllCaps = text Like UCase(text)
End Function
Sub PEjournalcheck()
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
                    ' Insert "no vol" if "vol" is not found along with author text
                        paraRange.Collapse Direction:=wdCollapseEnd
                        'paraRange.text = "[AQ: Please provide volume number and page range for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
                        With paraRange
                            .text = "[AQ: Please provide volume number and page range for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide volume number and page range for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                    rng.Collapse Direction:=wdCollapseEnd
                ElseIf volFound And Not firstPageFound Then
                    ' Insert "no firstpage" if "firstpage" is not found
                    paraRange.Collapse Direction:=wdCollapseEnd
                    'paraRange.text = "[AQ: Please provide complete page range for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
                        With paraRange
                            .text = "[AQ: Please provide complete page range for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide complete page range for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                    rng.Collapse Direction:=wdCollapseEnd
                ElseIf volFound And firstPageFound And Not lastPageFound Then
                    ' Insert "no lastpage" if "lastpage" is not found
                    paraRange.Collapse Direction:=wdCollapseEnd
                    'paraRange.text = "[AQ: Please provide last page for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
                        With paraRange
                            .text = "[AQ: Please provide last page for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide last page for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                    rng.Collapse Direction:=wdCollapseEnd
                End If
            End If
        Else
            ' Exit the loop if no more instances are found
            found = False
        End If
    Loop
End Sub
Sub PEBookcheck()
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
                    ' Insert "no vol" if "vol" is not found along with author text
                        paraRange.Collapse Direction:=wdCollapseEnd
                        'paraRange.text = "[AQ: Please provide volume number and page range for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
                        With paraRange
                            .text = "[AQ: Please provide publisher details for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide publisher details for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                    rng.Collapse Direction:=wdCollapseEnd
                End If
            End If
        Else
            ' Exit the loop if no more instances are found
            found = False
        End If
    Loop
End Sub
Sub PEjournalcheckrev1()
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

            ' Check for "doi" style within the selected paragraph
            doiFound = False
            For Each p In paraRange.words
                If p.style = doc.Styles("url") Then
                    doiFound = True
                    Exit For
                End If
            Next p
            
            ' Check for "author" style within the selected paragraph
            If Not volFound Or Not firstPageFound Or Not lastPageFound Or Not doiFound Then
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
                If Not volFound And Not doiFound Then
                    ' Insert "no vol" if "vol" is not found along with author text
                        paraRange.Collapse Direction:=wdCollapseEnd
                        'paraRange.text = "[AQ: Please provide volume number and page range for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
                        With paraRange
                            .text = "[AQ: Please provide volume number and page range for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide volume number and page range for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                    rng.Collapse Direction:=wdCollapseEnd
                ElseIf Not volFound And doiFound Then
                    'Insert "no lastpage" if "lastpage" is not found
                    paraRange.Collapse Direction:=wdCollapseEnd
                    'paraRange.text = "Advance online publication"
                        With paraRange
                            .text = "[AQ: Please provide volume number and page range for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide volume number and page range for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                    rng.Collapse Direction:=wdCollapseEnd
                ElseIf volFound And Not firstPageFound And doiFound Then
                    ' Insert "no firstpage" if "firstpage" is not found
                    paraRange.Collapse Direction:=wdCollapseEnd
                    'paraRange.text = "[AQ: Please provide complete page range for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
                        With paraRange
                            .text = "[AQ: Please provide page range for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide page range for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                    rng.Collapse Direction:=wdCollapseEnd
                ElseIf volFound And firstPageFound And Not lastPageFound Then
                    ' Insert "no lastpage" if "lastpage" is not found
                    paraRange.Collapse Direction:=wdCollapseEnd
                    'paraRange.text = "[AQ: Please provide last page for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
                        With paraRange
                            .text = "[AQ: Please provide complete page range for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide complete page range for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                    rng.Collapse Direction:=wdCollapseEnd
                ElseIf volFound And Not firstPageFound And Not lastPageFound And Not doiFound Then
                    ' Insert "no lastpage" if "lastpage" is not found
                    paraRange.Collapse Direction:=wdCollapseEnd
                    'paraRange.text = "[AQ: Please provide last page for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
                        With paraRange
                            .text = "[AQ: Please provide the page range, or, if the page range is unavailable, please provide the URL or DOI the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide the page range, or, if the page range is unavailable, please provide the URL or DOI the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                    rng.Collapse Direction:=wdCollapseEnd

                End If
            End If
        Else
            ' Exit the loop if no more instances are found
            found = False
        End If
    Loop
End Sub
Sub ReplaceJournalNames()
    Dim doc As Document
    Dim rng As range
    Dim findText As Variant
    Dim replaceText As String
    
    ' Set the target journal names and replacement text
    findText = Array("Proceedings of the National Academy of Sciences", _
                     "Proceedings of the National Academy of Sciences of the United States of America", _
                     "Proceedings of the National Academy of Sciences of USA")
    
    replaceText = "Proceedings of the National Academy of Sciences, USA"
    
    ' Check if a document is open in Word
    If Documents.count = 0 Then
        MsgBox "Please open a document.", vbExclamation
        Exit Sub
    End If
    
    ' Set reference to the active document
    Set doc = ActiveDocument
    
    ' Loop through each target text in the array
    For Each text In findText
        ' Set the range to search the entire document
        Set rng = doc.Content
        
        ' Loop through the document to find and replace specified journal names in "stl" style
        With rng.Find
            .ClearFormatting
            .style = "stl"
            .text = text
            .Replacement.text = replaceText
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False ' Disable wildcards to match exact text
            .Execute Replace:=wdReplaceAll
        End With
    Next text
    
    ' Notify user upon completion
    MsgBox "Replacement complete.", vbInformation
End Sub
Sub PEplosone()
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
        rng.Find.text = "plos one"
        
        If rng.Find.Execute Then
            ' Expand the range to cover the entire paragraph
            Set paraRange = rng.Paragraphs(1).range
            
            ' Exclude the paragraph mark
            paraRange.MoveEnd wdCharacter, -1
            
            ' Select the entire paragraph
            paraRange.Select
            
            ' Check for "url" style within the selected paragraph
            urlFound = False
            For Each p In paraRange.words
                If p.style = doc.Styles("url") Then
                    urlFound = True
                    Exit For
                End If
            Next p

            ' Check for "doi" style within the selected paragraph
            doiFound = False
            For Each p In paraRange.words
                If p.style = doc.Styles("doino") Then
                    doiFound = True
                    Exit For
                End If
            Next p
            
            ' Check for "author" style within the selected paragraph
            If Not doiFound Or Not urlFound Then
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
                If Not doiFound Or Not urlFound Then
                        paraRange.Collapse Direction:=wdCollapseEnd
                        With paraRange
                            .text = "[AQ: Please provide DOI number for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide DOI number for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                    rng.Collapse Direction:=wdCollapseEnd
                End If
            End If
        Else
            ' Exit the loop if no more instances are found
            found = False
        End If
    Loop
End Sub
Sub PEdoibasedonstyle()
    Dim doc As Document
    Dim rng As range
    Dim found As Boolean
    Dim styleName As String
    Dim paraRange As range
    Dim doiFound As Boolean
    Dim urlFound As Boolean
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
    
    ' Array of target texts to search for
    Dim targetTexts As Variant
    targetTexts = Array("plos one", "frontiers", "Journal of Vision")
    
    ' Loop through each target text
    For Each targetText In targetTexts
        ' Set the style name to search for
        styleName = "stl" ' Replace "YourStyleName" with the actual style name
        
        ' Set the range to search in the entire document
        Set rng = doc.Content
        
        ' Perform the find operation
        found = True
        Do While found
            rng.Find.ClearFormatting
            rng.Find.style = doc.Styles(styleName)
            rng.Find.text = targetText
            
            If rng.Find.Execute Then
                ' Expand the range to cover the entire paragraph
                Set paraRange = rng.Paragraphs(1).range
                
                ' Exclude the paragraph mark
                paraRange.MoveEnd wdCharacter, -1
                
                ' Select the entire paragraph
                paraRange.Select
                
                ' Check for "url" style within the selected paragraph
                urlFound = False
                For Each p In paraRange.words
                    If p.style = doc.Styles("url") Then
                        urlFound = True
                        Exit For
                    End If
                Next p
                
                ' Check for "doi" style within the selected paragraph
                doiFound = False
                For Each p In paraRange.words
                    If p.style = doc.Styles("doi") Then
                        doiFound = True
                        Exit For
                    End If
                Next p
                
                ' Check for "author" style within the selected paragraph
                If Not doiFound Or Not urlFound Then
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
                    If Not doiFound Or Not urlFound Then
                        paraRange.Collapse Direction:=wdCollapseEnd
                        With paraRange
                            .text = "[AQ: Please provide DOI number for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide DOI number for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                        rng.Collapse Direction:=wdCollapseEnd
                    End If
                End If
            Else
                ' Exit the loop if no more instances are found
                found = False
            End If
        Loop
    Next targetText
End Sub
Sub PEdoi()
    Dim doc As Document
    Dim rng As range
    Dim found As Boolean
    Dim styleName As String
    Dim paraRange As range
    Dim doiFound As Boolean
    Dim urlFound As Boolean
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
    
    ' Array of target texts to search for
    Dim targetTexts As Variant
    targetTexts = Array("plos one", "frontiers", "Journal of Vision")
    
    ' Loop through each target text
    For Each targetText In targetTexts
        ' Set the style name to search for
        styleName = "stl" ' Replace "YourStyleName" with the actual style name
        
        ' Set the range to search in the entire document
        Set rng = doc.Content
        
        ' Perform the find operation
        found = True
        Do While found
            rng.Find.ClearFormatting
            rng.Find.style = doc.Styles(styleName)
            rng.Find.text = targetText
            
            If rng.Find.Execute Then
                ' Expand the range to cover the entire paragraph
                Set paraRange = rng.Paragraphs(1).range
                
                ' Exclude the paragraph mark
                paraRange.MoveEnd wdCharacter, -1
                
                ' Select the entire paragraph
                paraRange.Select
                
            doiFound = False
            Dim wordsArray() As String
            wordsArray = Split(paraRange.text, " ") ' Split the paragraph into words

            For i = LBound(wordsArray) To UBound(wordsArray)
                If InStr(1, wordsArray(i), "/doi.org/") > 0 Then
                    doiFound = True
                    Exit For
                End If
            Next i
                
                ' Check for "author" style within the selected paragraph
                If Not doiFound Then
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
                        'Dim wordsArray() As String
                        wordsArray = Split(Trim(remainingWords))
                        authorText = wordsArray(0) & " & " & wordsArray(1)
                        rng.Collapse Direction:=wdCollapseEnd
                    ElseIf wordCount > 2 Then
                        Dim firstWord As String
                        firstWord = Split(Trim(remainingWords))(0)
                        authorText = firstWord & " et al."
                    End If
                    
                    ' Manipulate paragraphs based on found styles
                    If Not doiFound Then
                        paraRange.Collapse Direction:=wdCollapseEnd
                        With paraRange
                            .text = "[AQ: Please provide DOI number for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide DOI number for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                        rng.Collapse Direction:=wdCollapseEnd
                    End If
                End If
            Else
                ' Exit the loop if no more instances are found
                found = False
            End If
        Loop
    Next targetText
End Sub
Sub PEdoi2()
    Dim doc As Document
    Dim rng As range
    Dim found As Boolean
    Dim styleName As String
    Dim paraRange As range
    Dim doiFound As Boolean
    Dim urlFound As Boolean
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
    
    ' Array of target texts to search for
    Dim targetTexts As Variant
    targetTexts = Array("plos one", "frontiers", "Journal of Vision")
    
    ' Loop through each target text
    For Each targetText In targetTexts
        ' Set the style name to search for
        styleName = "stl" ' Replace "YourStyleName" with the actual style name
        
        ' Set the range to search in the entire document
        Set rng = doc.Content
        
        ' Perform the find operation
        found = True
        Do While found
            rng.Find.ClearFormatting
            rng.Find.style = doc.Styles(styleName)
            rng.Find.text = targetText
            
            If rng.Find.Execute Then
                ' Expand the range to cover the entire paragraph
                Set paraRange = rng.Paragraphs(1).range
                
                ' Exclude the paragraph mark
                paraRange.MoveEnd wdCharacter, -1
                
                ' Select the entire paragraph
                paraRange.Select
                
                doiFound = False
                Dim wordsArray() As String
                wordsArray = Split(paraRange.text, " ") ' Split the paragraph into words

                For i = LBound(wordsArray) To UBound(wordsArray)
                    If InStr(1, wordsArray(i), "/doi.org/") > 0 Then
                        doiFound = True
                        Exit For
                    End If
                Next i

                            ' Manipulate paragraphs based on found styles
                    If Not doiFound Then
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
                                'Dim wordsArray() As String
                                wordsArray = Split(Trim(remainingWords))
                                authorText = wordsArray(0) & " & " & wordsArray(1)
                                rng.Collapse Direction:=wdCollapseEnd
                            ElseIf wordCount > 2 Then
                                Dim firstWord As String
                                firstWord = Split(Trim(remainingWords))(0)
                                authorText = firstWord & " et al."
                            End If
                                paraRange.Collapse Direction:=wdCollapseEnd
                                With paraRange
                                    .text = "[AQ: Please provide DOI number for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                                    .Start = .End - Len("[AQ: Please provide DOI number for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                                    .style = "AQ" ' Apply the "AQ" style
                                    .Font.Bold = True ' Set the inserted text to bold
                                End With
                                rng.Collapse Direction:=wdCollapseEnd
                            End If
                    Else
                        ' Exit the loop if no more instances are found
                        found = False
                    End If
        Loop
    Next targetText
End Sub
