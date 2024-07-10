Attribute VB_Name = "Hammil"
Sub Hammildoicheckbasedonstyle()
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
            doiFound = False
            For Each p In paraRange.words
                If p.style = doc.Styles("url") Then
                    doiFound = True
                    Exit For
                End If
            Next p
           
            
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
                If Not doiFound Then
                    ' Insert "no vol" if "vol" is not found along with author text
                        paraRange.Collapse Direction:=wdCollapseEnd
                        'paraRange.text = "[AQ: Please provide volume number and page range for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
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
Sub Hammildoicheckbasedontext()
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

        doiFound = False
        Dim wordsArray() As String
        wordsArray = Split(paraRange.text, " ") ' Split the paragraph into words

        For i = LBound(wordsArray) To UBound(wordsArray)
            If InStr(1, wordsArray(i), "/doi.org/") > 0 Or InStr(1, wordsArray(i), "/doi/org/") > 0 Then
                doiFound = True
                Exit For
            End If
        Next i
            
            ' Check for "author" style within the selected paragraph
            'If Not doiFound Then
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
                    ' Insert "no vol" if "vol" is not found along with author text
                        paraRange.Collapse Direction:=wdCollapseEnd
                        'paraRange.text = "[AQ: Please provide volume number and page range for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
                        With paraRange
                            .text = "[AQ: Please provide DOI number for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide DOI number for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                    rng.Collapse Direction:=wdCollapseEnd
                End If
            'End If
                Else
            ' Exit the loop if no more instances are found
                    found = False
        End If
    Loop
End Sub



Sub Hammildoicheckbasedontext5()
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

        doiFound = False
        Dim wordsArray() As String
        wordsArray = Split(paraRange.text, " ") ' Split the paragraph into words

        For i = LBound(wordsArray) To UBound(wordsArray)
            If InStr(1, wordsArray(i), "/doi.org/") > 0 Or InStr(1, wordsArray(i), "/doi/org/") > 0 Then
                doiFound = True
                Exit For
            End If
        Next i
            
            ' Check for "author" style within the selected paragraph
            'If Not doiFound Then
                Dim remainingWords As String
                Dim wordCount As Integer
                remainingWords = ""
                For Each p In paraRange.words
                    If p.style = doc.Styles("author") Then
                        ' Remove all-capitalized words
                        If Len(p.text) > 1 Then
                            remainingWords = remainingWords & p.text & " "
                            authorText = remainingWords
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
                
                ' Manipulate based on word count

                
                ' Manipulate paragraphs based on found styles
                If Not doiFound Then
                    ' Insert "no vol" if "vol" is not found along with author text
                        paraRange.Collapse Direction:=wdCollapseEnd
                        'paraRange.text = "[AQ: Please provide volume number and page range for the reference " & leftQuote & "" & authorText & "" & Year & "" & rightQuote & "]"
                        With paraRange
                            .text = "[AQ: Please provide DOI number for the reference " & leftQuote & authorText & Year & rightQuote & "]"
                            .Start = .End - Len("[AQ: Please provide DOI number for the reference " & leftQuote & authorText & Year & rightQuote & "]") ' Move the range start back
                            .style = "AQ" ' Apply the "AQ" style
                            .Font.Bold = True ' Set the inserted text to bold
                        End With
                    rng.Collapse Direction:=wdCollapseEnd
                End If
            'End If
                Else
            ' Exit the loop if no more instances are found
                    found = False
        End If
    Loop
End Sub
