Attribute VB_Name = "General"
Sub ReplaceQuotes()
    Dim blnQuotes As Boolean
    blnQuotes = Application.Options.AutoFormatAsYouTypeReplaceQuotes
    Application.Options.AutoFormatAsYouTypeReplaceQuotes = True
    
    ' Replace single quotes
    ActiveDocument.Content.Find.Execute findText:="'", MatchCase:=False, _
        MatchWholeWord:=False, MatchWildcards:=True, Wrap:=wdFindStop, _
        Format:=False, ReplaceWith:="'", Replace:=wdReplaceAll
    
    ' Replace double quotes
    ActiveDocument.Content.Find.Execute findText:="""", MatchCase:=False, _
        MatchWholeWord:=False, MatchWildcards:=True, Wrap:=wdFindStop, _
        Format:=False, ReplaceWith:="""", Replace:=wdReplaceAll
    
    Application.Options.AutoFormatAsYouTypeReplaceQuotes = blnQuotes
End Sub
Sub listtotext()
'
' listtotext Macro
ActiveDocument.range.ListFormat.ConvertNumbersToText

End Sub
Sub outlineLevel()
ActiveDocument.Styles("H1").ParagraphFormat.outlineLevel = wdOutlineLevel1
ActiveDocument.Styles("H2").ParagraphFormat.outlineLevel = wdOutlineLevel2
ActiveDocument.Styles("H3").ParagraphFormat.outlineLevel = wdOutlineLevel3
ActiveDocument.Styles("H4").ParagraphFormat.outlineLevel = wdOutlineLevel4
'ActiveDocument.Styles("H5").ParagraphFormat.outlineLevel = wdOutlineLevel5
End Sub
Sub SetOutlineLevel()
Dim outlineLevel As WdOutlineLevel
Dim styleName As String
    If styleName = "H1" Then
        ActiveDocument.Styles("H1").ParagraphFormat.outlineLevel = outlineLevel
    ElseIf styleName = "H2" Then
        ActiveDocument.Styles("H2").ParagraphFormat.outlineLevel = outlineLevel
    ElseIf styleName = "H3" Then
        ActiveDocument.Styles("H3").ParagraphFormat.outlineLevel = outlineLevel
    ElseIf styleName = "H4" Then
        ActiveDocument.Styles("H4").ParagraphFormat.outlineLevel = outlineLevel
    ElseIf styleName = "H5" Then
        ActiveDocument.Styles("H5").ParagraphFormat.outlineLevel = outlineLevel
    End If
End Sub
Sub DeleteFigurePlaceholder()
    Dim doc As Document
    Dim para As paragraph
    
    Set doc = ActiveDocument
    
    ' Loop through all paragraphs in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph style is 'CL'
        If para.style = doc.Styles("CL") Then
            ' Check if the paragraph text contains the placeholder
            If InStr(para.range.text, "[FIGURE") > 0 Then
                ' Delete the paragraph
                para.range.Delete
            End If
        End If
    Next para
End Sub
Sub FormatTable()
    Dim doc As Document
    Dim tbl As TABLE
    
    ' Set a reference to the active document
    Set doc = ActiveDocument
    
    ' Loop through all tables in the document
    For Each tbl In doc.Tables
        ' Change the table width to 20 centimeters
        tbl.PreferredWidthType = wdPreferredWidthPoints
        tbl.PreferredWidth = CentimetersToPoints(20)
        
        tbl.Rows.Alignment = wdAlignRowLeft
        
        ' Set the table shading to "No Color"
        tbl.Shading.BackgroundPatternColor = wdColorAutomatic
    Next tbl
End Sub
Sub SortWordsAlphabetically()
    Dim selectedText As range
    Dim wordsArray() As String
    Dim sortedWords() As String
    Dim i As Integer
    
    ' Check if text is selected
    If Selection.Type <> wdSelectionIP Then
        ' Get the selected text
        Set selectedText = Selection.range
        ' Split the selected text into an array of words
        wordsArray = Split(selectedText.text, ",")
        
        ' Sort the words alphabetically
        For i = 0 To UBound(wordsArray)
            wordsArray(i) = Trim(wordsArray(i))
        Next i
        
        ' Sort the array
        QuickSort wordsArray, 0, UBound(wordsArray)
        
        ' Join the sorted words into a string
        sortedWords = wordsArray
        selectedText.text = Join(sortedWords, ", ")
    End If
End Sub
Sub QuickSort(arr() As String, low As Integer, high As Integer)
    Dim i As Integer, j As Integer
    Dim pivot As String, temp As String
    
    i = low
    j = high
    pivot = arr((low + high) \ 2)
    
    While i <= j
        While StrComp(arr(i), pivot, vbTextCompare) < 0
            i = i + 1
        Wend
        
        While StrComp(arr(j), pivot, vbTextCompare) > 0
            j = j - 1
        Wend
        
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Wend
    
    If low < j Then QuickSort arr, low, j
    If i < high Then QuickSort arr, i, high
End Sub
Sub SelectKeywords()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim keywordsRange As range
    Set keywordsRange = doc.Content
    
    ' Find the word "Keywords"
    With keywordsRange.Find
        .text = "Keywords"
        .style = "ABKWH"
        .Execute
    End With
    
    ' Check if the word "Keywords" is found
    If keywordsRange.Find.found Then
        ' Move to the next paragraph
        keywordsRange.Move Unit:=wdParagraph, count:=1
        
        ' Select the paragraph without the paragraph mark
        Dim paragraphRange As range
        Set paragraphRange = keywordsRange.Paragraphs(1).range
        paragraphRange.MoveEnd Unit:=wdCharacter, count:=-1 ' Exclude paragraph mark
        paragraphRange.style = "ABKW"
        paragraphRange.Select
    Else
        MsgBox "Word 'Keywords' not found."
    End If
End Sub
Sub Alpha_order_keywords()

Call SelectKeywords
Call SortWordsAlphabetically
End Sub
Sub SortWordsAlphabetically2()
    Dim selectedText As range
    Dim wordsArray() As String
    Dim sortedWords() As String
    Dim i As Integer
    
    ' Check if text is selected
    If Selection.Type <> wdSelectionIP Then
        ' Get the selected text
        Set selectedText = Selection.range
        ' Split the selected text into an array of words
        wordsArray = Split(selectedText.text, ";")
        
        ' Sort the words alphabetically
        For i = 0 To UBound(wordsArray)
            wordsArray(i) = Trim(wordsArray(i))
        Next i
        
        ' Sort the array
        QuickSort wordsArray, 0, UBound(wordsArray)
        
        ' Join the sorted words into a string
        sortedWords = wordsArray
        selectedText.text = Join(sortedWords, "; ")
    End If
End Sub
Sub EXWords40()
    Dim doc As Document
    Dim rng As range
    Dim quoteStart As Long
    Dim quoteEnd As Long
    Dim quoteText As String
    Dim wordCount As Integer
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Initialize the range to search from the beginning of the document
    Set rng = doc.Content
    rng.Start = 0
    
    Do While rng.Find.Execute(findText:="“*”", MatchWildcards:=True) = True
        quoteStart = rng.Start
        quoteEnd = rng.End
        
        ' Extract the text within the quotes
        quoteText = Mid(rng.text, 2, Len(rng.text) - 2)
        
        ' Count words inside the quotes
        wordCount = UBound(Split(quoteText, " ")) + 1
        
        ' Check if word count exceeds 40
        If wordCount > 40 Then
            ' Insert line breaks before and after the quote
            rng.MoveStartUntil Chr(10), wdForward ' Move to the beginning of the line
            rng.InsertBefore vbNewLine ' Insert a line break before the opening quote
            rng.MoveEndUntil Chr(10), wdForward ' Move to the end of the line
            rng.InsertAfter vbNewLine ' Insert a line break after the ending quote
            
            ' Apply style and formatting to the text within the quotes
            Set rng = doc.range(Start:=quoteStart + Len(vbNewLine), End:=quoteEnd + Len(vbNewLine))
            rng.style = "EX"
            rng.HighlightColorIndex = wdBrightGreen
        End If
        
        ' Move the range to continue searching from the end of the current quote
        rng.Start = quoteEnd + 1
        rng.End = doc.Content.End
    Loop
End Sub

Sub ReplaceMultipleSpaces()
    With Selection.Find
        .ClearFormatting
        .text = "( ){2,}"
        .Replacement.ClearFormatting
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Sub Replacetabswithspaces()
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.text = "^t"
.style = "REF"
.Replacement.text = " "
.Forward = True
.Wrap = wdFindContinue
.Format = False
.MatchCase = False
.MatchWholeWord = False
.MatchByte = False
.MatchAllWordForms = False
.MatchSoundsLike = False
.MatchWildcards = False
.MatchFuzzy = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub GetWordCount()
    Dim wordCount As Long
    
    ' Get the word count
    wordCount = ActiveDocument.ComputeStatistics(wdStatisticWords)
    
    ' Display the word count
    MsgBox "The document contains " & wordCount & " words.", vbInformation, "Word Count"
End Sub
Sub Pagecount()
    Dim wordCount As Long
    Dim pages As Double
    
    ' Get the word count
    wordCount = ActiveDocument.ComputeStatistics(wdStatisticWords)
    
    ' Calculate pages based on word count (assuming 350 words per page)
    pages = wordCount / 350
    
    ' Display the result
    MsgBox "Word count: " & wordCount & vbNewLine & "Page count: " & Format(pages, "#,##0.00") & " pages", vbInformation, "Word Count and Page Count"
End Sub


Sub InsertWordCount()

    Selection.text = ActiveDocument.ComputeStatistics(wdStatisticWords)
    Selection.Collapse (wdCollapseEnd) ' Prevent overwriting

End Sub
