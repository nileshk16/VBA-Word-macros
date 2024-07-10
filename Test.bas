Attribute VB_Name = "Test"
Sub FindBullet()
    Dim rngTarget As word.range
    Dim oPara As word.paragraph

    Set rngTarget = Selection.range
    With rngTarget
        Call .Collapse(wdCollapseEnd)
        .End = ActiveDocument.range.End

        For Each oPara In .Paragraphs
            If oPara.range.ListFormat.ListType = _
               WdListType.wdListBullet Then
                oPara.range.Select
                Exit For
            End If
        Next
    End With
End Sub
Sub APAFindAndHighlightPattern(pattern As String, messageText As String)
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim rng As range
    Set rng = doc.Content
    
    Dim found As Boolean
    found = False
    
    With rng.Find
        .ClearFormatting
        .text = pattern
        .Format = False
        .MatchWildcards = True
        
        Do While .Execute
            rng.HighlightColorIndex = wdRed ' Highlight the found text
            rng.Collapse wdCollapseEnd ' Move to the end of the found range
            found = True
        Loop
    End With
    
    If found Then
        MsgBox messageText & " Pattern Found and highlighted.", vbInformation, "Result"
    Else
        MsgBox messageText & " Pattern not found.", vbExclamation, "Result"
    End If
End Sub
Sub CombinedPatternCheck()
    FindAndHighlightPattern "[0-9]{4}, ", "^#^#^#^#"
    FindAndHighlightPattern "[A-Za-z] [0-9]{4}", "^$ ^#^#^#^#"
    FindAndHighlightPattern "et al. [0-9]", "et al. [any digit]"
End Sub

Sub FindAndHighlightPattern(pattern As String, messageText As String)
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim rng As range
    Set rng = doc.Content
    
    Dim found As Boolean
    found = False
    
    With rng.Find
        .ClearFormatting
        .text = pattern
        .Format = False
        .MatchWildcards = True
        
        Do While .Execute
            rng.HighlightColorIndex = wdRed ' Highlight the found text
            rng.Collapse wdCollapseEnd ' Move to the end of the found range
            found = True
        Loop
    End With
    
    If found Then
        MsgBox messageText & " Pattern Found and highlighted.", vbInformation, "Result"
    Else
        MsgBox messageText & " Pattern not found.", vbExclamation, "Result"
    End If
End Sub
Sub Numberedchecklist()
    Dim findList() As Variant
    Dim replaceList() As Variant
    Dim i As Long
    Dim count As Long
    
    ' Characters and phrases to find
    findList = Array("et al. ", "dsm", "ibid", "in press", "n.d.", "forthcoming", "under review", "blind", "this issue", "this volume", "personal communication", "Suppl", "[0-9]{4}, ")
    
    ' Colors to highlight
    replaceList = Array(wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed)
    
    ' Loop through the find list and highlight the matches
    For i = 0 To UBound(findList)
        count = 0 ' Reset count for each character or phrase
        
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = findList(i)
            .Replacement.text = findList(i)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            ' Highlight the matches with the corresponding color and count the number of items found
            Do While .Execute
                Selection.range.HighlightColorIndex = replaceList(i)
                count = count + 1
            Loop
        End With
        
        ' Display the number of items found for each character or phrase in a message box
        ' Display the number of items found for each character or phrase in a custom message box
        response = MsgBox(count & " items found for character or phrase: " & findList(i) & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Highlighting Progress")
        
        ' Check the response, and exit the loop if the user clicks "No"
        If response = vbNo Then
            Exit For
        End If
    Next i
End Sub
Sub NumberedREFchecklist()
    Dim findList() As Variant
    Dim replaceList() As Variant
    Dim i As Long
    Dim count As Long
    
    ' Characters and phrases to find
    findList = Array("  ", "..", ". .", ". ,", ", ,", "?.", ");", "((", "))", "( ", " )", " ;", "et al. ", "doi: ", ".^p")
    
    ' Colors to highlight
    replaceList = Array(wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed)
    
    ' Loop through the find list and highlight the matches
    For i = 0 To UBound(findList)
        count = 0 ' Reset count for each character or phrase
        
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = findList(i)
            .style = "REF"
            .Replacement.text = findList(i)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            ' Highlight the matches with the corresponding color and count the number of items found
            Do While .Execute
                Selection.range.HighlightColorIndex = replaceList(i)
                count = count + 1
            Loop
        End With
        
        ' Display the number of items found for each character or phrase in a message box
        ' Display the number of items found for each character or phrase in a custom message box
        response = MsgBox(count & " items found for character or phrase: " & findList(i) & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Highlighting Progress")
        
        ' Check the response, and exit the loop if the user clicks "No"
        If response = vbNo Then
            Exit For
        End If
    Next i
End Sub
Sub AMACombinedPatternCheck()
    AMAFindAndHighlightPattern "[A-Za-z]^p ", "entermark"
    AMAFindAndHighlightPattern "; [0-9]", ": [any digit]"
    AMAFindAndHighlightPattern ": [0-9]", ": [any digit]"
End Sub

Sub AMAFindAndHighlightPattern(pattern As String, messageText As String)
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim rng As range
    Set rng = doc.Content
    
    Dim found As Boolean
    found = False
    
    With rng.Find
        .ClearFormatting
        .text = pattern
        .Format = False
        '.MatchWildcards = True
        
        Do While .Execute
            rng.HighlightColorIndex = wdRed ' Highlight the found text
            rng.Collapse wdCollapseEnd ' Move to the end of the found range
            found = True
        Loop
    End With
    
    If found Then
        MsgBox messageText & " Pattern Found and highlighted.", vbInformation, "Result"
    Else
        MsgBox messageText & " Pattern not found.", vbExclamation, "Result"
    End If
End Sub
Sub EXWords401()
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
        
        ' Check if the range has text
        If rng.text <> "" Then
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
        End If
        
        ' Move the range to continue searching from the end of the current quote
        rng.Start = quoteEnd + 1
        rng.End = doc.Content.End
    Loop
End Sub
Sub EXWords401single()
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
    
    Do While rng.Find.Execute(findText:="‘*’", MatchWildcards:=True) = True
        quoteStart = rng.Start
        quoteEnd = rng.End
        
        ' Check if the range has text
        If rng.text <> "" Then
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
        End If
        
        ' Move the range to continue searching from the end of the current quote
        rng.Start = quoteEnd + 1
        rng.End = doc.Content.End
    Loop
End Sub

Sub EXWords402()
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
        
        ' Check if the range has text
        If rng.text <> "" Then
            ' Extract the text within the quotes
            quoteText = Mid(rng.text, 2, Len(rng.text) - 2)
            
            ' Count words inside the quotes
            wordCount = UBound(Split(quoteText, " ")) + 1
            
            ' Check if word count exceeds 40 and the style is "TEXT" or "TEXT IND"
            If wordCount > 40 And (rng.ParagraphFormat.style.NameLocal = "TEXT" Or rng.ParagraphFormat.style.NameLocal = "TEXT IND") Then
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
                            ' Find and delete the specific text string
                With doc.range
                    .Find.Execute findText:="[PE: More than 40 words found inside quotes.]", MatchCase:=False
                    Do While .Find.found
                        .Delete
                        .Find.Execute
                    Loop
                End With
        End If
        
        ' Move the range to continue searching from the end of the current quote
        rng.Start = quoteEnd + 1
        rng.End = doc.Content.End
    Loop
End Sub

Sub highlightWords()

  Dim range As range
  'targetlist = Array("target1", "target2", "target3") ' put list of terms to find here
  targetlist = Array("  ", "..", ". .", ". ,", ", ,", "?.", ");", "((", "))", "( ", " )", " ;", "et al ", "et al,", ", et al", "dsm", "ibid", "in press", "n.d.", "forthcoming", "under review", "blind", "this issue", "this volume", "personal communication", "Suppl")


  For Each Target In targetlist

    Set range = ActiveDocument.range

    With range.Find
      .text = Target
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False

      Do While .Execute(Forward:=True) = True
        range.HighlightColorIndex = wdRed
      Loop

    End With
  Next

End Sub
Sub HighLightHeShe()
    Dim vFindText As Variant
    Dim oRng As range
    Dim i As Long

    vFindText = Array("he", "his")
    For i = 0 To UBound(vFindText)
        Set oRng = ActiveDocument.range
        With oRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            Do While .Execute(findText:=vFindText(i), _
                              MatchWholeWord:=True, _
                              Forward:=True, _
                              Wrap:=wdFindStop) = True
                oRng.HighlightColorIndex = wdTurquoise
                oRng.Collapse wdCollapseEnd
            Loop
        End With
    Next

    vFindText = Array("she", "her")
    For i = 0 To UBound(vFindText)
        Set oRng = ActiveDocument.range
        With oRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            Do While .Execute(findText:=vFindText(i), _
                              MatchWholeWord:=True, _
                              Forward:=True, _
                              Wrap:=wdFindStop) = True
                oRng.HighlightColorIndex = wdPink
                oRng.Collapse wdCollapseEnd
            Loop
        End With
    Next

lbl_Exit:
    Exit Sub
End Sub

Sub highlightWords11()
    Dim range As range
    Dim Target As Variant
    Dim occurrences As Integer

    'targetlist = Array("target1", "target2", "target3") ' put list of terms to find here
    targetlist = Array("  ", "..", ". .", ". ,", ", ,", "?.", ");", "((", "))", "( ", " )", " ;", "et al ", "et al,", ", et al", "dsm", "ibid", "in press", "n.d.", "forthcoming", "under review", "blind", "this issue", "this volume", "personal communication", "Suppl")

    For Each Target In targetlist
        Set range = ActiveDocument.range
        occurrences = 0

        With range.Find
            .text = Target
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            Do While .Execute(Forward:=True) = True
                range.HighlightColorIndex = wdRed
                occurrences = occurrences + 1
            Loop
        End With

        ' Display the number of occurrences for each target
        MsgBox "Number of occurrences of '" & Target & "': " & occurrences, vbInformation
    Next
End Sub

Sub highlightWords12()
    Dim range As range
    Dim Target As Variant
    Dim occurrences As Integer
    Dim summaryMessage As String

    'targetlist = Array("target1", "target2", "target3") ' put list of terms to find here
    targetlist = Array("  ", "..", ". .", ". ,", ", ,", "?.", ");", "((", "))", "( ", " )", " ;", "et al ", "et al,", ", et al", "dsm", "ibid", "in press", "n.d.", "forthcoming", "under review", "blind", "this issue", "this volume", "personal communication", "Suppl")

    For Each Target In targetlist
        Set range = ActiveDocument.range
        occurrences = 0

        With range.Find
            .text = Target
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            Do While .Execute(Forward:=True) = True
                range.HighlightColorIndex = wdRed
                occurrences = occurrences + 1
            Loop
        End With

        ' Accumulate information for summary
        summaryMessage = summaryMessage & "Number of occurrences of '" & Target & "': " & occurrences & vbCrLf
    Next

    ' Display the summary message box
    MsgBox "Summary of items found:" & vbCrLf & summaryMessage, vbInformation
End Sub
Sub highlightWords13()
    Dim range As range
    Dim Target As Variant
    Dim occurrences As Integer
    Dim summaryMessage As String

    'targetlist = Array("target1", "target2", "target3") ' put list of terms to find here
    targetlist = Array("  ", "..", ". .", ". ,", ", ,", "?.", ");", "((", "))", "( ", " )", " ;", "et al ", "et al,", ", et al", "dsm", "ibid", "in press", "n.d.", "forthcoming", "under review", "blind", "this issue", "this volume", "personal communication", "Suppl")

    For Each Target In targetlist
        Set range = ActiveDocument.range
        occurrences = 0

        With range.Find
            .text = Target
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            Do While .Execute(Forward:=True) = True
                range.HighlightColorIndex = wdRed
                occurrences = occurrences + 1
            Loop
        End With

        ' Accumulate information for summary if occurrences > 0
        If occurrences > 0 Then
            summaryMessage = summaryMessage & "Number of occurrences of '" & Target & "': " & occurrences & vbCrLf
        End If
    Next

    ' Display the summary message box if there are any found occurrences
    If Len(summaryMessage) > 0 Then
        MsgBox "Summary of items found:" & vbCrLf & summaryMessage, vbInformation
    End If
End Sub


Sub AbbreviateJournals()
    Dim ws As Worksheet
    Dim rng As range
    Dim i As Integer
    
    ' Set the worksheet with your data
    Set ws = ThisWorkbook.Sheets("E:\Nilesh\Journal database\Medical journal database.xlsx") ' Change YourSheetName to the actual sheet name
    
    ' Set the range with your data (assuming it starts from the second row)
    Set rng = ws.range("A2:B" & ws.Cells(ws.Rows.count, "A").End(xlUp).Row)
    
    ' Loop through each row in the range
    For i = 1 To rng.Rows.count
        ' Replace full journal names with abbreviations in Word document
        With ActiveDocument.Content.Find
            .text = rng.Cells(i, 1).Value
            .Replacement.text = rng.Cells(i, 2).Value
            .Execute Replace:=wdReplaceAll
        End With
    Next i
End Sub

