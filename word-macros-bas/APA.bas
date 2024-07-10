Attribute VB_Name = "APA"
'APA 7
Sub USAPA()
Dim answer As String
    
    ' Prompt the user for input using a message box
answer = MsgBox("Please compare the head levels with MS before running the macro. Do you want to continue?", vbYesNo, "Continue or Stop")
    
    ' Check the user's response
If answer = vbYes Then
        ' Continue with the macro
    MsgBox "Continuing with the macro..."
        ' Add your code here for the continuation

'Trun on track changes and turn off trcak moves
Call TurnOnTrackChanges
'General points
Call listtotext
Call ReplaceQuotes

'APA Front matter

Call APARunninghead



'Formating
Call APAFormating
Call APATableWords
'Removes endperiods
Call APARemoveEndPeriod

'Delete TS insert query
Call DeleteFigurePlaceholder


'Title case
'Call TitleCaseAPA

Call TitleCaseAPAMicroUS102

'Check Head level
Call outlineLevel
Call CheckHeadlevels2
Call CheckHeadlevels3

'Check Table first column
Call TableEmptyCells
'Format Table
Call FormatTable
Call AQcolumn

'Call EXWords40
'Call ReplaceMultipleSpaces

MsgBox "APA style formatted successfully. Please check and confirm."
Else
        ' Stop the macro
        MsgBox "Stopping the macro."
        Exit Sub ' Exit the subroutine if the user chooses "No"
End If
End Sub
Sub APARunninghead()
    Dim doc As Document
    Dim para As paragraph
    Set doc = ActiveDocument
    'LRH, RRH
    For Each para In doc.Paragraphs
        If para.style = "LRH" Then
            Set rng = para.range
            rng.MoveEnd wdCharacter, -1 ' Move the end of the range before the paragraph mark
            rng.text = rng.text & " XX(X)"
        End If
    Next para
    
    For Each para In doc.Paragraphs
        If para.style = "LRH" Or para.style = "RRH" Then
            para.range.Font.Italic = True
        End If
    Next para
End Sub
Sub APAFormating()
    Dim doc As Document
    Dim para As paragraph
    
    Set doc = ActiveDocument
    
    For Each para In doc.Paragraphs
        Select Case para.style
            Case "H1"
                With para.range
                    .Font.Bold = True
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                End With
            Case "H2", "H4", "CPB"
                para.range.Font.Bold = True
            Case "H3"
                With para.range
                    .Font.Bold = True
                    .Font.Italic = True
                End With
            Case "CP"
                para.range.Font.Italic = True
        End Select
    Next para
End Sub
Sub APARemoveEndPeriod()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim para As paragraph
    Dim styleName As Variant
    
    For Each para In doc.Paragraphs
        styleName = para.style
        
        If styleName = "CP" Or styleName = "CPB" Or styleName = "H3" Then
            Dim paraText As String
            paraText = para.range.text
            
            If Right(paraText, 2) = "." & vbCr Then
                para.range.Characters(Len(paraText) - 1).Delete
            End If
        End If
    Next para
End Sub
Sub CheckHeadlevels2()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim para As paragraph
    Dim level1Name As String
    Dim level2Name As String
    Dim hdg1Count As Integer
    Dim hdg2Count As Integer
    Dim insertionRange As range
    Dim previousPara As paragraph
    
    hdg1Count = 0
    hdg2Count = 0
    
    For Each para In doc.Paragraphs
        If para.style <> "" Then
            If para.outlineLevel = wdOutlineLevel1 Then
                ' Check Heading Level 1 count and Heading Level 2 count
                If hdg1Count > 0 And hdg2Count = 1 Then
                    Set insertionRange = previousPara.range
                    insertionRange.Collapse wdCollapseEnd
                    insertionRange.InsertAfter vbCrLf & "[AQ: There is only one H2 “" & level2Name & "” given under the H1 “" & level1Name & "”. Please consider adding another H2 in this section or allow us to delete the heading “" & level2Name & "”, as APA style requires at least two subheadings under each heading level.]" & vbCrLf
                    insertionRange.Font.Bold = True
                    insertionRange.style = "AQ" ' Apply style "AQ" to the inserted text
                End If
                level1Name = Left(para.range.text, Len(para.range.text) - 1)
                hdg1Count = hdg1Count + 1
                hdg2Count = 0
            ElseIf para.outlineLevel = wdOutlineLevel2 Then
                ' Increment Heading Level 2 count and remove the last character (line break)
                level2Name = Left(para.range.text, Len(para.range.text) - 1)
                hdg2Count = hdg2Count + 1
            End If
        End If
        Set previousPara = para ' Store the previous paragraph
    Next para
    
    ' Check for the last heading group
    If hdg1Count > 0 And hdg2Count = 1 Then
        Set insertionRange = previousPara.range
        insertionRange.Collapse wdCollapseEnd
        insertionRange.InsertAfter vbCrLf & "[AQ: There is only one H2 """ & level2Name & """ given under the H1 """ & level1Name & """. Please consider adding another H2 in this section or allow us to delete the heading """ & level2Name & """, as APA style requires at least two subheadings under each heading level.]" & vbCrLf
        insertionRange.Font.Bold = True
        insertionRange.style = "AQ" ' Apply style "AQ" to the inserted text
    End If
End Sub
Sub CheckHeadlevels3()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim para As paragraph
    Dim level2Name As String
    Dim level3Name As String
    Dim hdg2Count As Integer
    Dim hdg3Count As Integer
    Dim insertionRange As range
    Dim previousPara As paragraph
    
    hdg2Count = 0
    hdg3Count = 0
    
    For Each para In doc.Paragraphs
        If para.style <> "" Then
            If para.outlineLevel = wdOutlineLevel2 Then
                ' Check Heading Level 2 count and Heading Level 3 count
                If hdg2Count > 0 And hdg3Count = 2 Then
                    Set insertionRange = previousPara.range
                    insertionRange.Collapse wdCollapseEnd
                    insertionRange.InsertAfter vbCrLf & "[AQ: There is only one H3 “" & level3Name & "” given under the H2 “" & level2Name & "”. Please consider adding another H3 in this section or allow us to delete the heading “" & level3Name & "”, as APA style requires at least two subheadings under each heading level.]" & vbCrLf
                    insertionRange.Font.Bold = True
                    insertionRange.style = "AQ" ' Apply style "AQ" to the inserted text
                End If
                level2Name = Left(para.range.text, Len(para.range.text) - 2)
                hdg2Count = hdg2Count + 2
                hdg3Count = 0
            ElseIf para.outlineLevel = wdOutlineLevel3 Then
                ' Increment Heading Level 3 count and remove the last character (line break)
                level3Name = Left(para.range.text, Len(para.range.text) - 2)
                hdg3Count = hdg3Count + 2
            End If
        End If
        Set previousPara = para ' Store the previous paragraph
    Next para
    
    ' Check for the last heading group
    If hdg2Count > 0 And hdg3Count = 2 Then
        Set insertionRange = previousPara.range
        insertionRange.Collapse wdCollapseEnd
        insertionRange.InsertAfter vbCrLf & "[AQ: There is only one H3 """ & level3Name & """ given under the H2 """ & level2Name & """. Please consider adding another H3 in this section or allow us to delete the heading """ & level3Name & """, as APA style requires at least two subheadings under each heading level.]" & vbCrLf
        insertionRange.Font.Bold = True
        insertionRange.style = "AQ" ' Apply style "AQ" to the inserted text
    End If
End Sub
Sub TableEmptyCells()
    Dim doc As Document
    Dim tbl As TABLE
    Dim cell As cell
    Dim emptyCellFound As Boolean
    'Dim styleName As String
    'styleName = "AQ"  ' Define the style name

    ' Set a reference to the active document
    Set doc = ActiveDocument

    ' Loop through all tables in the active document
    For Each tbl In doc.Tables
        ' Check if the table has at least one cell
        If tbl.range.Cells.count > 0 Then
            ' Get the first cell in the table
            Set cell = tbl.cell(1, 1)
            ' Check if the first cell is empty
            If Trim(cell.range.text) = vbCr & Chr(7) Then
                emptyCellFound = True
                ' Insert text above the table
                doc.range(tbl.range.Start - 1).InsertBefore "[AQ: Please provide column head for the first column in Table X.]"
            End If
        End If
    Next tbl

    ' Check if empty cells were found in any table and display a message
    'If emptyCellFound Then
        'MsgBox "Text has been inserted above tables with empty cells."
    'Else
        'MsgBox "No empty cell found in any table."
    'End If
End Sub
Sub AQcolumn()
    Dim textToFormat As String
    textToFormat = "[AQ: Please provide column head for the first column in Table X.]"
    
    ' Replace the text in your Word document
    With ActiveDocument.Content.Find
        .ClearFormatting
        .text = textToFormat
        .Replacement.ClearFormatting
        .Replacement.text = textToFormat
        .Replacement.Font.Bold = True
        .Replacement.Font.Italic = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Sub TitleCaseAPA()
    Dim lcList As String
    Dim wrd As Integer
    Dim sTest As String
    Dim p As paragraph
    
    ' List of lowercase words, surrounded by spaces
    lcList = " is a and had the that an and as but for if nor or so yet a an the as at by for in of off on per to up via vs are"
    
    For Each p In ActiveDocument.Paragraphs
        Select Case p.style
            Case "H1", "H2", "H3", "H4", "CP", "AT"
                p.range.Case = wdTitleWord
                
                For wrd = 2 To p.range.words.count
                    sTest = Trim(p.range.words(wrd))
                    sTest = " " & LCase(sTest) & " "
                    If InStr(lcList, sTest) Then
                        p.range.words(wrd).Case = wdLowerCase
                    End If
                Next wrd
        End Select
    Next p
End Sub
Sub APATableWords()
    Dim rng As range
    Dim wordList() As Variant
    Dim word As Variant
    
    ' List of words to search for
    wordList = Array("SD", "SE", "M", "p", "R", "r", "B", "t", "n", "N", "F", "d")
    
    ' Loop through each word in the list
    For Each word In wordList
        ' Find the word in the specified styles
        Set rng = ActiveDocument.range
        With rng.Find
            .text = word
            .Forward = True
            .Format = True
            .MatchWholeWord = True
            .MatchCase = True
            .Font.Italic = False ' Check for non-italic text
            Do While .Execute
                If rng.style = "TCH" Or rng.style = "TT" Then
                    rng.Font.Italic = True ' Apply italic formatting
                End If
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End With
    Next word
End Sub
Sub APAchecklist()

Call Nameanddatechecklist
Call CombinedPatternCheck

End Sub
Sub SubTitleCaseAPA10()
    Dim lcList As String
    Dim wrd As Integer
    Dim sTest As String
    Dim p As paragraph
    Dim i As Integer
    
    ' List of lowercase words, surrounded by spaces
    lcList = " is a and had the that an and as but for if nor or so yet a an the as at by for in of off on per to up via vs are"
    
    For Each p In ActiveDocument.Paragraphs
        Select Case p.style
            Case "H1", "H2", "H3", "H4", "CP", "AT"
                For wrd = 2 To p.range.words.count
                    sTest = Trim(p.range.words(wrd))
                    sTest = " " & LCase(sTest) & " "
                    If InStr(lcList, sTest) Then
                        p.range.words(wrd).Case = wdLowerCase
                    Else
                        ' Check for words with mixed case
                        Dim mixedCase As Boolean
                        mixedCase = False
                        
                        ' Check if the word has at least one uppercase letter
                        For i = 1 To Len(sTest)
                            If Mid(sTest, i, 1) Like "[A-Z]" Then
                                mixedCase = True
                                Exit For
                            End If
                        Next i
                        
                        ' Convert mixed-case words to lowercase
                        If mixedCase Then
                            p.range.words(wrd).text = LCase(sTest)
                        Else
                            ' Convert to Title Case APA style for other words
                            p.range.words(wrd).Case = wdTitleWord
                        End If
                    End If
                Next wrd
        End Select
    Next p
End Sub
Sub SubTitleCaseAPA20()
    Dim lcList As String
    Dim wrd As range
    Dim p As paragraph
    
    ' List of lowercase words, separated by spaces
    lcList = " is a and had the that an and as but for if nor or so yet a an the as at by for in of off on per to up via vs are"
    
    For Each p In ActiveDocument.Paragraphs
        Select Case p.style
            Case "H1", "H2", "H3", "H4", "CP", "AT"
                For Each wrd In p.range.words
                    Dim wordText As String
                    wordText = Trim(wrd.text)
                    
                    ' Check if the word is in the list of lowercase words
                    If InStr(1, lcList, " " & LCase(wordText) & " ") > 0 Then
                        wrd.Case = wdLowerCase
                    ElseIf UCase(wordText) = wordText Then
                        ' Skip changing casing for words with all uppercase letters
                    Else
                        ' Convert to Title Case APA style for other words
                        wrd.Case = wdTitleWord
                    End If
                Next wrd
        End Select
    Next p
End Sub
Sub SubTitleCaseAPAMicroUS()
    Dim lcList As String
    Dim wrd As range
    Dim p As paragraph
    
    ' List of lowercase words, separated by spaces
    lcList = " is a and had the that an and as but for if nor or so yet a an the as at by for in of off on per to up via vs are"
    
    For Each p In ActiveDocument.Paragraphs
    Select Case p.style
        Case "H1", "H2", "H3", "H4", "CP", "AT"
            For Each wrd In p.range.words
                Dim wordText As String
                wordText = Trim(wrd.text)
                
                ' Check if the word is in the list of lowercase words
                If InStr(1, lcList, " " & LCase(wordText) & " ") > 0 Then
                    wrd.Case = wdLowerCase
                ElseIf UCase(wordText) = wordText Then
                    ' Skip changing casing for words with all uppercase letters
                Else
                    ' Check for words with mixed case
                    Dim mixedCase As Boolean
                    mixedCase = False
                    
                    ' Check if the word has at least one uppercase letter
                    For i = 1 To Len(wordText)
                        If Mid(wordText, i, 1) Like "[A-Z]" Then
                            mixedCase = True
                            Exit For
                        End If
                    Next i
                    
                    ' Skip changing casing for words with mixed case
                    If mixedCase Then
                        ' Do nothing or handle as needed
                    Else
                        ' Convert to Title Case APA style for other words
                        wrd.Case = wdTitleWord
                    End If
                End If
            Next wrd
    End Select
Next p
End Sub

Sub TitleCaseAPAMicroUS102()
    Dim lcList As String
    Dim wrd As range
    Dim p As paragraph
    
    ' List of lowercase words, separated by spaces
    lcList = " is a and had the that an and as but for if nor or so yet a an the as at by for in of off on per to up via vs are"
    
    For Each p In ActiveDocument.Paragraphs
        Select Case p.style
            Case "H1", "H2", "H3", "H4", "CP", "AT"
                For Each wrd In p.range.words
                    Dim wordText As String
                    wordText = Trim(wrd.text)
                    
                    ' Check if the word has a character count greater than one
                    If Len(wordText) > 1 Then
                        ' Check if the word is in the list of lowercase words
                        If InStr(1, lcList, " " & LCase(wordText) & " ") > 0 Then
                            'wrd.Case = wdLowerCase
                        ElseIf UCase(wordText) = wordText Then
                            ' Skip changing casing for words with all uppercase letters
                        Else
                            ' Check for words with mixed case
                            Dim mixedCase As Boolean
                            mixedCase = False
                            
                            ' Check if the word has at least one uppercase letter
                            For i = 1 To Len(wordText)
                                If Mid(wordText, i, 1) Like "[A-Z]" Then
                                    mixedCase = True
                                    Exit For
                                End If
                            Next i
                            
                            ' Skip changing casing for words with mixed case
                            If mixedCase Then
                                ' Do nothing or handle as needed
                            Else
                                ' Convert to Title Case APA style for other words
                                wrd.Case = wdTitleWord
                            End If
                        End If
                    End If
                Next wrd
        End Select
    Next p
End Sub


Sub CheckHeadlevels21()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim para As paragraph
    Dim level1Name As String
    Dim level2Name As String
    Dim hdg1Count As Integer
    Dim hdg2Count As Integer
    Dim insertionRange As range
    Dim previousPara As paragraph
    
    hdg1Count = 0
    hdg2Count = 0
    
    For Each para In doc.Paragraphs
        If para.style <> "" Then
            If para.outlineLevel = wdOutlineLevel1 Then
                ' Check Heading Level 1 count and Heading Level 2 count
                If hdg1Count > 0 And hdg2Count = 1 Then
                    Set insertionRange = previousPara.range
                    insertionRange.Collapse wdCollapseEnd
                    insertionRange.InsertAfter vbCrLf & "[AQ: There is only one H2 “" & level2Name & "” given under the H1 “" & level1Name & "”. Please consider adding another H2 in this section or allow us to delete the heading “" & level2Name & "”, as APA style requires at least two subheadings under each heading level.]" & vbCrLf
                    insertionRange.Font.Bold = True
                    insertionRange.style = "AQ" ' Apply style "AQ" to the inserted text
                End If
                level1Name = Left(para.range.text, Len(para.range.text) - 1)
                hdg1Count = hdg1Count + 1
                hdg2Count = 0
            ElseIf para.outlineLevel = wdOutlineLevel2 Then
                ' Increment Heading Level 2 count and remove the last character (line break)
                level2Name = Left(para.range.text, Len(para.range.text) - 1)
                hdg2Count = hdg2Count + 1
            End If
        End If
        Set previousPara = para ' Store the previous paragraph
    Next para
    
    ' Check for the last heading group
    If hdg1Count > 0 And hdg2Count = 1 Then
        Set insertionRange = previousPara.range
        insertionRange.Collapse wdCollapseEnd
        insertionRange.InsertAfter vbCrLf & "[AQ: There is only one H2 """ & level2Name & """ given under the H1 """ & level1Name & """. Please consider adding another H2 in this section or allow us to delete the heading """ & level2Name & """, as APA style requires at least two subheadings under each heading level.]" & vbCrLf
        insertionRange.Font.Bold = True
        insertionRange.style = "AQ" ' Apply style "AQ" to the inserted text
    End If
End Sub
