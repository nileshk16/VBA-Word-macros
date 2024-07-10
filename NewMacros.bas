Attribute VB_Name = "NewMacros"
Sub etal()
Attribute etal.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.etal"
'
' etal Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Italic = True
    With Selection.Find
        .text = "et al."
        .Replacement.text = "et al."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        '.Font.Italic = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub FindReplaceInWord()

    Dim Wbk As Workbook: Set Wbk = ThisWorkbook
    Dim wrd As New word.Application
    Dim Dict As Object
    Dim RefList As range, RefElem As range

    wrd.Visible = True
    Dim WDoc As Document
    Set WDoc = wrd.Documents.Open("C:\Users\Admin\Downloads\Input.docx") 'Modify as necessary.

    Set Dict = CreateObject("Scripting.Dictionary")
    Set RefList = Wbk.Sheets("Sheet1").range("A1:A3") 'Modify as necessary.

    With Dict
        For Each RefElem In RefList
            If Not .Exists(RefElem) And Not IsEmpty(RefElem) Then
                .Add RefElem.Value, RefElem.Offset(0, 1).Value
            End If
        Next RefElem
    End With

    For Each Key In Dict
        With WDoc.Content.Find
            .Execute findText:=Key, ReplaceWith:=Dict(Key)
        End With
    Next Key


End Sub
Sub renumber()
'Updateby ExtendOffice
Dim xWordApp As word.Application
Dim xDoc As word.Document
Dim xRng As range
Dim i As Integer
Dim xFileDlg As FileDialog
On Error GoTo ExitSub
Set xFileDlg = Application.FileDialog(msoFileDialogFilePicker)
xFileDlg.AllowMultiSelect = False
xFileDlg.Filters.Add "Word Document", "*.docx; *.doc; *.docm"
xFileDlg.FilterIndex = 2
If xFileDlg.Show <> -1 Then GoTo ExitSub
Set xRng = Application.InputBox("Please select the lists of find and replace texts (Press Ctrl key to select two same size ranges):", "Kutools for Excel", , , , , , 8)
If xRng.Areas.count <> 2 Then
  MsgBox "Please select two columns (press Ctrl key), the two ranges have the same size.", vbInformation + vbOKOnly, "Kutools for Excel"
  GoTo ExitSub
End If
If (xRng.Areas.Item(1).Rows.count <> xRng.Areas.Item(2).Rows.count) Or _
  (xRng.Areas.Item(1).Columns.count <> xRng.Areas.Item(2).Columns.count) Then
  MsgBox "Please select two columns (press Ctrl key), the two ranges have the same size.", vbInformation + vbOKOnly, "Kutools for Excel"
  GoTo ExitSub
End If
Set xWordApp = CreateObject("Word.application")
xWordApp.Visible = True
Set xDoc = xWordApp.Documents.Open(xFileDlg.SelectedItems.Item(1))
For i = 1 To xRng.Areas.Item(1).Cells.count
  With xDoc.Application.Selection.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .text = xRng.Areas.Item(1).Cells.Item(i).Value
    .Replacement.text = xRng.Areas.Item(2).Cells.Item(i).Value
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
  End With
  xDoc.Application.Selection.Find.Execute Replace:=wdReplaceAll
Next
ExitSub:
  Set xRng = Nothing
  Set xFileDlg = Nothing
  Set xWordApp = Nothing
  Set xDoc = Nothing
End Sub

Sub endash()
With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "-"
        .Replacement.text = "^="
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
        .text = "="
        .Replacement.text = "^t^&"
        .Execute Replace:=wdReplaceAll
    End With

End Sub
Sub TSFigure()
With ActiveDocument.Content.Find
 .ClearFormatting
 .text = "[FIGURE"
 .style = "CL"
 .Replacement.ClearFormatting
 .Replacement.text = "[TS: PLEASE INSERT FIGURE"
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub
Sub TSFigure2()
With ActiveDocument.Content.Find
 .ClearFormatting
 .text = "ABOUT HERE]"
 .style = "CL"
 .Replacement.ClearFormatting
 .Replacement.text = "ABOUT HERE.]"
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub
Sub TSinsert()
    
    Call TSFigure
    
    
    Call TSFigure2
    
    
End Sub
Sub CheckHyperlinks()
    Dim doc As Document
    Dim hyperlink As hyperlink
    
    Set doc = ActiveDocument
    
    For Each hyperlink In doc.Hyperlinks
        If hyperlink.address <> "" Then
            ' Temporarily store the original formatting of the hyperlink range
            Dim originalRange As range
            Set originalRange = hyperlink.range.Duplicate
            
            ' Apply formatting to the hyperlink based on whether it is working or not
            hyperlink.range.HighlightColorIndex = wdBrightGreen ' Green color for working links
            
            Dim isWorking As Boolean
            isWorking = TestHyperlink(hyperlink.address)
            
            If Not isWorking Then
                ' Hyperlink is not working
                hyperlink.range.InsertAfter "[AQ: Please note that the given URL in 'XXXXXX' does not lead to the desired web page. Please provide an active URL.]"
                'hyperlink.range.HighlightColorIndex = wdRed ' Red color for invalid links
            End If
            
            ' Restore the original formatting of the hyperlink range
            RestoreHyperlinkFormatting hyperlink.range, originalRange
            
        End If
    Next hyperlink
    MsgBox "Hyperlink checking completed!", vbInformation
End Sub
Function TestHyperlink(ByVal address As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Disable certificate validation for HTTPS links
    request.Option(4) = &H3300 ' WinHttpRequestOption_SslErrorIgnoreFlags
    
    ' Enable automatic redirects
    request.Option(6) = True ' WinHttpRequestOption_EnableRedirects
    
    request.Open "GET", address, False
    request.send
    
    TestHyperlink = (request.Status = 200)
    
    Set request = Nothing
    Exit Function
    
ErrorHandler:
    TestHyperlink = False
    Set request = Nothing
End Function
Sub RestoreHyperlinkFormatting(ByVal targetRange As range, ByVal originalRange As range)
    targetRange.Font.Color = originalRange.Font.Color
    targetRange.Font.Bold = originalRange.Font.Bold
    targetRange.Font.Underline = originalRange.Font.Underline
    targetRange.Font.Italic = originalRange.Font.Italic
End Sub

Sub REFAftercolon()
    Dim searchTerm As String
    Dim rng As range
    
    ' Set the search term with wildcard
    searchTerm = ": [A-Z]"
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    'rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight the search term with the specified style
    With rng.Find
        .text = searchTerm
        .style = ActiveDocument.Styles("REF")
        '.Format = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        
        ' Loop through each found item
        Do While .Execute
            rng.HighlightColorIndex = wdRed
            rng.Collapse wdCollapseEnd
            'rng.Case = wdLowerCase
        Loop
    End With
End Sub
Sub REFAftercolon2()
    Dim searchTerm As String
    Dim rng As range
    
    ' Set the search term with wildcard
    searchTerm = ": [a-z]"
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    'rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight the search term with the specified style
    With rng.Find
        .text = searchTerm
        .style = ActiveDocument.Styles("REF")
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        
        ' Loop through each found item
        Do While .Execute
            rng.HighlightColorIndex = wdRed
            rng.Collapse wdCollapseEnd
            'rng.Case = wdLowerCase
        Loop
    End With
End Sub
Sub Keywordschangestyle()
    Dim searchTerm As String
    Dim rng As range
    
    ' Set the search term
    searchTerm = "Keywords: "
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Find and change the style of the search term
    With rng.Find
        .text = searchTerm
        .style = ActiveDocument.Styles("ABKW")
        .Format = True
        '.Font.Italic = False
        '.Font.Bold = False
        .Forward = True
        .Wrap = wdFindStop
        
        ' Loop through each found item
        Do While .Execute
            rng.Font.Bold = False
            rng.Font.Italic = False
            'rng.InsertAfter vbCr ' Insert an extra line brea
            rng.style = ActiveDocument.Styles("ABKWH")
            rng.Collapse wdCollapseEnd
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub
Sub REFpwithoutspace()
    Dim searchTerm As String
    Dim rng As range
    
    ' Set the search term with wildcard
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    'rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight the search term with the specified style
    With rng.Find
        .text = "p.^#"
        .style = ActiveDocument.Styles("REF")
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
Sub ATAftercolon1()
    Dim searchTerm As String
    Dim rng As range
    
    ' Set the search term with wildcard
    searchTerm = ": ([A-Z])"
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    'rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight the search term with the specified style
    With rng.Find
        .text = searchTerm
        .style = ActiveDocument.Styles("AT")
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
Sub FindSuperscriptInAUStyle()
    Dim rng As range
    
    Set rng = ActiveDocument.Content
    
    'rng.HighlightColorIndex = wdNoHighlight
    
    With rng.Find
        .ClearFormatting
        .Font.Superscript = True
        .style = "AU"
        .Execute Format:=True
        
        While .found
            rng.HighlightColorIndex = wdBlue
            rng.Collapse Direction:=wdCollapseEnd
            .Execute Format:=True
        Wend
    End With
End Sub
Sub CheckSentencesInStyle()
    Dim doc As Document
    Dim rng As range
    Dim styleName As String
    Dim sentences() As Variant
    Dim sentence As Variant
    Dim found As Boolean
    
    ' Set the style name
    styleName = "EH"
    
    ' Set the sentences to be checked
    sentences = Array("Ethics approval and consent to participate", _
                      "Consent for publication", _
                      "Author contributions", _
                      "Acknowledgments", _
                      "Funding", _
                      "Competing interests", _
                      "Availability of data and materia")
    
    ' Set the document to be checked
    Set doc = ActiveDocument
    
    ' Set the range to the entire document
    Set rng = doc.range
    
    ' Loop through each sentence
    For Each sentence In sentences
        ' Reset the found flag
        found = False
        
        ' Reset the range to the entire document
        rng.SetRange Start:=doc.range.Start, End:=doc.range.End
        
        ' Search for the sentence in the specified style
        With rng.Find
            .ClearFormatting
            .text = sentence
            .style = doc.Styles(styleName)
            .Forward = True
            .Wrap = wdFindStop
            .Execute
        End With
        
        ' Check if the sentence was found
        If rng.Find.found Then
            found = True
            Exit For
        End If
    Next sentence
    
    ' Display the result
    If found Then
        MsgBox "All sentences found in style '" & styleName & "'."
    Else
        MsgBox "One or more sentences not found in style '" & styleName & "'."
    End If
End Sub
Sub pwithoutspace()
    Dim searchTerm As String
    Dim rng As range
    
    ' Set the search term with wildcard

    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    'rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight the search term with the specified style
    With rng.Find
        .text = "p.^#"
        .Replacement.text = ""
        .style = "REF"
        .Forward = True
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Loop through each found item
        Do While .Execute
            rng.HighlightColorIndex = wdRed
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub
Sub commaand()
    Dim searchTerm As String
    Dim rng As range
    
    ' Set the search term with wildcard
    searchTerm = ", and"
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    'rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight the search term with the specified style
    With rng.Find
        .text = searchTerm
        .style = ActiveDocument.Styles("AU")
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
Sub ABKW()
    Dim findArray() As Variant
    Dim character As Variant
    Dim count As Integer
    
    ' Add the characters you want to find and highlight in the array below
    findArray = Array("Background:", "Design:", "Results:", "Methods:", "Conclusion:")
    
    count = 0
    
    ' Loop through each character in the array and highlight them in red
    For Each character In findArray
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = character
            .Font.Bold = False
            .Replacement.Highlight = True
            .Forward = True
            .style = "ABKW"
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Do While Selection.Find.Execute
            Selection.range.HighlightColorIndex = wdRed
            Selection.Collapse wdCollapseEnd
            count = count + 1
        Loop
    Next character
    
    End Sub
Sub HighlightTablePattern()
    Dim doc As Document
    Dim rng As range
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    ' Create a regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = "Table \d+\." ' Pattern to match 'Table' followed by any sequence of digits and a dot
    
    ' Set the document to search
    Set doc = ActiveDocument
    
    ' Set the range to search within the entire document
    Set rng = doc.Content
    
    ' Clear any existing selection and formatting
    doc.Select
    Selection.range.HighlightColorIndex = wdNoHighlight
    
    ' Find all matches in the document
    Set matches = regex.Execute(rng.text)
    
    ' Loop through each match and highlight it in green
    For Each match In matches
        Set rng = doc.range(Start:=match.FirstIndex, End:=match.FirstIndex + match.Length)
        rng.HighlightColorIndex = wdGreen
    Next match
End Sub
Sub nameanddate()
    Dim findArray() As Variant
    Dim character As Variant
    
    ' Add the characters you want to find and highlight in the array below
    findArray = Array("et al., ^#^#^#^#")
    
    ' Loop through each character in the array and highlight them in red
    For Each character In findArray
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = character
            .Replacement.Highlight = True
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Do While Selection.Find.Execute
            Selection.range.HighlightColorIndex = wdRed
            Selection.Collapse wdCollapseEnd
        Loop
    Next character
     MsgBox count & " characters highlighted in red.", vbInformation, "Highlighting Completed"
End Sub

Sub FormatH1Tags()
    Dim doc As Document
    Dim rng As range
    Dim shapeRange As shapeRange
    
    Set doc = ActiveDocument
    
    ' Loop through all shapes in the document
    For Each Shape In doc.Shapes
        If Shape.Type = msoTextBox Then
            Set shapeRange = Shape.TextFrame.TextRange
            
            ' Loop through all paragraphs in the shape
            For Each para In shapeRange.Paragraphs
                If InStr(para.range.text, "<H1>") > 0 And InStr(para.range.text, "</H1>") > 0 Then
                    ' Set bold and center alignment for the H1 tag
                    para.range.Font.Bold = True
                    para.range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                End If
            Next para
        End If
    Next Shape
End Sub

Sub HighlightText()
    Dim doc As Document
    Dim rng As range
    Dim findRange As range
    
    Set doc = ActiveDocument
    
    ' Set the range to search for text
    Set rng = doc.Content
    
    ' Create a Find object
    Set findRange = rng.Duplicate
    With findRange.Find
        .ClearFormatting
        .Font.AllCaps = True
        .style = doc.Styles("AU")
        .text = ""
        .Forward = True
        .Wrap = wdFindStop
    End With
    
    ' Loop through each found instance
    Do While findRange.Find.Execute
        ' Check if the found text ends with a dot
        If Right(findRange.text, 1) = "." Then
            ' Highlight the text in green
            findRange.HighlightColorIndex = wdGreen
        Else
            ' Highlight the text in red
            findRange.HighlightColorIndex = wdRed
        End If
        
        ' Move the range to the next instance
        findRange.Collapse wdCollapseEnd
    Loop
    
    ' Clear the findRange object
    Set findRange = Nothing
    
    ' Clear the rng object
    Set rng = Nothing
    
    ' Clear the doc object
    Set doc = Nothing
End Sub
Sub LRH()
    Dim doc As Document
    Dim para As paragraph
    
    Set doc = ActiveDocument
    
    For Each para In doc.Paragraphs
        If para.style = "LRH" Then
            para.range.Font.Italic = True
        End If
    Next para
End Sub
Sub RRH()
    Dim doc As Document
    Dim para As paragraph
    
    Set doc = ActiveDocument
    
    For Each para In doc.Paragraphs
        If para.style = "RRH" Then
            para.range.Font.Italic = True
        End If
    Next para
End Sub
Sub TY()
    Dim doc As Document
    Dim para As paragraph
    
    Set doc = ActiveDocument
    
    For Each para In doc.Paragraphs
        If para.style = "TY" Then
            para.range.Font.Italic = True
        End If
    Next para
End Sub

Sub LRHXX()
    Dim doc As Document
    Dim para As paragraph
    Dim rng As range
    
    Set doc = ActiveDocument
    
    For Each para In doc.Paragraphs
        If para.style = "LRH" Then
            Set rng = para.range
            rng.MoveEnd wdCharacter, -1 ' Move the end of the range before the paragraph mark
            rng.text = rng.text & " XX(X)"
        End If
    Next para
End Sub
Sub LRHX()
    Dim doc As Document
    Dim para As paragraph
    Dim rng As range
    
    Set doc = ActiveDocument
    
    For Each para In doc.Paragraphs
        If para.style = "LRH" Then
            Set rng = para.range
            rng.MoveEnd wdCharacter, -1 ' Move the end of the range before the paragraph mark
            rng.text = rng.text & " X(X)"
        End If
    Next para
End Sub
Sub UKFrontmatter()

Call LRHX
Call LRH
Call RRH
Call TY
Call MoveTY

End Sub
Sub APAFrontmatter()

Call LRHXX
Call LRH
Call RRH

End Sub
Sub AMAFrontmatter()

Call LRHXX
Call LRH
Call RRH
Call RRHdot
End Sub
Sub FindAndHighlightDuplicateWordsInStyleAF()
    Dim doc As Document
    Dim rng As range
    Dim words As Variant
    Dim word As range
    Dim style As String
    Dim duplicates As Collection
    Dim duplicateWord As Variant
    Dim colorIndex As Integer
    Dim highlightColors() As Long
    ReDim highlightColors(10 To 20) ' Array to store custom highlight colors
    
    ' Set the style name to search for duplicates
    style = "AF"
    
    Set doc = ActiveDocument
    Set rng = doc.Content
    
    ' Initialize collection to store duplicate words
    Set duplicates = New Collection
    
    ' Set the range to the entire document
    rng.WholeStory
    
    ' Loop through each word in the document
    For Each word In rng.words
        ' Check if the word is in the specified style
        If word.style = style Then
            ' Convert word to lowercase for case-insensitive comparison
            Dim lowerCaseWord As String
            lowerCaseWord = LCase(word.text)
            
            ' Check if the word is already in the duplicates collection
            On Error Resume Next
            duplicateWord = duplicates(lowerCaseWord)
            On Error GoTo 0
            
            ' If word is not already in duplicates collection, add it
            If IsEmpty(duplicateWord) Then
                duplicates.Add lowerCaseWord, CStr(lowerCaseWord)
            Else
                ' Check if a highlight color has already been assigned to this word
                If Not IsNumeric(duplicateWord) Then
                    ' Assign a new highlight color for this word
                    colorIndex = colorIndex + 1
                    duplicateWord = colorIndex
                End If
                
                ' Highlight the duplicate word with the assigned color
                word.HighlightColorIndex = highlightColors(duplicateWord)
            End If
        End If
    Next word
    
    ' Display the duplicate words
    If duplicates.count > 0 Then
        For Each duplicateWord In duplicates
            MsgBox "Duplicate word found in style 'AF': " & duplicateWord, vbInformation, "Duplicate Word"
        Next duplicateWord
    Else
        MsgBox "No duplicate words found in style 'AF'.", vbInformation, "Duplicate Word"
    End If
    
    ' Clear memory
    Set doc = Nothing
    Set rng = Nothing
    Set duplicates = Nothing
End Sub
Sub MoveTY()
    Dim doc As Document
    Dim rangeTY As range
    Dim rangeAT As range
    Dim startTY As Long
    Dim endTY As Long
    Dim startAT As Long
    
    ' Set the document object
    Set doc = ActiveDocument
    
    ' Set the range for 'TY' style text
    Set rangeTY = doc.range
    
    ' Set the range for 'AT' style text
    Set rangeAT = doc.range
    
    ' Find the 'TY' style text
    With rangeTY.Find
        .ClearFormatting
        .style = "TY"
        .Execute
        If .found Then
            startTY = rangeTY.Start
            endTY = rangeTY.End
        End If
    End With
    
    ' Find the 'AT' style text
    With rangeAT.Find
        .ClearFormatting
        .style = "AT"
        .Execute
        If .found Then
            startAT = rangeAT.Start
        End If
    End With
    
    ' Move the 'TY' style text above the 'AT' style text
    If startTY < startAT Then
        doc.range(startTY, endTY).Cut
        doc.range(startAT, startAT).Paste
    End If
End Sub
Sub CountWordsInStyle()
    Dim doc As Document
    Dim styleName As String
    Dim wordCount As Integer
    Dim word As range
    
    ' Set the style name to count words
    styleName = "ABKW"
    
    ' Set the document reference
    Set doc = ActiveDocument
    
    ' Reset the word count
    wordCount = 0
    
    ' Loop through each word in the document
    For Each word In doc.words
        ' Check if the word has the specified style
        If word.style = styleName Then
            wordCount = wordCount + 1
        End If
    Next word
    
    ' Display the word count
    MsgBox "Number of words with style '" & styleName & "': " & wordCount
End Sub
Sub InsertNoteAtBeginning()
    Dim rng As range
    Set rng = ActiveDocument.Content
    rng.Collapse wdCollapseStart
    rng.text = "Note. " & rng.text
    rng.style = "CPSO"
End Sub
Sub qwe()
    Dim doc As Document
    Dim para As paragraph
    Dim rng As range
    
    Set doc = ActiveDocument
    
    For Each para In doc.Paragraphs
        If para.style = "LRH" Then
            Set rng = para.range
            rng.MoveStart wdCharacter, 1 ' Move the start of the range after the paragraph mark
            rng.text = "X(X) " & rng.text
        End If
    Next para
End Sub
Sub AfterColonAS()
    Dim rng As range
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .text = ":*"
        .style = "AT"
        .MatchWildcards = True
        
        Do While .Execute
            rng.MoveStartUntil ": "
            rng.MoveStart 2
            
            If rng.End = rng.Paragraphs.Last.range.End Then
                rng.End = rng.End + 1
            Else
                rng.MoveEndUntil vbCr
            End If
            
            rng.InsertParagraphBefore
            rng.Collapse wdCollapseEnd
            rng.style = "AS" ' Change style to "AS"
        Loop
    End With
End Sub
Sub AfterColonAS22()
    Dim rng As range
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .text = ":*"
        .style = "ABKW"
        .MatchWildcards = True
        
        Do While .Execute
            rng.MoveStartUntil ": "
            rng.MoveStart 2
            
            If rng.End = rng.Paragraphs.Last.range.End Then
                rng.End = rng.End + 1
            Else
                rng.MoveEndUntil vbCr
            End If
            
            rng.InsertParagraphBefore
            rng.Collapse wdCollapseEnd
            rng.style = "AS" ' Change style to "AS"
        Loop
    End With
End Sub
Sub MoveTYa()
    Dim doc As Document
    Dim rangeTY As range
    Dim rangeAT As range
    Dim startTY As Long
    Dim startAT As Long
    
    ' Set the document object
    Set doc = ActiveDocument
    
    ' Set the range for 'TY' style text
    Set rangeTY = doc.Content
    
    ' Set the range for 'AT' style text
    Set rangeAT = doc.Content
    
    ' Find the 'TY' style text
    With rangeTY.Find
        .ClearFormatting
        .style = "TY"
        .Execute
        If .found Then
            startTY = rangeTY.Start
        End If
    End With
    
    ' Find the 'AT' style text
    With rangeAT.Find
        .ClearFormatting
        .style = "AT"
        .Execute
        If .found Then
            startAT = rangeAT.Start
        End If
    End With
    
    ' Move the 'TY' style text below the 'AT' style text
    If startTY > startAT Then
        rangeTY.Cut
        doc.range(startAT, startAT).Collapse wdCollapseEnd
        doc.range.Paste
    End If
End Sub
Sub AFdot()
    Dim searchTerm As String
    Dim rng As range
    
    ' Set the search term with wildcard

    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    'rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight the search term with the specified style
    With rng.Find
        .text = "^$."
        .Forward = True
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .style = "AF"
        
        ' Loop through each found item
        Do While .Execute
            rng.HighlightColorIndex = wdGreen
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub

Sub AUdot()
    Dim searchTerm As String
    Dim rng As range
    
    ' Set the search term with wildcard

    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    'rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight the search term with the specified style
    With rng.Find
        .text = "^$."
        .Forward = True
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .style = "AU"
        
        ' Loop through each found item
        Do While .Execute
            rng.HighlightColorIndex = wdBlue
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub
Sub ATAftercolon2()
    Dim searchTerm As String
    Dim rng As range
    
    ' Set the search term with wildcard
    searchTerm = ": ([a-z])"
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    'rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight the search term with the specified style
    With rng.Find
        .text = searchTerm
        .style = ActiveDocument.Styles("AT")
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
Sub pwithspace()
    Dim searchTerm As String
    Dim rng As range
    
    ' Set the search term with wildcard

    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    'rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight the search term with the specified style
    With rng.Find
        .text = "p. ^#"
        .Replacement.text = ""
        .style = "REF"
        .Forward = True
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Loop through each found item
        Do While .Execute
            rng.HighlightColorIndex = wdRed
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub
Sub macros2()
Attribute macros2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.macros2"
'
' macros2 Macro
'
'
    Selection.range.Case = wdTitleSentence
    Selection.range.Case = wdTitleSentence
    Selection.EscapeKey
    Selection.EscapeKey
End Sub
Sub AUdotremover()
With ActiveDocument.Content.Find
 .ClearFormatting
 .text = "."
 .style = "AU"
 .Replacement.ClearFormatting
 .Replacement.text = ""
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub

Sub FindFourDigitNumbers()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim rng As range
    Set rng = doc.Content
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "\b\w+,\s(\d{4})" ' Matches any word followed by a comma, space, and a four-digit number
    
    Dim matches As Object
    Set matches = regex.Execute(rng.text)
    
    Dim match As Object
    For Each match In matches
        rng.Find.ClearFormatting
        rng.Find.text = match.Value
        rng.Find.Execute
        
        ' Highlight the found text
        If rng.Find.found Then
            rng.HighlightColorIndex = wdYellow
        End If
    Next match
End Sub

Sub FindSentencesWithMoreThan40Words()
    Dim rng As range
    Dim doc As Document
    Dim sentence As range
    Dim sentenceCount As Integer
    Dim wordCount As Integer
    
    ' Set the document to search in (change the document name if needed)
    Set doc = ActiveDocument
    
    ' Set the range to search in (the whole document in this case)
    Set rng = doc.Content
    
    ' Clear previous search results
    doc.range.HighlightColorIndex = wdNoHighlight
    
    ' Loop through all sentences in the document
    For Each sentence In rng.sentences
        ' Check if the sentence is enclosed within double quotes
        If InStr(1, sentence, Chr(34)) > 0 Then
            ' Count the number of words in the sentence
            wordCount = UBound(Split(sentence.text, " ")) + 1
            
            ' Check if the sentence has more than 40 words
            If wordCount > 40 Then
                ' Highlight the sentence
                sentence.Select
                Selection.range.HighlightColorIndex = wdYellow
                
                ' Increment the count of sentences
                sentenceCount = sentenceCount + 1
            End If
        End If
    Next sentence
    
    ' Display the number of sentences found
    MsgBox "Found " & sentenceCount & " sentence(s) with more than 40 words."
End Sub

Sub FindWordsInsideQuotesUsingFind()
    Dim rng As range
    Dim doc As Document
    Dim quotePattern As String
    Dim found As Boolean
    
    ' Set the document to search in (change the document name if needed)
    Set doc = ActiveDocument
    
    ' Set the range to search in (the whole document in this case)
    Set rng = doc.Content
    
    ' Clear previous search results
    doc.range.HighlightColorIndex = wdNoHighlight
    
    ' Set the pattern to search for words inside quotes
    quotePattern = "([“”])([^“”^13^32]+)([“”])"
    
    ' Find the first occurrence of the pattern
    found = rng.Find.Execute(findText:=quotePattern, MatchWildcards:=True)
    
    ' Loop through all occurrences of the pattern
    Do While found
        ' Highlight the found range
        rng.HighlightColorIndex = wdYellow
        
        ' Find the next occurrence of the pattern
        found = rng.Find.Execute(findText:=quotePattern, MatchWildcards:=True)
    Loop
    
    ' Display a message indicating the search is complete
    MsgBox "Search complete."
End Sub

Sub AfterColonABKW()
    Dim rng As range
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .text = "Keywords:*"
        .style = "ABKWH"
        .MatchWildcards = True
        
        Do While .Execute
            rng.MoveStartUntil ": "
            rng.MoveStart 2
            
            If rng.End = rng.Paragraphs.Last.range.End Then
                rng.End = rng.End + 1
            Else
                rng.MoveEndUntil vbCr
            End If
            
            rng.InsertParagraphBefore
            rng.Collapse wdCollapseEnd
            rng.style = "ABKW" ' Change style to "AS"
        Loop
    End With
End Sub

Sub HighlightConflictsOfInterest()
    Dim doc As Document
    Dim rng As range
    Dim searchString As String
    
    searchString = "The author(s) declared no potential conflicts of interest with respect to the research, authorship, and/or publication of this article."
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Set the range to search in
    Set rng = doc.range
    
    ' Set the search parameters
    With rng.Find
        .text = searchString
        .MatchWholeWord = True
        .MatchCase = False
        .Wrap = wdFindStop
        .Format = False
    End With
    
    ' Search for the sentence and highlight it if found
    Do While rng.Find.Execute
        If rng.style = "AN" Then
            rng.HighlightColorIndex = wdGreen
        End If
        rng.Collapse wdCollapseEnd
    Loop
End Sub
Sub HighlightFunding()
    Dim doc As Document
    Dim rng As range
    Dim searchString As String
    
    searchString = "The author(s) received no financial support for the research, authorship, and/or publication of this article."
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Set the range to search in
    Set rng = doc.range
    
    ' Set the search parameters
    With rng.Find
        .text = searchString
        .MatchWholeWord = True
        .MatchCase = False
        .Wrap = wdFindStop
        .Format = False
    End With
    
    ' Search for the sentence and highlight it if found
    Do While rng.Find.Execute
        If rng.style = "AN" Then
            rng.HighlightColorIndex = wdGreen
        End If
        rng.Collapse wdCollapseEnd
    Loop
End Sub
Sub HighlightConflictsAndFunding()
    Dim doc As Document
    Dim rng As range
    Dim searchString1 As String
    Dim searchString2 As String
    Dim found As Boolean
    
    searchString1 = "Declaration of Conflicting Interests"
    searchString2 = "Funding"
    found = False
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Set the range to search in
    Set rng = doc.range
    
    ' Set the search parameters
    With rng.Find
        .text = searchString1
        .MatchWholeWord = True
        .MatchCase = False
        .Wrap = wdFindStop
        .Format = False
    End With
    
    ' Search for the first sentence and highlight it if found
    Do While rng.Find.Execute
        If rng.style = "EH" Then
            rng.HighlightColorIndex = wdGreen
            found = True
        End If
        rng.Collapse wdCollapseEnd
    Loop
    
    ' Reset the range to search again for the second sentence
    Set rng = doc.range
    
    ' Set the search parameters for the second sentence
    With rng.Find
        .text = searchString2
        .MatchWholeWord = True
        .MatchCase = False
        .Wrap = wdFindStop
        .Format = False
    End With
    
    ' Search for the second sentence and highlight it if found
    Do While rng.Find.Execute
        If rng.style = "EH" Then
            rng.HighlightColorIndex = wdGreen
            found = True
        End If
        rng.Collapse wdCollapseEnd
    Loop
    
    ' Display message box if neither of the sentences is found
    If Not found Then
        MsgBox "The sentences were not found in the specified style."
    End If
End Sub
Sub HighlightConflictsAndFundin1g()
    Dim doc As Document
    Dim rng As range
    Dim searchString1 As String
    Dim searchString2 As String
    Dim found As Boolean
    
    searchString1 = "Declaration of Conflicting Interests"
    searchString2 = "Funding"
    found = False
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Set the range to search in
    Set rng = doc.range
    
    ' Set the search parameters
    With rng.Find
        .text = searchString1
        .MatchWholeWord = True
        .MatchCase = False
        .Wrap = wdFindStop
        .Format = False
    End With
    
    ' Search for the first sentence and highlight it if found
    Do While rng.Find.Execute
        If rng.style = "EH" Then
            rng.HighlightColorIndex = wdGreen
            found = True
            Exit Do
        End If
        rng.Collapse wdCollapseEnd
    Loop
    
    ' Reset the range to the end of the document if style 'EH' is not found
    If Not found Then
        Set rng = doc.range
        rng.Collapse wdCollapseEnd
    End If
    
    ' Set the search parameters for the second sentence
    With rng.Find
        .text = searchString2
        .MatchWholeWord = True
        .MatchCase = False
        .Wrap = wdFindStop
        .Format = False
    End With
    
    ' Search for the second sentence and highlight it if found
    Do While rng.Find.Execute
        If rng.style = "EH" Then
            rng.HighlightColorIndex = wdGreen
            found = True
            Exit Do
        End If
        rng.Collapse wdCollapseEnd
    Loop
    
    ' Insert the sentence 'Conflict and Funding not found' after the first occurrence of the style 'EH' if neither of the sentences is found
    If Not found Then
        rng.InsertAfter vbCrLf & "Conflict and Funding not found"
        rng.style = "EH"
        rng.Collapse wdCollapseEnd
        rng.HighlightColorIndex = wdGreen
    End If
End Sub
Sub AddSpaceAfterURL()
    Dim doc As Document
    Dim rng As range
    Dim urlStyle As style
    
    ' Set the document variable to the active document
    Set doc = ActiveDocument
    
    ' Set the range variable to the entire document
    Set rng = doc.Content
    
    ' Set the style variable to the 'url' style
    Set urlStyle = doc.Styles("url")
    
    ' Loop through each paragraph in the range
    For Each para In rng.Paragraphs
        ' Check if the paragraph's style is 'url'
        If para.style = urlStyle.Name Then
            ' Add a space at the end of the paragraph
            para.range.InsertAfter " "
        End If
    Next para
    
    ' Clean up the objects
    Set rng = Nothing
    Set doc = Nothing
    Set urlStyle = Nothing
    
    ' Notify the user when the operation is complete
    MsgBox "Space added after 'url' style paragraphs."
End Sub



Sub FindCapitalWordsWithPeriod()
    Dim rng As range
    Dim pattern As String
    
    ' Set the search pattern for a capital word followed by a period
    pattern = "([A-Z][a-z]+\. )"
    
    ' Set the range to search within the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    'rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight each occurrence of the pattern
    With rng.Find
        .ClearFormatting
        .text = pattern
        .Forward = True
        .MatchWildcards = True
        
        Do While .Execute
            rng.HighlightColorIndex = wdYellow
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    ' Show a message box with the number of occurrences found
    MsgBox "Found " & rng.ComputeStatistics(wdStatisticWords) & " occurrences.", vbInformation
End Sub

Sub HighlightConflictsAndFundinga()
    Dim doc As Document
    Dim rng As range
    Dim searchStrings() As String
    Dim searchString As Variant
    Dim found As Boolean
    
    searchStrings = Array("Ethics approval and consent to participate", "Consent for publication", "Author contributions", "Acknowledgments", "Funding", "Competing interests", "Availability of data and materials")
    found = False
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Set the range to search in
    Set rng = doc.range
    
    ' Iterate through each search string
    For Each searchString In searchStrings
        ' Set the search parameters
        With rng.Find
            .text = searchString
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindStop
            .Format = False
        End With
        
        ' Search for the sentence and highlight it if found
        Do While rng.Find.Execute
            If rng.style = "EH" Then
                rng.HighlightColorIndex = wdGreen
                found = True
            End If
            rng.Collapse wdCollapseEnd
        Loop
        
        ' Reset the range to search again
        Set rng = doc.range
    Next searchString
    
    ' Display message box if none of the sentences is found
    If Not found Then
        MsgBox "The sentences were not found in the specified style."
    End If
End Sub

Sub SelectTextByStyle()
    Dim doc As Document
    Dim rng As range
    Dim styleName As String
    
    ' Set the style name
    styleName = "ABKW"
    
    ' Check if a document is open
    If Documents.count = 0 Then
        MsgBox "No document is open.", vbExclamation
        Exit Sub
    End If
    
    ' Set the active document
    Set doc = ActiveDocument
    
    ' Set the range to search for the specified style
    Set rng = doc.Content
    
    ' Find the text with the specified style
    With rng.Find
        .ClearFormatting
        .style = styleName
        .Execute
    End With
    
    ' Check if the style is found
    If rng.Find.found Then
        ' Select the found range
        rng.Select
        MsgBox "Text with style '" & styleName & "' has been selected.", vbInformation
    Else
        MsgBox "No text with style '" & styleName & "' found.", vbExclamation
    End If
End Sub

Sub SelectTextByStylea()
    Dim doc As Document
    Dim rng As range
    Dim styleName As String
    Dim keywordRange As range
    
    ' Set the style name
    styleName = "ABKW"
    
    ' Check if a document is open
    If Documents.count = 0 Then
        MsgBox "No document is open.", vbExclamation
        Exit Sub
    End If
    
    ' Set the active document
    Set doc = ActiveDocument
    
    ' Set the range to search for the specified style
    Set rng = doc.Content
    
    ' Find the text with the specified style
    With rng.Find
        .ClearFormatting
        .style = styleName
        .Execute
    End With
    
    ' Check if the style is found
    If rng.Find.found Then
        ' Store the found range in a variable
        Set keywordRange = rng.Duplicate
        
        ' Move the range to the end of the found text
        keywordRange.MoveEndUntil "Keywords:"
        
        ' Select the found range
        keywordRange.Select
        
        MsgBox "Text with style '" & styleName & "' has been selected.", vbInformation
    Else
        MsgBox "No text with style '" & styleName & "' found.", vbExclamation
    End If
End Sub



Sub FindPatternWithWildcards()
    Dim rng As range
    Set rng = ActiveDocument.Content ' Change to the desired range if necessary
    
    With rng.Find
        .ClearFormatting
        .text = ": ([a-z])"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .style = "AT"
        Do While .Execute
            ' Convert the found pattern to uppercase
            rng.text = UCase(rng.text)
        Loop
    End With
End Sub
Sub ConvertH1ToTitleCase()
    Dim doc As Document
    Dim rng As range
    Dim para As paragraph
    Dim wordArr() As String
    Dim word As String
    Dim ignoreWords() As Variant
    Dim i As Integer

    ' Words to ignore in title case
    ignoreWords = Array("a", "an", "the", "and", "but", "or", "in", "on", "at", "for", "to", "of")

    Set doc = ActiveDocument

    For Each para In doc.Paragraphs
        ' Check if the paragraph style is 'H1'
        If para.style = "H1" Then
            Set rng = para.range
            wordArr = Split(rng.text, " ")

            ' Loop through each word and apply title case rules
            For i = LBound(wordArr) To UBound(wordArr)
                word = Trim(wordArr(i))

                ' Ignore small words unless they are the first word
                If i = LBound(wordArr) Or UCase(word) <> UCase(ignoreWords(i)) Then
                    ' Capitalize the first letter of the word
                    If Len(word) > 0 Then
                        Mid(word, 1, 1) = UCase(Mid(word, 1, 1))
                    End If
                End If

                wordArr(i) = word
            Next i

            ' Join the words back together and update the range
            rng.text = Join(wordArr, " ")
        End If
    Next para
End Sub
Sub Nameanddatechecklist()
    Dim findList() As Variant
    Dim replaceList() As Variant
    Dim i As Long
    Dim count As Long
    
    ' Characters and phrases to find
    findList = Array("  ", "..", ". .", ". ,", ", ,", "?.", ");", "((", "))", "( ", " )", " ;", "et al ", "et al,", ", et al", "dsm", "ibid", "in press", "n.d.", "forthcoming", "under review", "blind", "this issue", "this volume", "personal communication", "Suppl")
    
    ' Colors to highlight
    replaceList = Array(wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed)
    
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
Sub Nameanddatecheck1()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim rng As range
    Set rng = doc.Content
    
    Dim found As Boolean
    found = False
    
    With rng.Find
        '.ClearFormatting
        .text = "[0-9]{4}, "
        .Format = False
        .MatchWildcards = True
        
        Do While .Execute
            rng.HighlightColorIndex = wdRed ' Highlight the found text
            rng.Collapse wdCollapseEnd ' Move to the end of the found range
            found = True
        Loop
    End With
    
    If found Then
        MsgBox "^#^#^#^#, Pattern Found and highlighted.", vbInformation, "Result"
    Else
        MsgBox "^#^#^#^#,  Pattern not found.", vbExclamation, "Result"
    End If
End Sub
Sub Nameanddatecheck2()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim rng As range
    Set rng = doc.Content
    
    Dim found As Boolean
    found = False
    
    With rng.Find
        '.ClearFormatting
        .text = "[A-Za-z] [0-9]{4}"
        .Format = False
        .MatchWildcards = True
        
        Do While .Execute
            rng.HighlightColorIndex = wdRed ' Highlight the found text
            rng.Collapse wdCollapseEnd ' Move to the end of the found range
            found = True
        Loop
    End With
    
    If found Then
        MsgBox "^$ ^#^#^#^# Pattern Found and highlighted.", vbInformation, "Result"
    Else
        MsgBox "^$ ^#^#^#^# Pattern not found.", vbExclamation, "Result"
    End If
End Sub
Sub Nameanddatecheck3()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim rng As range
    Set rng = doc.Content
    
    Dim found As Boolean
    found = False
    
    With rng.Find
        '.ClearFormatting
        .text = "et al. ^#^#^#^#"
        .Format = False
        .MatchWildcards = False
        
        Do While .Execute
            rng.HighlightColorIndex = wdRed ' Highlight the found text
            rng.Collapse wdCollapseEnd ' Move to the end of the found range
            found = True
        Loop
    End With
    
    If found Then
        MsgBox "et al. ^#, Pattern Found and highlighted.", vbInformation, "Result"
    Else
        MsgBox "et al. ^#,  Pattern not found.", vbExclamation, "Result"
    End If
End Sub
Sub Authornamechecklist()

Call Nameanddatechecklist

Call Nameanddatecheck1

Call Nameanddatecheck2

Call Nameanddatecheck3

End Sub
Sub ConvertATStyleToTitleCase()
    Dim rng As range
    Dim sentence As String
    Dim words() As String
    Dim word As Variant
    Dim exceptions() As Variant
    Dim firstWord As Boolean
    
    ' Define the style name for which you want to apply title case (change 'AT' to your desired style name)
    Const targetStyle As String = "AT"
    
    ' Add any exceptions here (words that should not be title case)
    exceptions = Array("a", "an", "the", "and", "but", "or", "nor", "for", "so", "yet", "as", "at", "by", "in", "of", "on", "to", "up", "with")
    
    ' Loop through each paragraph in the active document
    For Each rng In ActiveDocument.StoryRanges
        Do
            If rng.style.NameLocal = targetStyle Then
                ' Store the paragraph text and split it into words
                sentence = Trim(rng.text)
                words = Split(sentence, " ")
                
                ' Loop through each word and apply title case
                firstWord = True
                For Each word In words
                    ' Check if the word is in the exceptions list or should be capitalized regardless
                    If Not IsInArray(LCase(word), exceptions) Or firstWord Then
                        rng.Find.ClearFormatting
                        rng.Find.text = word
                        rng.Find.Replacement.ClearFormatting
                        rng.Find.Replacement.text = StrConv(word, vbProperCase)
                        rng.Find.Execute Replace:=wdReplaceAll
                    End If
                    firstWord = False
                Next word
            End If
            Set rng = rng.NextStoryRange
        Loop Until rng Is Nothing
    Next rng
    
    MsgBox "Title case applied to style '" & targetStyle & "' successfully!", vbInformation
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    ' Function to check if a string is in the array
    Dim element As Variant
    On Error GoTo ErrorHandler
    
    For Each element In arr
        If element = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
    Exit Function
    
ErrorHandler:
    IsInArray = False
End Function

Sub ConvertH1ToTitleaaCase()
    Dim doc As Document
    Dim rng As range
    Dim para As paragraph

    Set doc = ActiveDocument

    For Each para In doc.Paragraphs
        ' Check if the paragraph style is 'H1'
        If para.style = "H1" Then
            Set rng = para.range
            ' Convert the text to title case
            rng.Case = wdTitleWord
        End If
    Next para
End Sub
Sub ConvertH1ToTitleCaseWithIgnore()
    Dim doc As Document
    Dim rng As range
    Dim para As paragraph
    Dim prepositions As Variant
    Dim word As Variant

    ' List of common prepositions (add more as needed)
    prepositions = Array("in", "on", "at", "to", "with", "from", "of", "by", "about", "above", "below", "under", "between", "among", "through", "into", "onto", "upon")

    Set doc = ActiveDocument

    For Each para In doc.Paragraphs
        ' Check if the paragraph style is 'Heading 1'
        If para.style = "H1" Then
            Set rng = para.range
            ' Convert the text to title case
            For Each word In Split(rng.text, " ")
                If Not IsPreposition(LCase(word), prepositions) Then
                    ' Convert the word to title case
                    rng.MoveStart wdCharacter, Len(word) + 1
                    rng.MoveEnd wdCharacter, -1
                    rng.text = UCase(Left(word, 1)) & LCase(Mid(word, 2))
                End If
            Next word
        End If
    Next para
End Sub

Function IsPreposition(word As String, prepositions As Variant) As Boolean
    Dim i As Long
    For i = LBound(prepositions) To UBound(prepositions)
        If word = prepositions(i) Then
            IsPreposition = True
            Exit Function
        End If
    Next i
    IsPreposition = False
End Function
Sub TitleCase()
    Dim lcList As String
    Dim wrd As Integer
    Dim sTest As String

    ' list of lowercase words, surrounded by spaces
    lcList = " is a and had that the an and as but for if nor or so yet a an the as at by for in of off on per to up via vs are"

    Selection.range.Case = wdTitleWord

    For wrd = 2 To Selection.range.words.count
        sTest = Trim(Selection.range.words(wrd))
        sTest = " " & LCase(sTest) & " "
        If InStr(lcList, sTest) Then
            Selection.range.words(wrd).Case = wdLowerCase
        End If
    Next wrd
End Sub
Sub TitleCaseAT()
    Dim lcList As String
    Dim wrd As Integer
    Dim sTest As String

    ' list of lowercase words, surrounded by spaces
    lcList = " is a and had that an and as but for if nor or so yet a an the as at by for in of off on per to up via vs are"

    For Each p In ActiveDocument.Paragraphs
        If p.style = "H1" Then
            p.range.Case = wdTitleWord

            For wrd = 2 To p.range.words.count
                sTest = Trim(p.range.words(wrd))
                sTest = " " & LCase(sTest) & " "
                If InStr(lcList, sTest) Then
                    p.range.words(wrd).Case = wdLowerCase
                End If
            Next wrd
        End If
    Next p
End Sub
Sub TitleCaseAMA()
    Dim lcList As String
    Dim wrd As Integer
    Dim sTest As String
    Dim p As paragraph
    
    ' List of lowercase words, surrounded by spaces
    lcList = " is a and had that an and as but for if nor or so yet a an the as at by for in of off on per to up via vs are"
    
    For Each p In ActiveDocument.Paragraphs
        Select Case p.style
            Case "H1", "H2"
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

Sub FindOccurrencesInStyleCL()
    Dim doc As Document
    Dim rng As range
    Dim count As Integer
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Set the range to search the entire document
    Set rng = doc.Content
    
    ' Reset the counter
    count = 0
    
    ' Loop through each occurrence of the text in the specified style
    With rng.Find
        .ClearFormatting
        .style = doc.Styles("CL")
        .text = "ABOUT HERE]"
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            ' Increment the counter
            count = count + 1
        Loop
    End With
    
    ' Display the count of occurrences
    MsgBox "Number of occurrences of 'ABOUT HERE]' in style 'CL': " & count, vbInformation, "Occurrences Count"
    
    ' Clean up
    Set rng = Nothing
    Set doc = Nothing
End Sub
Sub Delete_InsertFigure()
    Dim doc As Document
    Dim rng As range
    Dim count As Integer
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Set the range to search the entire document
    Set rng = doc.Content
    
    ' Reset the counter
    'count = 0
    
    ' Loop through each occurrence of the text in the specified style
    With rng.Find
        .ClearFormatting
        .style = doc.Styles("CL")
        .text = "ABOUT HERE]"
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            ' Delete the entire paragraph
            rng.Paragraphs(1).range.Delete
            ' Increment the counter
            'count = count + 1
        Loop
    End With
    
    ' Display the count of deleted occurrences
    MsgBox "Insert Figure Deleted"
    
    ' Clean up
    Set rng = Nothing
    Set doc = Nothing
End Sub

Sub FindAndHighlightNumbersWithEndash()
    Dim rng As range
    Dim findText As String
    
    ' Set the range where you want to search for numbers with endash
    Set rng = ActiveDocument.Content
    
    ' Set the text to find (number with endash)
    findText = "[0-9]–[0-9]"
    
    ' Loop through the range and find all instances of the specified text
    With rng.Find
        .ClearFormatting
        .text = findText
        .Format = True
        .style = "REF"
        .MatchWildcards = True ' Enable wildcard matching
        .Wrap = wdFindStop ' Stop searching when the end of the document is reached
        Do While .Execute
            If rng.Find.found Then
                ' Highlight the found text
                rng.HighlightColorIndex = wdGreen
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub
Sub nameanddatec5()
    Dim rng As range
    Dim findText As String
    
    ' Set the range where you want to search for "p." followed by a single-digit number
    Set rng = ActiveDocument.Content
    
    ' Set the text to find ("p." followed by a single-digit number)
    findText = "p\.[0-9]{1}"
    
    ' Loop through the range and find all instances of the specified text
    With rng.Find
        .ClearFormatting
        .text = findText
        .Format = False
        .MatchWildcards = True ' Enable wildcard matching
        .Wrap = wdFindStop ' Stop searching when the end of the document is reached
        Do While .Execute
            If rng.Find.found Then
                ' Highlight the found text (including "p." and the single-digit number)
                rng.HighlightColorIndex = wdYellow
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub
Sub APARefa1()
    Dim rng As range
    Dim findText As String
    
    ' Set the range where you want to search for numbers with endash
    Set rng = ActiveDocument.Content
    
    ' Set the text to find (number with endash)
    findText = "-([a-z])"
    
    ' Loop through the range and find all instances of the specified text
    With rng.Find
        '.ClearFormatting
        .text = findText
        .Format = True
        .style = "REF"
        '.MatchWildcards = True ' Enable wildcard matching
        .Wrap = wdFindStop ' Stop searching when the end of the document is reached
        Do While .Execute
            If rng.Find.found And rng.style <> "url" Then ' Check if the style is not "url"
                ' Highlight the found text
                rng.HighlightColorIndex = wdRed
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub
Sub APARefa2()
    Dim rng As range
    Dim findText As String
    
    ' Set the range where you want to search for numbers with endash
    Set rng = ActiveDocument.Content
    
    ' Set the text to find (number with endash)
    findText = "- ([a-z])"
    
    ' Loop through the range and find all instances of the specified text
    With rng.Find
        .ClearFormatting
        .text = findText
        .Format = True
        .style = "REF"
        .MatchWildcards = True ' Enable wildcard matching
        .Wrap = wdFindStop ' Stop searching when the end of the document is reached
        Do While .Execute
            If rng.Find.found Then
                ' Highlight the found text
                rng.HighlightColorIndex = wdRed
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub
Sub APARefa3()
    Dim rng As range
    Dim findText As String
    
    ' Set the range where you want to search for numbers with endash
    Set rng = ActiveDocument.Content
    
    ' Set the text to find (number with endash)
    findText = "^=([a-z])"
    
    ' Loop through the range and find all instances of the specified text
    With rng.Find
        '.ClearFormatting
        .text = findText
        .Format = True
        .style = "REF"
        '.MatchWildcards = True ' Enable wildcard matching
        .Wrap = wdFindStop ' Stop searching when the end of the document is reached
        Do While .Execute
            If rng.Find.found And rng.style <> "url" Then ' Check if the style is not "url"
                ' Highlight the found text
                rng.HighlightColorIndex = wdRed
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub
Sub APARefa4()
    Dim rng As range
    Dim findText As String
    
    ' Set the range where you want to search for numbers with endash
    Set rng = ActiveDocument.Content
    
    ' Set the text to find (number with endash)
    findText = "^= ([a-z])"
    
    ' Loop through the range and find all instances of the specified text
    With rng.Find
        .ClearFormatting
        .text = findText
        .Format = True
        .style = "REF"
        .MatchWildcards = True ' Enable wildcard matching
        .Wrap = wdFindStop ' Stop searching when the end of the document is reached
        Do While .Execute
            If rng.Find.found Then
                ' Highlight the found text
                rng.HighlightColorIndex = wdRed
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub
Sub APARefa5()
    Dim rng As range
    Dim findText As String
    
    ' Set the range where you want to search for numbers with endash
    Set rng = ActiveDocument.Content
    
    ' Set the text to find (number with endash)
    findText = "^$^p"
    
    ' Loop through the range and find all instances of the specified text
    With rng.Find
        '.ClearFormatting
        .text = findText
        .Format = True
        .style = "REF"
        '.MatchWildcards = True ' Enable wildcard matching
        .Wrap = wdFindStop ' Stop searching when the end of the document is reached
        Do While .Execute
            If rng.Find.found And rng.style <> "url" Then ' Check if the style is not "url"
                ' Highlight the found text
                rng.HighlightColorIndex = wdRed
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub
Sub APARefa6()
    Dim rng As range
    Dim findText As String
    
    ' Set the range where you want to search for numbers with endash
    Set rng = ActiveDocument.Content
    
    ' Set the text to find (number with endash)
    findText = "^#^p"
    
    ' Loop through the range and find all instances of the specified text
    With rng.Find
        '.ClearFormatting
        .text = findText
        .Format = True
        .style = "REF"
        '.MatchWildcards = True ' Enable wildcard matching
        .Wrap = wdFindStop ' Stop searching when the end of the document is reached
        Do While .Execute
            If rng.Find.found And rng.style <> "url" Then ' Check if the style is not "url"
                ' Highlight the found text
                rng.HighlightColorIndex = wdRed
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub
Sub APARefa7()
    Dim rng As range
    Dim findText As String
    
    ' Set the range where you want to search for numbers with endash
    Set rng = ActiveDocument.Content
    
    ' Set the text to find (number with endash)
    findText = "^#-^#"
    
    ' Loop through the range and find all instances of the specified text
    With rng.Find
        '.ClearFormatting
        .text = findText
        .Format = True
        .style = "REF"
        '.MatchWildcards = True ' Enable wildcard matching
        .Wrap = wdFindStop ' Stop searching when the end of the document is reached
        Do While .Execute
            If rng.Find.found And rng.style <> "url" Then ' Check if the style is not "url"
                ' Highlight the found text
                rng.HighlightColorIndex = wdRed
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub
Sub APARefa8()
    Dim rng As range
    Dim findText As String
    
    ' Set the range where you want to search for numbers with endash
    Set rng = ActiveDocument.Content
    
    ' Set the text to find (number with endash)
    findText = ""
    
    ' Loop through the range and find all instances of the specified text
    With rng.Find
        .ClearFormatting
        .text = findText
        .Format = True
        .Font.Italic = True
        .style = "REF"
        '.MatchWildcards = True ' Enable wildcard matching
        .Wrap = wdFindStop ' Stop searching when the end of the document is reached
        Do While .Execute
            If rng.Find.found Then
                ' Highlight the found text
                rng.HighlightColorIndex = wdBrightGreen
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub
Sub APARefa9()
    Dim rng As range
    Dim findText As String
    
    ' Set the range where you want to search for numbers with endash
    Set rng = ActiveDocument.Content
    
    ' Set the text to find (number with endash)
    findText = "p.^#"
    
    ' Loop through the range and find all instances of the specified text
    With rng.Find
        '.ClearFormatting
        .text = findText
        .Format = True
        .style = "REF"
        '.MatchWildcards = True ' Enable wildcard matching
        .Wrap = wdFindStop ' Stop searching when the end of the document is reached
        Do While .Execute
            If rng.Find.found And rng.style <> "doino" Then ' Check if the style is not "url"
                ' Highlight the found text
                rng.HighlightColorIndex = wdRed
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
End Sub
Sub APAREFcheck()

Call APARefa1
Call APARefa2
Call APARefa3
Call APARefa4
Call APARefa5
Call APARefa6
Call APARefa7
Call APARefa8
Call APARefa9
Call REFAftercolon2
End Sub
Sub RRHdot()
With ActiveDocument.Content.Find
 .ClearFormatting
 .text = "."
 .style = "RRH"
 .Replacement.ClearFormatting
 .Replacement.text = ""
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub
Sub CountWordsWithStyle()
    On Error Resume Next ' Enable error handling
    
    Dim doc As Document
    Dim rng As range
    Dim wordCount As Integer
    Dim word As range
    
    ' Set the document and range variables
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "No active document found.", vbExclamation, "Error"
        Exit Sub
    End If
    
    Set rng = doc.Content
    If rng Is Nothing Then
        MsgBox "Document content range not found.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Reset word count
    wordCount = 0
    
    ' Loop through each word in the range
    For Each word In rng.words
        ' Check if the word has the specified style
        If word.style = "ABKW" Then
            wordCount = wordCount + 1
        End If
    Next word
    
    ' Display the word count
    MsgBox "Number of words with style 'ABKW': " & wordCount, vbInformation, "Word Count"
    
    On Error GoTo 0 ' Disable error handling
End Sub

Sub CountWordaasInStyle()
    Dim doc As Document
    Dim styleName As String
    Dim wordCount As Long
    Dim word As range
    
    ' Set the style name to count words for
    styleName = "ABKW"
    
    ' Check if the style exists
    If Not StyleExists(styleName) Then
        MsgBox "Style '" & styleName & "' not found in the document."
        Exit Sub
    End If
    
    ' Set the active document
    Set doc = ActiveDocument
    
    ' Initialize word count
    wordCount = 0
    
    ' Loop through each word and check style
    For Each word In doc.words
        If word.style = styleName Then
            wordCount = wordCount + 1
        End If
    Next word
    
    ' Display the word count
    MsgBox "Word count in style '" & styleName & "': " & wordCount
End Sub

Function StyleExists(styleName As String) As Boolean
    Dim s As style
    On Error Resume Next
    Set s = ActiveDocument.Styles(styleName)
    On Error GoTo 0
    StyleExists = Not (s Is Nothing)
End Function
Sub ConvertToSentenceCase()
    Dim selectedRange As range
    Dim sentence As String
    Dim i As Long
    
    ' Check if text is selected
    If Selection.Type = wdSelectionIP Then
        MsgBox "No text selected.", vbExclamation, "Selection Error"
        Exit Sub
    End If
    
    ' Get the selected text and convert to sentence case
    Set selectedRange = Selection.range
    sentence = selectedRange.text
    sentence = StrConv(sentence, vbProperCase)
    
    ' Correct the first character of each sentence to uppercase
    For i = 2 To Len(sentence)
        If Mid(sentence, i - 1, 1) = "." Or Mid(sentence, i - 1, 1) = "!" Or Mid(sentence, i - 1, 1) = "?" Then
            Mid(sentence, i, 1) = UCase(Mid(sentence, i, 1))
        End If
    Next i
    
    ' Update the selected range with the sentence case text
    selectedRange.text = sentence
End Sub

Sub TitleCaseWithExceptions()
    Dim lcList As String
    Dim wrd As Integer
    Dim sTest As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' List of lowercase words, surrounded by spaces
    lcList = " is a and had that an and as but for if nor or so yet a an the as at by for in of off on per to up via vs are"

    ' Set up regular expression pattern to match words with four or more letters
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = "\b\w{4,}\b"

    Selection.range.Case = wdTitleWord

    ' Iterate through the matches
    For Each match In regex.Execute(Selection.range.text)
        sTest = match.Value
        sTest = " " & LCase(sTest) & " "
        If InStr(lcList, sTest) Then
            Selection.range.Find.Execute findText:=sTest, MatchCase:=False, ReplaceWith:=sTest, Replace:=wdReplaceAll
        End If
    Next match
End Sub
Sub CheckAPAStyleReference()
    Dim refText As String
    Dim regex As Object
    Dim match As Object
    
    ' Get the selected text (reference)
    refText = Selection.text
    
    ' Create a regular expression pattern for a basic APA book reference
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = "^[A-Z][a-zA-Z',.\s]+\. \([12]\d{3}\)\. [A-Z][a-zA-Z\s]+\. [A-Z][a-zA-Z\s]+$"
    
    ' Check if the reference matches the pattern
    If regex.Test(refText) Then
        MsgBox "This reference appears to follow basic APA style."
    Else
        MsgBox "This reference does not follow basic APA style."
    End If
End Sub

Sub SelectMultipleStyles()
    Dim rng As range
    Dim styleName As String
    Dim targetStyles() As Variant
    Dim i As Integer
    
    ' Define an array of style names you want to select
    ' Add more style names or modify this list as needed
    targetStyles = Array("H1", "H2")
    
    ' Initialize the range object
    Set rng = ActiveDocument.Content
    
    ' Loop through the array of style names
    For i = LBound(targetStyles) To UBound(targetStyles)
        styleName = targetStyles(i)
        
        ' Clear any previous selections
        rng.Select
        
        ' Search for text with the specified style
        With rng.Find
            .ClearFormatting
            .style = ActiveDocument.Styles(styleName)
            .text = ""
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Execute
        End With
        
        ' If text with the style is found, select it
        If rng.Find.found Then
            rng.Select
        End If
    Next i
End Sub

Sub Test()
Dim i As Long
Dim FileName As String
Application.FileDia1og(msoEi1eDia10gOpen).Show
FileName Application.Fi1eDia10g(msoFi1eDia10gOpen).Selectedltems(1)
ScreeUpdating -i(Fa1se)
Line2: On Error GoTo Linel
Documents.Open FileName, , True, , i & ""
MsgBox "Password i'"
Application.ScreenOpdating -True
Exit Sub
Linel: i = i + 1
Resume Line2
ScreeUpdating = True
End Sub

Sub teaast()
    Dim i As Long
    Dim FileName As String
    
    ' Display the File Open dialog to select a file
    With Application.FileDialog(msoFileDialogOpen)
        If .Show = -1 Then
            FileName = .SelectedItems(1)
        Else
            Exit Sub ' User canceled the dialog
        End If
    End With
    
    Application.ScreenUpdating = False
    i = 1
    
Line2:
    On Error GoTo Line1
    
    ' Open the document with a password (if needed)
    Documents.Open FileName, , True, , "" & i
    
    ' If no error occurs, the password is correct
    MsgBox "Password is '" & i & "'"
    
    Application.ScreenUpdating = True
    Exit Sub
    
Line1:
    i = i + 1
    Resume Line2
End Sub
Sub RunPythonScript()
    Dim pythonPath As String
    Dim scriptPath As String
    Dim cmd As String

    ' Set the path to your Python executable.
    pythonPath = "C:\Program Files\Python311\python.exe" ' Replace with your Python executable path.

    ' Set the path to your Python script.
    scriptPath = "C:\Users\Admin\Downloads\sentencecase.py"
    ' Debugging: Print paths to the Immediate Window.
    Debug.Print "Python Path: " & pythonPath
    Debug.Print "Script Path: " & scriptPath
    
    ' Build the command to run the Python script.
    cmd = pythonPath & " " & scriptPath
    
    ' Debugging: Print the full command to the Immediate Window.
    Debug.Print "Command: " & cmd


    ' Run the Python script.
    Call Shell(cmd, vbNormalFocus)
End Sub
Sub TABLE()
With ActiveDocument.Content.Find
 .ClearFormatting
 .text = "Table"
 .style = "CPB"
 .MatchCase = True
 .Replacement.ClearFormatting
 .Replacement.text = "TABLE"
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub
Sub tablecolen()
With ActiveDocument.Content.Find
 .ClearFormatting
 .text = "."
 .style = "CPB"
 .MatchCase = True
 .Replacement.ClearFormatting
 .Replacement.text = ":"
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub
Sub UKplosone()
With ActiveDocument.Content.Find
 .ClearFormatting
 .text = "Plos One"
 .style = "stl"
 .MatchCase = False
 .Replacement.ClearFormatting
 .Replacement.text = "PLoS ONE"
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub
Sub APAplosone()
With ActiveDocument.Content.Find
 .ClearFormatting
 .text = "Plos One"
 .style = "stl"
 .MatchCase = False
 .Replacement.ClearFormatting
 .Replacement.text = "PLOS ONE"
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub
Sub Addemspace()
    Dim doc As Document
    Dim para As paragraph
    Dim styleName As String
    Dim emSpace As String
    
    ' Set the style name and em-space character
    styleName = "CPB"
    emSpace = ChrW(&H2003) ' Em-space Unicode character
    
    ' Check if the document is open
    If Documents.count > 0 Then
        Set doc = ActiveDocument
    Else
        MsgBox "No document is open!", vbExclamation
        Exit Sub
    End If
    
    ' Loop through all paragraphs in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph has the specified style
        If para.style.NameLocal = styleName Then
            ' Check if the paragraph has a line break at the end
            If para.range.Characters.Last = vbCr Then
                ' Add an em-space character before the line break
                para.range.Characters(para.range.Characters.count - 1).InsertAfter emSpace
            End If
        End If
    Next para
End Sub

Sub AFauthornote()
With ActiveDocument.Content.Find
 .ClearFormatting
 .text = "AUTHORS' NOTE."
 .style = "AF"
 .MatchCase = True
 .Replacement.ClearFormatting
 .Replacement.text = "Authors' Note:"
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub
Sub InsertTextBeforeStyle()
    Dim doc As Document
    Dim para As paragraph
    Dim styleName As String
    Dim insertText As String
    
    ' Set the style name and the text to insert
    styleName = "TY"
    insertText = "[TS: PLEASE SET THE FORMATTING OF LRH, RRH, AUTHORS’ NOTE, ARTICLE TITLE, AUTHOR NAMES, HEADINGS (H1, H2, H3), FIGURES, TABLES, AND REFERENCES AS PER CJB STYLE GUIDE.]" & vbCr
   
    ' Check if the document is open
    If Documents.count > 0 Then
        Set doc = ActiveDocument
    Else
        MsgBox "No document is open!", vbExclamation
        Exit Sub
    End If
    
    ' Loop through all paragraphs in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph has the specified style
        If para.style.NameLocal = styleName Then
            ' Insert the specified text before the paragraph
            
            para.range.InsertBefore insertText
            para.range.style = doc.Styles("CL")
        End If
    Next para
End Sub
Sub FindReplaceEtAl()
    Dim rng As range
    Dim findText As String
    Dim replaceText As String

    ' Set the text to find and replace
    findText = "et al."
    replaceText = "et al."

    ' Loop through each story in the document
    For Each rng In ActiveDocument.StoryRanges
        ' Find the text and apply italic formatting
        With rng.Find
            .ClearFormatting
            .text = findText
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            ' Execute the find operation
            Do While .Execute
                ' Check if the found text is not already in italic format
                If Not rng.Font.Italic Then
                    ' Apply italic formatting
                    rng.Font.Italic = True
                End If
            Loop
        End With
    Next rng
End Sub
Sub HighlightAndUppercase()
    Dim searchTerm As String
    Dim rng As range
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    ' Set the search term with wildcard for highlighting
    searchTerm = ": [a-z]"
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Clear any existing highlighting
    rng.HighlightColorIndex = wdNoHighlight
    
    ' Find and highlight the search term with the specified style
    With rng.Find
        .text = searchTerm
        .style = ActiveDocument.Styles("REF")
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        
        ' Loop through each found item
        Do While .Execute
            rng.HighlightColorIndex = wdRed
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    ' Create a regular expression object for uppercase conversion
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Set the pattern to match text after a colon and lowercase letters
    regex.pattern = ": ([a-z])"
    
    ' Set the document range
    Set rng = ActiveDocument.range
    
    ' Find all matches in the document range
    Set matches = regex.Execute(rng.text)
    
    ' Loop through each match and apply formatting
    For Each match In matches
        rng.Start = match.FirstIndex
        rng.End = match.FirstIndex + match.Length
        rng.style = "REF"
        rng.text = UCase(rng.text)
    Next match
End Sub
Sub CheckHyperlinksaaa()
    Dim doc As Document
    Dim hyperlink As hyperlink
    
    Set doc = ActiveDocument
    
    For Each hyperlink In doc.Hyperlinks
        If hyperlink.address <> "" Then
            ' Temporarily store the original formatting of the hyperlink range
            Dim originalRange As range
            Set originalRange = hyperlink.range.Duplicate
            
            ' Apply formatting to the hyperlink based on whether it is working or not
            hyperlink.range.HighlightColorIndex = wdBrightGreen ' Green color for working links
            
            Dim isWorking As Boolean
            isWorking = TestHyperlink(hyperlink.address)
            
            If Not isWorking Then
                ' Hyperlink is not working
                hyperlink.range.InsertAfter "[AQ: Please note that the given URL in 'XXXXXX' does not lead to the desired web page. Please provide an active URL.]"
                hyperlink.range.HighlightColorIndex = wdRed ' Red color for invalid links
            End If
            
            ' Restore the original formatting of the hyperlink range
            originalRange.Copy
            hyperlink.range.Paste
        End If
    Next hyperlink
    MsgBox "Hyperlink checking completed!", vbInformation
End Sub
Sub DeleteAllDOIParagraphs()
    Dim doc As Document
    Dim para As paragraph
    Dim searchString As String
    searchString = "https://doi.org/"

    Set doc = ActiveDocument

    For Each para In doc.Paragraphs
        If InStr(para.range.text, searchString) > 0 Then
            para.range.Select
            Selection.Delete
        End If
    Next para
End Sub
Sub ConvertURLTextsToHyperlinksInDoc()
  Dim objDoc As Document
 
  Set objDoc = ActiveDocument
 
  word.Options.AutoFormatReplaceHyperlinks = True
  objDoc.range.AutoFormat
End Sub
Sub ConvertURLTextsToHyperlinksInDocaa()
  Dim objDoc As Document
  Dim rng As range

  Set objDoc = ActiveDocument
  Set rng = objDoc.Content ' You can adjust the range as needed

  ' Enable the AutoFormat option to replace plain text with hyperlinks
  Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = True

  ' Apply the AutoFormat to the specified range
  rng.AutoFormat
End Sub

Sub ConvertURLToHyperlinksInREF()
  Dim objDoc As Document
  Dim rng As range
  Dim styleName As String

  ' Set the name of the style you want to target (e.g., "REF")
  styleName = "REF"

  Set objDoc = ActiveDocument
  Set rng = objDoc.Content ' You can adjust the range as needed

  ' Loop through the document and check if the text is in the specified style
  For Each par In rng.Paragraphs
    If par.style = styleName Then
      ' Enable the AutoFormat option to replace plain text with hyperlinks
      Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = True
      ' Apply the AutoFormat to the paragraph
      par.range.AutoFormat
    End If
  Next par
End Sub

Sub RemoveHyperlinksREF()
    Dim hyperlink As hyperlink
    Dim doc As Document
    Dim styleName As String
    styleName = "REF" ' Change this to the desired style name
    
    Set doc = ActiveDocument
    
    ' Repeat the loop until there are no hyperlinks with the specified style left
    Do While HyperlinksByStyleCount(doc, styleName) > 0
        For Each hyperlink In doc.Hyperlinks
            If hyperlink.range.style = styleName Then
                hyperlink.Delete
            End If
        Next hyperlink
    Loop
    MsgBox "Hyperlinks with style '" & styleName & "' Removed"
End Sub

Function HyperlinksByStyleCount(doc As Document, styleName As String) As Long
    Dim count As Long
    Dim hyperlink As hyperlink
    
    count = 0
    
    For Each hyperlink In doc.Hyperlinks
        If hyperlink.range.style = styleName Then
            count = count + 1
        End If
    Next hyperlink
    
    HyperlinksByStyleCount = count
End Function
Sub ExpandJournalNames()
    ' Define the abbreviation and its expanded form as key-value pairs in a dictionary
    Dim journalNames As Object
    Set journalNames = CreateObject("Scripting.Dictionary")
    
    ' Add journal abbreviations and their expansions to the dictionary
    journalNames.Add "JACS", "Journal of the American Chemical Society"
    journalNames.Add "Nature", "Nature Publishing Group"
    ' Add more journal abbreviations and expansions as needed
    
    ' Loop through each abbreviation in the document and replace it with its expanded form
    For Each abbr In journalNames.keys
        With Selection.Find
            .ClearFormatting
            .text = abbr
            .Replacement.text = journalNames(abbr)
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next abbr
End Sub

Sub AddAsteriskAtEnd()
'removes italic, use wildcard method
    Dim para As paragraph
    Dim rng As range
    
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph starts with an asterisk and has the style "REF"
        If Left(para.range.text, 1) = "*" And para.style = "REF" Then
            Set rng = para.range
            rng.MoveEnd wdCharacter, -1 ' Move the end of the range before the paragraph mark
            rng.text = rng.text & "*" ' Add asterisk at the end
        End If
    Next para
End Sub
Sub RemoveAsteriskAtStart()
'removes astriek at start, can use find and replace
    Dim para As paragraph
    Dim asteriskPos As Long
    
    For Each para In ActiveDocument.Paragraphs
        If para.style = "REF" Then
            If Left(para.range.text, 1) = "*" Then
                asteriskPos = InStr(para.range.text, "*")
                If asteriskPos = 1 Then
                    para.range.Characters(1).Delete
                End If
            End If
        End If
    Next para
End Sub
Sub AddAsteriskAtStart()
'add astreik at start if there is astriek at end
    Dim para As paragraph
    Dim doc As Document
    Dim rng As range
    
    ' Set the reference to the active document
    Set doc = ActiveDocument
    
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph contains the desired style ("REF") and an asterisk
        If para.style = "REF" And InStr(para.range.text, "*") > 0 Then
            para.range.InsertBefore "*"
        End If
    Next para
End Sub

Sub AddAsteriskToSelectedParagraphs()
'not necessary now
    Dim para As paragraph
    Dim selectedRange As range
    
    ' Check if there is a selection in the document
    If Selection.Type = wdSelectionNormal Then
        Set selectedRange = Selection.range
        ' Loop through each paragraph in the selected range
        For Each para In selectedRange.Paragraphs
            ' Add an asterisk at the beginning of the paragraph text
            para.range.InsertBefore "*"
        Next para
    Else
        ' If no selection is made, display a message
        MsgBox "Please select the paragraphs you want to modify.", vbExclamation, "No Selection"
    End If
End Sub
Sub astriek()
'not necessary now
    Dim doc As Document
    Dim para As paragraph
    Dim rng As range
    
    Set doc = ActiveDocument
    
    For Each para In doc.Paragraphs
        If para.style = "REF" Then
            Set rng = para.range
            rng.MoveEnd wdCharacter, -1 ' Move the end of the range before the paragraph mark
            rng.text = rng.text & "*"
        End If
    Next para
End Sub

Sub SelectedParagraphsWithAsterisk()
'not necessary now
    Dim para As paragraph
    Dim rng As range

    ' Check if there is selected text
    If Selection.Type = wdSelectionIP Or Selection.Type = wdSelectionNormal Then
        ' Loop through each paragraph in the selected range
        For Each para In Selection.Paragraphs
            ' Check if the paragraph contains an asterisk
            If InStr(para.range.text, "*") > 0 Then
                para.range.InsertBefore "*"
            End If
        Next para
    Else
        ' No text is selected
        MsgBox "Please select some text before running this macro."
    End If
End Sub

Sub AddAsteriskAtEndq()
'not necessary now
    Dim para As paragraph
    
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph starts with an asterisk and has the style "REF"
        If Left(para.range.text, 1) = "*" And para.style = "REF" Then
            Dim paraText As String
            Dim rngStart As Long
            Dim paraLength As Long
            
            paraText = para.range.text
            paraLength = Len(paraText)
            
            ' Get the start position of the paragraph content (exclude paragraph mark)
            rngStart = para.range.Start
            
            ' Preserve formatting by inserting an asterisk just before the paragraph mark
            para.range.Collapse wdCollapseEnd
            para.range.MoveStartUntil CSet:="*"
            para.range.MoveStart wdCharacter, -1
            para.range.text = "*" & paraText & "*"
            
            ' Restore the original formatting for the newly modified text
            ActiveDocument.range(rngStart, rngStart + paraLength).style = "REF"
        End If
    Next para
End Sub
Sub MoveTextByStyle()
    Dim doc As Document
    Dim rng As range
    Dim foundRange As range
    
    ' Define the document
    Set doc = ActiveDocument
    
    ' Define the initial range to search the entire document
    Set rng = doc.Content
    
    ' Set the starting point of the search to the beginning of the document
    rng.Start = 0
    
    ' Loop to find the text with the "AF" style
    Do
        ' Look for the next instance of the "AF" style
        Set foundRange = rng.Duplicate
        With foundRange.Find
            .ClearFormatting
            .style = doc.Styles("AF")
            .text = ""
            .Forward = True
            .Execute
        End With
        
        ' If found, move the text to the desired location (e.g., "AU" style)
        If foundRange.Find.found Then
            foundRange.Cut
            doc.Styles("AU").Paragraphs.Last.range.Collapse Direction:=wdCollapseEnd
            doc.Styles("AU").Paragraphs.Last.range.Paste
        End If
        
        ' Move the search range to the end of the found range
        rng.Start = foundRange.End
        rng.End = doc.Content.End
    Loop While foundRange.Find.found
End Sub

Sub MoveParagraphAFBelowAU()
    Dim styleAF As String
    Dim styleAU As String
    Dim paraAF As word.paragraph
    Dim paraAU As word.paragraph
    Dim foundAU As Boolean

    ' Specify the style names
    styleAF = "AF"
    styleAU = "AU"
    foundAU = False ' Flag to check if AU paragraph is found

    ' Loop through all paragraphs in the document
    For Each paraAF In ActiveDocument.Paragraphs
        If paraAF.style = styleAF Then ' Check for paragraphs with styleAF
            For Each paraAU In ActiveDocument.Paragraphs
                If paraAU.style = styleAU Then ' Check for paragraphs with styleAU
                    ' Move the paragraph with styleAF below the paragraph with styleAU
                    paraAF.range.Cut
                    paraAU.range.Collapse wdCollapseEnd
                    paraAU.range.Paste
                    foundAU = True
                    Exit For ' Exit the loop after moving one AF paragraph below AU
                End If
            Next paraAU
            If foundAU Then Exit For ' Exit the loop after the first move
        End If
    Next paraAF
End Sub
Sub CapitalizeWordAfterColon()
    Application.ScreenUpdating = False
    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .text = ": ([a-z])"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWholeWord = False
        .MatchSoundsLike = False
        .MatchCase = False
        .MatchWildcards = True
        .MatchAllWordForms = False
        With .Replacement.Font
          .AllCaps = True
          .SmallCaps = False
        End With
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Application.ScreenUpdating = True
End Sub
Sub LowercaseAfterColon()

Dim found As Boolean
Dim range As word.range

found = False

Do
    With Selection.Find
        .ClearFormatting
        .text = ": ([A-Z])"
        .style = "atl"
        .Replacement.ClearFormatting
        .Replacement.text = ": \1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    Selection.Find.Execute

    If Selection.Find.found Then
        Selection.range.Case = wdLowerCase
        Selection.range.HighlightColorIndex = wdBrightGreen
        Selection.Collapse Direction:=wdCollapseEnd
        found = True
    Else
        found = False
    End If
Loop Until found = False

End Sub
Sub TYPageBreak()
    Dim pgBreak As range
    For Each pgBreak In ActiveDocument.StoryRanges
        With pgBreak.Find
            .ClearFormatting
            .text = "^m"
            .style = "TY"
            .Forward = True
            .Wrap = wdFindStop
            Do While .Execute
                pgBreak.Select
                Selection.MoveStartUntil (Chr(12))
                Selection.MoveEndUntil (Chr(12))
                Selection.range.style = "Normal"
            Loop
        End With
    Next pgBreak
End Sub
Sub TYPageBreak1()
    Dim pgBreak As range
    For Each pgBreak In ActiveDocument.StoryRanges
        With pgBreak.Find
            .ClearFormatting
            .text = "^m"
            .Forward = True
            .Wrap = wdFindStop
            Do While .Execute
                If pgBreak.Characters.First.Information(wdWithInTable) = False Then
                    pgBreak.Collapse wdCollapseStart
                    pgBreak.MoveEnd wdCharacter, 1
                    pgBreak.style = ActiveDocument.Styles("TY")
                    pgBreak.Collapse wdCollapseEnd
                    pgBreak.MoveStart wdCharacter, -1
                    pgBreak.style = ActiveDocument.Styles("Normal")
                End If
            Loop
        End With
    Next pgBreak
End Sub
Sub RemoveHyperlinksAndSave()
    Dim hyperlink As hyperlink
    Dim doc As Document
    Dim currentPath As String
    Dim newFolderPath As String
    Dim newFilePath As String
    
    ' Get the current path of the active document
    Set doc = ActiveDocument
    currentPath = doc.FullName
    
    ' Extract the folder path from the current document path
    newFolderPath = Left(currentPath, InStrRev(currentPath, "\"))
    
    ' Specify the new folder name where you want to save the document
    newFolderPath = newFolderPath & "Compare\" ' Modify this with your desired folder name
    
    ' Check if the folder exists, create it if it doesn't
    If Dir(newFolderPath, vbDirectory) = "" Then
        MkDir newFolderPath
    End If
    
    ' Repeat the loop until there are no hyperlinks left
    Do While doc.Hyperlinks.count > 0
        For Each hyperlink In doc.Hyperlinks
            hyperlink.Delete
        Next hyperlink
    Loop
    
    ' Save the document with the same name in the new folder
    newFilePath = newFolderPath & doc.Name
    doc.SaveAs2 FileName:=newFilePath
    MsgBox "Hyperlinks Removed and Document Saved in the New Folder"
End Sub
Sub CompareDocument()
 ActiveDocument.Compare Name:="C:\Users\Admin\Downloads\MSS1219592\Compare\MSS1219592_pre.docx", _
 CompareTarget:=wdCompareTargetNew
End Sub

Sub CompareDocumentInSamePath()
    Dim currentPath As String
    Dim preDocPath As String
    
    ' Get the path of the active document
    currentPath = ActiveDocument.Path & "\"
    
    ' Search for a file ending with "pre.docx" in the same directory
    preDocPath = Dir(currentPath & "*pre.docx")
    
    If preDocPath <> "" Then
        ' Compare the active document with the found "pre.docx" file
        ActiveDocument.Compare Name:=currentPath & preDocPath, CompareTarget:=wdCompareTargetNew
    Else
        MsgBox "Pre-document ending with 'pre.docx' not found in the same directory."
    End If
End Sub
Sub ActivateDocument()
    Dim doc As Document
    Dim targetDoc As Document
    Dim docCount As Integer
    
    ' Set the target document name or criteria to identify the document you want to activate
    Dim targetName As String
    targetName = "MSS1219592_CLN.docx" ' Change this to your target document's name
    
    ' Get the count of open documents
    docCount = Documents.count
    
    ' Loop through each open document
    For Each doc In Documents
        ' Check if the document matches the target criteria (e.g., document name)
        If doc.Name = targetName Then
            ' Set the target document
            Set targetDoc = doc
            Exit For
        End If
    Next doc
    
    ' Check if the target document was found
    If Not targetDoc Is Nothing Then
        ' Activate the target document
        targetDoc.Activate
        MsgBox "Document activated: " & targetDoc.Name
    Else
        MsgBox "Document '" & targetName & "' not found among " & docCount & " open documents."
    End If
End Sub
Sub ActivateDocumentCLN()
    Dim doc As Document
    Dim targetDoc As Document
    Dim docCount As Integer
    
    ' Get the count of open documents
    docCount = Documents.count
    
    ' Loop through each open document
    For Each doc In Documents
        ' Check if the document name ends with "CLN.docx"
        If Right(doc.Name, 8) = "CLN.docx" Then
            ' Set the target document
            Set targetDoc = doc
            Exit For
        End If
    Next doc
    
    ' Check if the target document was found
    If Not targetDoc Is Nothing Then
        ' Activate the target document
        targetDoc.Activate
        MsgBox "Document activated: " & targetDoc.Name
    Else
        MsgBox "No document with name ending in 'CLN.docx' found among " & docCount & " open documents."
    End If
End Sub
Sub ActivateDocumentpre()
    Dim doc As Document
    Dim targetDoc As Document
    Dim docCount As Integer
    
    ' Get the count of open documents
    docCount = Documents.count
    
    ' Loop through each open document
    For Each doc In Documents
        ' Check if the document name ends with "CLN.docx"
        If Right(doc.Name, 8) = "pre.docx" Then
            ' Set the target document
            Set targetDoc = doc
            Exit For
        End If
    Next doc
    
    ' Check if the target document was found
    If Not targetDoc Is Nothing Then
        ' Activate the target document
        targetDoc.Activate
        MsgBox "Document activated: " & targetDoc.Name
    Else
        MsgBox "No document with name ending in 'CLN.docx' found among " & docCount & " open documents."
    End If
End Sub
Sub FormatDocument()
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

Sub FindAndReplace(textToFind As String, findStyle As String, replaceText As String)
    With ActiveDocument.Content.Find
        .ClearFormatting
        .text = textToFind
        .style = findStyle
        .Replacement.ClearFormatting
        .Replacement.text = replaceText
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
End Sub

Sub CombineCodes()
    FindAndReplace ".^p", "H3", ""
    FindAndReplace ".^p", "CPB", "^p"
    FindAndReplace ".^p", "CP", "^p"
End Sub

Sub RemoveEndPeriodFromCPStyle()
    Dim para As paragraph
    
    For Each para In ActiveDocument.Paragraphs
        If para.style = "CP" Then
            Dim paraText As String
            paraText = para.range.text
            
            If Right(paraText, 2) = "." & vbCr Then
                para.range.Characters(Len(paraText) - 1).Delete
            End If
        End If
    Next para
End Sub
Sub RemoveEndPeriodFromStyles()
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
Sub FindTextWithStyleAndSelectParagraph()
    Dim doc As Document
    Dim rng As range
    Dim found As Boolean
    Dim styleName As String
    Dim paraRange As range
    
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
            
            
            ' Select the entire paragraph
            paraRange.Select
            
            
            ' Move the range to the end of the paragraph to continue searching
            rng.Collapse Direction:=wdCollapseEnd
        Else
            ' Exit the loop if no more instances are found
            found = False
        End If
    Loop
End Sub



Sub Nameanddatechecklist22()
    Dim findList() As Variant
    Dim replaceList() As Variant
    Dim i As Long
    Dim count As Long
    Dim summary As String
    
    ' Characters and phrases to find
    findList = Array("  ", "..", ". .", ". ,", ", ,", "?.", ");", "((", "))", "( ", " )", " ;", "et al ", "et al,", ", et al", "dsm", "ibid", "in press", "n.d.", "forthcoming", "under review", "blind", "this issue", "this volume", "personal communication", "Suppl", "[0-9]{4}, ")
    
    ' Colors to highlight
    replaceList = Array(wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed)
    
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
        
        ' Store count and phrase in the summary string
        summary = summary & count & " items found for character or phrase: " & findList(i) & vbCrLf
    Next i
    
    ' Display the summary of items found at the end
    MsgBox "Summary of items found:" & vbCrLf & summary, vbInformation, "Highlighting Summary"
End Sub

Sub Nameanddatechecklist24()
    Dim findList() As Variant
    Dim replaceList() As Variant
    Dim i As Long
    Dim count As Long
    Dim summary As String
    
    ' Characters and phrases to find
    findList = Array("  ", "..", ". .", ". ,", ", ,", "?.", ");", "((", "))", "( ", " )", " ;", "et al ", "et al,", ", et al", "dsm", "ibid", "in press", "n.d.", "forthcoming", "under review", "blind", "this issue", "this volume", "personal communication", "Suppl")
    
    ' Colors to highlight
    replaceList = Array(wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed)
    
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
        
        ' Store count and phrase in the summary string if count is greater than zero
        If count > 0 Then
            summary = summary & count & " items found for character or phrase: " & findList(i) & vbCrLf
        End If
    Next i
    
    ' Display the summary of items found (if any) at the end
    If Len(summary) > 0 Then
        MsgBox "Summary of items found:" & vbCrLf & summary, vbInformation, "Highlighting Summary"
    Else
        MsgBox "No items found.", vbInformation, "Highlighting Summary"
    End If
End Sub

Sub Nameanddatechecklist30()
    Dim findList() As Variant
    Dim replaceList() As Variant
    Dim i As Long
    Dim count As Long
    Dim summary As String
    
    ' Characters and phrases to find
    findList = Array("  ", "..", ". .", ". ,", ", ,", "?.", ");", "((", "))", "( ", " )", " ;", "et al ", "et al,", ", et al", "dsm", "ibid", "in press", "n.d.", "forthcoming", "under review", "blind", "this issue", "this volume", "personal communication", "Suppl")
    
    ' Patterns to find
    Dim patternsToFind() As Variant
    patternsToFind = Array("[0-9]{4}, ", "[A-Za-z] [0-9]{4}", "et al. [0-9]")
    
    ' Colors for highlighting
    replaceList = Array(wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed)
    
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
        
        ' Store count and phrase in the summary string if count is greater than zero
        If count > 0 Then
            summary = summary & count & " items found for character or phrase: " & findList(i) & vbCrLf
        End If
    Next i
    
    ' Loop through the patterns and highlight the matches
    For i = LBound(patternsToFind) To UBound(patternsToFind)
        count = 0 ' Reset count for each pattern
        
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = patternsToFind(i)
            .MatchWildcards = True ' Use wildcards for pattern matching
            .Forward = True
            .Wrap = wdFindContinue
            
            ' Highlight the matches with the corresponding color and count the number of items found
            Do While .Execute
                Selection.range.HighlightColorIndex = replaceList(UBound(findList) + i + 1)
                count = count + 1
            Loop
        End With
        
        ' Store count and pattern in the summary string if count is greater than zero
        If count > 0 Then
            summary = summary & count & " items found for pattern: " & patternsToFind(i) & vbCrLf
        End If
    Next i
    
    ' Display the summary of items found (if any) at the end
    If Len(summary) > 0 Then
        MsgBox "Summary of items found:" & vbCrLf & summary, vbInformation, "Highlighting Summary"
    Else
        MsgBox "No items found.", vbInformation, "Highlighting Summary"
    End If
End Sub
Sub Nameanddatechecklist32()
    Dim findList() As Variant
    Dim replaceList() As Variant
    Dim i As Long
    Dim count As Long
    Dim summary As String
    
    ' Characters and phrases to find
    findList = Array("  ", "..", ". .", ". ,", ", ,", "?.", ");", "((", "))", "( ", " )", " ;", "et al ", "et al,", ", et al", "dsm", "ibid", "in press", "n.d.", "forthcoming", "under review", "blind", "this issue", "this volume", "personal communication", "Suppl")
    
    ' Patterns to find
    Dim patternsToFind() As Variant
    patternsToFind = Array("[0-9]{4}, ", "[A-Za-z] [0-9]{4}", "et al. [0-9]")
    
    ' Colors for highlighting
    ReDim replaceList(0 To UBound(findList) + UBound(patternsToFind) + 1) As Variant
    
    ' Set colors for character and phrase matches
    For i = LBound(findList) To UBound(findList)
        replaceList(i) = wdRed
    Next i
    
    ' Set colors for pattern matches
    For i = 0 To UBound(patternsToFind)
        replaceList(UBound(findList) + i + 1) = wdBlue ' Adjust the color index for patterns
    Next i
    
    ' Loop through the find list and highlight the matches for characters/phrases
    For i = LBound(findList) To UBound(findList)
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
        
        ' Store count and phrase in the summary string if count is greater than zero
        If count > 0 Then
            summary = summary & count & " items found for character or phrase: " & findList(i) & vbCrLf
        End If
    Next i
    
    ' Loop through the patterns and highlight the matches for patterns
    For i = LBound(patternsToFind) To UBound(patternsToFind)
        count = 0 ' Reset count for each pattern
        
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = patternsToFind(i)
            .MatchWildcards = True ' Use wildcards for pattern matching
            .Forward = True
            .Wrap = wdFindContinue
            
            ' Highlight the matches with the corresponding color and count the number of items found
            Do While .Execute
                Selection.range.HighlightColorIndex = replaceList(UBound(findList) + i + 1)
                count = count + 1
            Loop
        End With
        
        ' Store count and pattern in the summary string if count is greater than zero
        If count > 0 Then
            summary = summary & count & " items found for pattern: " & patternsToFind(i) & vbCrLf
        End If
    Next i
    
    ' Display the summary of items found (if any) at the end
    If Len(summary) > 0 Then
        MsgBox "Summary of items found:" & vbCrLf & summary, vbInformation, "Check_list Summary"
    Else
        MsgBox "No items found.", vbInformation, "Check_list Summary"
    End If
End Sub




Sub Nameanddatechecklist50()
    On Error Resume Next ' Enable error handling

    ' Constants for color indices
    Const RED_COLOR_INDEX As Long = 255
    ' ... other color constants ...

    ' Prompt for confirmation before making changes
    If MsgBox("Do you want to highlight items?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then
        Exit Sub ' User chose not to proceed
    End If

    ' Use document range instead of relying on Selection
    Dim doc As Document
    Set doc = ActiveDocument ' Assumes you are working with the active document

    Dim findList() As Variant
    Dim replaceList() As Variant
    Dim patternsToFind() As Variant
    ' ... your existing arrays ...
    findList = Array("  ", "..", ". .", ". ,", ", ,", "?.", ");", "((", "))", "( ", " )", " ;", "et al ", "et al,", ", et al", "dsm", "ibid", "in press", "n.d.", "forthcoming", "under review", "blind", "this issue", "this volume", "personal communication", "Suppl")
    ' Patterns to find
    patternsToFind = Array("[0-9]{4}, ", "[A-Za-z] [0-9]{4}", "et al. [0-9]")
    
    ' Colors for highlighting
    replaceList = Array(wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed, wdRed)
    ' Loop through the find list and highlight the matches
    Call HighlightItems(doc.range, findList, replaceList, RED_COLOR_INDEX)

    ' Loop through the patterns and highlight the matches
    Call HighlightPatterns(doc.range, patternsToFind, replaceList, RED_COLOR_INDEX)

    On Error GoTo 0 ' Reset error handling to default behavior
End Sub

Sub HighlightItems(rng As range, findList() As Variant, replaceList() As Variant, highlightColor As WdColorIndex)
    Dim i As Long
    Dim count As Long
    Dim summary As String
    
    For i = 0 To UBound(findList)
        count = 0
        
        With rng.Find
            ' ... your existing Find settings ...
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
                rng.HighlightColorIndex = highlightColor
                count = count + 1
            Loop
        End With
        
        ' Store count and phrase in the summary string if count is greater than zero
        If count > 0 Then
            summary = summary & count & " items found for character or phrase: " & findList(i) & vbCrLf
        End If
    Next i
    
    ' Display the summary of items found (if any) at the end
    If Len(summary) > 0 Then
        MsgBox "Summary of items found:" & vbCrLf & summary, vbInformation, "Highlighting Summary"
    Else
        MsgBox "No items found.", vbInformation, "Highlighting Summary"
    End If
End Sub

Sub HighlightPatterns(rng As range, patternsToFind() As Variant, replaceList() As Variant, highlightColor As WdColorIndex)
    Dim i As Long
    Dim count As Long
    Dim summary As String
    
    For i = LBound(patternsToFind) To UBound(patternsToFind)
        count = 0
        
        With rng.Find
            ' ... your existing Find settings ...
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = patternsToFind(i)
            .MatchWildcards = True ' Use wildcards for pattern matching
            .Forward = True
            .Wrap = wdFindContinue
            
            
            ' Highlight the matches with the corresponding color and count the number of items found
            Do While .Execute
                rng.HighlightColorIndex = highlightColor
                count = count + 1
            Loop
        End With
        
        ' Store count and pattern in the summary string if count is greater than zero
        If count > 0 Then
            summary = summary & count & " items found for pattern: " & patternsToFind(i) & vbCrLf
        End If
    Next i
    
    ' Display the summary of items found (if any) at the end
    If Len(summary) > 0 Then
        MsgBox "Summary of items found:" & vbCrLf & summary, vbInformation, "Highlighting Summary"
    Else
        MsgBox "No items found.", vbInformation, "Highlighting Summary"
    End If
End Sub

Sub ChatGPT()
 'Updateby Extendoffice
    Dim status_code As Integer
    Dim response As String
    OPENAI = "https://api.openai.com/v1/chat/completions"
    api_key = "sk-***************************** "
    If api_key = "" Then
        MsgBox "Please enter the API key."
        Exit Sub
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "Please select text."
        Exit Sub
    End If
    SendTxt = Replace(Replace(Replace(Replace(Selection.text, vbCrLf, ""), vbCr, ""), vbLf, ""), Chr(34), Chr(39))
    SendTxt = "{""model"": ""gpt-3.5-turbo"", ""messages"": [{""role"":""system"", ""content"":""You are a Word assistant""} ,{""role"":""user"", ""content"":""" & SendTxt & """}]}"
    Set Http = CreateObject("MSXML2.XMLHTTP")
    With Http
        .Open "POST", OPENAI, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & api_key
        .send SendTxt
      status_code = .Status
      response = .responseText
    End With
    If status_code = 200 Then
        Set regex = CreateObject("VBScript.RegExp")
        With regex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .pattern = """content"": ""(.*)"""
        End With
        Set matches = regex.Execute(response)
        If matches.count > 0 Then
            response = matches(0).SubMatches(0)
            response = Replace(Replace(response, "\n", vbCrLf), "\""", Chr(34))
            Selection.range.InsertAfter vbNewLine & response
        End If
    Else
        Debug.Print response
    End If
    Set Http = Nothing
End Sub

