Attribute VB_Name = "shortcut"
Sub highlightremover() '
' highlightremover Macro'
    Options.DefaultHighlightColorIndex = wdNoHighlight
    Selection.range.HighlightColorIndex = wdNoHighlight
End Sub
Sub highlighter()
' highlighter Macro'
    Options.DefaultHighlightColorIndex = wdYellow
    Selection.range.HighlightColorIndex = wdYellow
End Sub
Sub OpenBrowser(strAddress As String)
    Dim chromePath As String
    
    ' Replace with the actual path to the Chrome executable
    chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    
    ' Open the URL in a new Chrome process
    Shell """" & chromePath & """ """ & strAddress & """", vbNormalFocus
End Sub
Sub SearchOnGoogle()
    Dim strText As String
        If Selection.Type <> wdSelectionIP Then
            strText = Selection.text
            strText = Trim(strText)
            selectedStyle = Selection.style
            Selection.Copy
        Else
            MsgBox ("Please select text first!")
            Exit Sub
        End If
        ' Search selected text on Google
        'OpenBrowser "https://www.google.com/search?num=20&hl=en&q=" & strText
        ' Check if the selected text is a URL starting with "https://"
    If LCase(Left(strText, 8)) = "https://" Or LCase(Left(strText, 7)) = "http://" Or LCase(Left(strText, 4)) = "www." Then
        OpenBrowser strText
    ElseIf selectedStyle = "stl" Then
        OpenBrowser "https://www.google.com/search?num=20&hl=en&q=" & strText & " abbreviation"
    Else
        OpenBrowser "https://www.google.com/search?num=20&hl=en&q=" & strText
    End If
End Sub
Sub SearchOnGoogleAbb()
    Dim strText As String
        If Selection.Type <> wdSelectionIP Then
            strText = Selection.text
            strText = Trim(strText)
            Selection.Copy
        Else
            MsgBox ("Please select text first!")
            Exit Sub
        End If
        
        ' Search selected text on Google
        OpenBrowser "https://www.google.com/search?num=20&hl=en&q=" & strText & " abbreviation"
End Sub
Sub SearchOnGoogle2()
    Dim strText As String
        If Selection.Type <> wdSelectionIP Then
            strText = Selection.text
            strText = Trim(strText)
            Selection.Copy
        Else
            MsgBox ("Please select text first!")
            Exit Sub
        End If
        
        ' Search selected text on Google
        OpenBrowser strText
End Sub
Sub TurnOffTrackChanges()
    ' Turn off Track Changes
    ActiveDocument.TrackRevisions = False
    ActiveDocument.ShowRevisions = False
End Sub
Sub AcceptAllChanges()
    ' Accept all changes in the document
    ActiveDocument.AcceptAllRevisions
End Sub
Sub RemoveHyperlinks()
    Dim hyperlink As hyperlink
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Repeat the loop until there are no hyperlinks left
    Do While doc.Hyperlinks.count > 0
        For Each hyperlink In doc.Hyperlinks
            hyperlink.Delete
        Next hyperlink
    Loop
    MsgBox "Hyperlinks Removed"
End Sub
Sub InseartTabinTable()
    Dim tbl As TABLE
    Dim cell As cell
    Dim selectedRange As range

    ' Check if a table is selected
    If Selection.Information(wdWithInTable) Then
        ' Set the selectedRange to the currently selected cells
        Set selectedRange = Selection.range
        
        ' Loop through each cell in the selectedRange
        For Each cell In selectedRange.Cells
            ' Check if the current cell is in the first column
            If cell.ColumnIndex = 1 Then
                cell.range.InsertBefore vbTab
            End If
        Next cell
    Else
        MsgBox "Please select a table before running this macro."
    End If
End Sub
Sub TurnOnTrackChanges()
    ' Turn on track changes
    ActiveDocument.TrackRevisions = True
    
     ActiveDocument.TrackMoves = False
    
    ' Turn off track moves
    'ActiveDocument.ShowRevisions = wdRevisionsNone
End Sub

