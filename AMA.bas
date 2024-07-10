Attribute VB_Name = "AMA"
Sub TitleCaseAMAMicroUS102()
    Dim lcList As String
    Dim wrd As range
    Dim p As paragraph
    
    ' List of lowercase words, separated by spaces
    lcList = " is a and had the that an and as but for if nor or so yet a an the as at by for in of off on per to up via vs are"
    
    For Each p In ActiveDocument.Paragraphs
        Select Case p.style
            Case "H1", "H2", "H3"
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
