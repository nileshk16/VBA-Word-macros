Attribute VB_Name = "TA"
Sub TA_PE()


    ''General
    Call ReplaceQuotes
        
    ''Frontmatter
    
    Call TY
    
    Call MoveTY
    
    Call LRHX
    
    Call LRH
    
    Call etal
    
    Call ATAftercolon1
    
    'Call ATAftercolonLowerCaser
    
    Call FindSuperscriptInAUStyle
    
    Call commaand
    
    Call AFdot
    
    Call AUdot
    
    Call Keywordschangestyle
    
    Call AfterColonABKW
    
    Call SelectKeywords
    
    Call SortWordsAlphabetically
    
    ''Headlevels
    
    Call TAH3dot
    
    ''Tabels and Figures
    
    Call TSinsert
    
    ''Backmatter
    
    Call TACompettingintrestes
    
    Call TAcontributions
    

    ''References
    Call REFAftercolon
    
    Call pwithoutspace

End Sub
Sub TAshaderemover()
Dim oStory As range
    For Each oStory In ActiveDocument.StoryRanges
        oStory.Font.Shading.BackgroundPatternColor = wdColorAutomatic
        If oStory.StoryType <> wdMainTextStory Then
            While Not (oStory.NextStoryRange Is Nothing)
                Set oStory = oStory.NextStoryRange
                oStory.Font.Shading.BackgroundPatternColor = wdColorAutomatic
            Wend
        End If
    Next oStory
lbl_Exit:
    Set oStory = Nothing
    Exit Sub
End Sub
Sub TAfontcolor()
With Selection.Find
.ClearFormatting
Selection.Find.Font.Color = wdColorPink
        .text = ""
        .Replacement.ClearFormatting
        .Replacement.Font.Color = wdColorBlack 'I added this line
        .Execute Replace:=wdReplaceAll, Forward:=True, _
         Wrap:=wdFindContinue
    End With
End Sub
Sub TACompettingintrestes()
'
' Macro7 Macro
With Selection.Find
 .ClearFormatting
 .text = "The authors declare that there is no conflict of interest."
 .Replacement.ClearFormatting
 .Replacement.text = "The authors declared no potential conflicts of interest with respect to the research, authorship, and/or publication of this article."
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub
Sub TAcontributions()
'
' re Macro
'
  With ActiveDocument.Content.Find
 .ClearFormatting
 .text = "Author contribution(s)"
 .Replacement.ClearFormatting
 .Replacement.text = "Author contributions"
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub
Sub TAH3dot()
With ActiveDocument.Content.Find
 .ClearFormatting
 .text = "."
 .style = "H3"
 .Replacement.ClearFormatting
 .Replacement.text = ""
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub

Sub TAclean()
    
    Call TAfontcolor
    
    Call TurnOffTrackChanges
    
    Call AcceptAllChanges
    
    
    
    Call RemoveHyperlinks
    
    Call TAshaderemover
    
    
End Sub

