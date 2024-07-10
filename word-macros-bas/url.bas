Attribute VB_Name = "url"
Function CheckURL(strURL As String) As Boolean
  Dim objDemand As Object
  Dim varResult As Variant
 
  On Error GoTo ErrorHandler
  Set objDemand = CreateObject("WinHttp.WinHttpRequest.5.1")
 
  With objDemand
    .Open "GET", strURL, False
    .send
    varResult = .StatusText
  End With
 
  Set objDemand = Nothing
 
  If varResult = "OK" Then
    CheckURL = True
  Else
    CheckURL = False
  End If
 
ErrorHandler:
End Function

Sub ReturnURLCheck()
  Dim objLink As hyperlink
  Dim strLinkText As String
  Dim strLinkAddress As String
  Dim strResult As String
  Dim nInvalidLink As Integer, nTotalLinks As Integer
  Dim objDoc As Document
 
  Application.ScreenUpdating = False
 
  Set objDoc = ActiveDocument
  nTotalLinks = objDoc.Hyperlinks.count
  nInvalidLink = 0
 
  With objDoc
    For Each objLink In .Hyperlinks
      strLinkText = objLink.range.text
      strLinkAddress = objLink.address
 
      If Not CheckURL(strLinkAddress) Then
        nInvalidLink = nInvalidLink + 1
        strResult = frmCheckURLs.txtShowResult.text
        frmCheckURLs.txtShowResult.text = strResult & nInvalidLink & ". Invalid Link Information:" & vbNewLine & _
                                          "Displayed Text: " & strLinkText & vbNewLine & _
                                           "Address: " & strLinkAddress & vbNewLine & vbNewLine
      End If
    Next objLink
 
    frmCheckURLs.txtTotalLinks.text = nTotalLinks
    frmCheckURLs.txtNumberOfInvalidLinks.text = nInvalidLink
    frmCheckURLs.Show Modal
 
  End With
  Application.ScreenUpdating = True
End Sub

Sub HighlightInvalidLinks()
  Dim objLink As hyperlink
  Dim strLinkAddress As String
  Dim strResult As String
  Dim objDoc As Document
 
  Set objDoc = ActiveDocument
 
  With objDoc
    For Each objLink In .Hyperlinks
      strLinkAddress = objLink.address
 
      If Not CheckURL(strLinkAddress) Then
        objLink.range.HighlightColorIndex = wdYellow
      End If
    Next objLink
  End With
End Sub

