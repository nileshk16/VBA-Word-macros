Attribute VB_Name = "Samples"
Sub GetFileName()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Display the file name in a message box
    MsgBox "File Name: " & doc.Name, vbInformation, "Current File Name"
    
    ' Clean up the object
    Set doc = Nothing
End Sub
Sub ExecuteMacroBasedOnFileName()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Get the file name
    Dim FileName As String
    FileName = doc.Name
    
    ' Check if the file name starts with "TA"
    If Left(FileName, 2) = "TA" Then
        ' Execute Macro1
        Call TA_PE
    End If
    
    ' Clean up the object
    Set doc = Nothing
End Sub
Sub RunMacroBasedOnInput()
    Dim userInput As String
    userInput = InputBox("Enter 1 to run Macro1 or 2 to run Macro2:", "Macro Selection")
    
    Select Case userInput
        Case "1"
            Call Macro11
        Case "2"
            Call Macro22
        Case Else
            MsgBox "Invalid input. Please enter 1 or 2.", vbExclamation
    End Select
End Sub
Sub Macro11()
    ' Place the code for Macro1 here
    MsgBox "Running Macro1"
End Sub
Sub Macro22()
    ' Place the code for Macro2 here
    MsgBox "Running Macro2"
End Sub
Sub InsertCurlyQuotes()
    Dim leftQuote As String
    Dim rightQuote As String
    
    ' Unicode values for left and right curly quotes
    leftQuote = ChrW(8220)
    rightQuote = ChrW(8221)
    
    ' Insert left and right curly quotes
    Selection.TypeText text:=leftQuote & "Quoted text" & rightQuote
End Sub
