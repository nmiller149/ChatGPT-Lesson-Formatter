Attribute VB_Name = "Lesson_Formatter"
' TODO: Format Markdown Sections and bullet points. Extra credit for formatting tables.

Sub Format_Lesson()
Attribute Format_Lesson.VB_Description = "Formats ChatGPT lessons copied and pasted into MS Word"
Attribute Format_Lesson.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Format_Lesson"
'
' Format_Lesson Macro
' Formats ChatGPT lessons copied and pasted into MS Word
'
' Note that the document should only contain either LaTex or Unicode equations and MS Word must have the input type set accordingly.
'
    'Format Whole Page to normal style
    Selection.WholeStory
    'Selection.Style = ActiveDocument.Styles("Normal")
    Selection.HomeKey Unit:=wdLine
    
    
    'Format Lesson Title
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "Lesson:"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Found = Selection.Find.Execute
    If Found Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.Style = ActiveDocument.Styles("Title")
        Selection.MoveDown
    End If
    
    'Format Each Section
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "Section *:"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Do
        Found = Selection.Find.Execute
        If Found Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.Style = ActiveDocument.Styles("Heading 1")
            Selection.MoveDown
        End If
    Loop While Found
    
    ' Move the selection to the top of the page (Page 1)
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    
    'Format Each Subsection (GPT3)
    Call format_gpt3_subsections
    
    
    'Format Each Subsection (GPT4/Markdown)
    'Call format_gpt4_subsections
    Call format_markdown
    
    
    ' Format All Equations
    Find_Convert_Block_Eq
    Find_Convert_In_Line
    Find_Convert_UnicodeMath
    
    MsgBox "Formatting Completed!"
End Sub


Sub Find_Convert_UnicodeMath()
'
' Find_Convert_UnicodeMath Macro
' Finds and builds all UnicodeMath equations to MS Word formatted equations
'
    ' Move the selection to the top of the page (Page 1)
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "```*```"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Do
        Found = Selection.Find.Execute
        If Found Then
            Selection.Text = Mid(Selection.Text, 4, Len(Selection.Text) - 6)
            Selection.OMaths.Add Range:=Selection.Range
            Selection.OMaths.BuildUp
            Selection.MoveRight
            Selection.MoveRight
        End If
    Loop While Found
End Sub


Sub Find_Convert_In_Line()
'
' Find_Convert_In_Line Macro
' Finds and converts all In-Line LaTeX equations to MS Word formatted equations
'
    ' Move the selection to the top of the page (Page 1)
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "\\\(*\\\)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Do
        Found = Selection.Find.Execute
        If Found Then
            Selection.Text = Mid(Selection.Text, 3, Len(Selection.Text) - 4)
            Selection.OMaths.Add Range:=Selection.Range
            Selection.OMaths.BuildUp
            Selection.MoveRight
            Selection.MoveRight
        End If
    Loop While Found
End Sub


Sub Find_Convert_Block_Eq()
'
' Find_Convert_Block_Eq Macro
' Finds and converts all Block LaTeX equations to MS Word formatted equations
'
    ' Move the selection to the top of the page (Page 1)
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "\\\[*\\\]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Do
        Found = Selection.Find.Execute
        If Found Then
            Selection.Text = clean_matrix_block_equation(Selection.Text)
            Selection.Text = Mid(Selection.Text, 4, Len(Selection.Text) - 5)
            Selection.OMaths.Add Range:=Selection.Range
            Selection.OMaths.BuildUp
            Selection.MoveRight
            Selection.MoveRight
        End If
    Loop While Found
End Sub

Private Sub format_gpt3_subsections()
'
' format_subsections private subroutine
'
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "Subsection *.*:"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Do
        Found = Selection.Find.Execute
        If Found Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.Style = ActiveDocument.Styles("Heading 2")
            Selection.MoveDown
            ' ERASE EMPTY LINE AFTER SUBSECTION
            ' Expand the Selection to the entire line
            Selection.Expand Unit:=wdLine
            ' Get the text of the selected line
            selectedLine = Selection.Text
            ' Check if the selected line is empty
                    ' Initialize the flag to track if non-ASCII characters are found
            hasNonASCII = False
            ' Loop through each character in the line
            For i = 1 To Len(selectedLine)
                ' Get the ASCII code of the character
                charCode = Asc(Mid(selectedLine, i, 1))
                
                ' Check if the character is outside the range of printable ASCII characters
                If charCode < 32 Or charCode > 126 Then
                    hasNonASCII = True
                    Exit For
                End If
            Next i
            ' If non-ASCII characters are found, delete the entire line
            If hasNonASCII Then
                Selection.Delete
            End If
            Selection.MoveDown
        End If
    Loop While Found

End Sub


Private Sub format_gpt4_subsections()
'
' format_subsections private subroutine
'
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "###"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Do
        Found = Selection.Find.Execute
        If Found Then
            ' ERASE Markdown Subsection Delimitters
            Selection.Text = Right(Selection.Text, Len(Selection.Text) - 3)
            Selection.Style = ActiveDocument.Styles("Heading 2")
            Selection.MoveDown
            ' ERASE EMPTY LINE AFTER SUBSECTION
            ' Expand the Selection to the entire line
            Selection.Expand Unit:=wdLine
            ' Get the text of the selected line
            selectedLine = Selection.Text
            ' Check if the selected line is empty
                    ' Initialize the flag to track if non-ASCII characters are found
            hasNonASCII = False
            ' Loop through each character in the line
            For i = 1 To Len(selectedLine)
                ' Get the ASCII code of the character
                charCode = Asc(Mid(selectedLine, i, 1))
                
                ' Check if the character is outside the range of printable ASCII characters
                If charCode < 32 Or charCode > 126 Then
                    hasNonASCII = True
                    Exit For
                End If
            Next i
            ' If non-ASCII characters are found, delete the entire line
            If hasNonASCII Then
                Selection.Delete
            End If
            Selection.MoveDown
        End If
    Loop While Found

End Sub


Private Sub format_markdown()
'
' format_markdown subroutine
' formats titles, sections, subsetctions of markdown language to Word format
'
'
    ' Move the selection to the top of the page (Page 1)
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    
    ' Titles
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "^p# "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Do
        Found = Selection.Find.Execute
        If Found Then
            ' ERASE Markdown Delimitters
            Selection.Text = ""
            Selection.Style = ActiveDocument.Styles("Title")
            Selection.MoveDown
        End If
    Loop While Found
    
    ' Move the selection to the top of the page (Page 1)
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    
    ' Sections
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "^p## "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Do
        Found = Selection.Find.Execute
        If Found Then
            ' ERASE Markdown Delimitters
            Selection.Text = ""
            Selection.Style = ActiveDocument.Styles("Heading 1")
            Selection.MoveDown
        End If
    Loop While Found
    
    ' Move the selection to the top of the page (Page 1)
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    
    ' subsections
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "^p### "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Do
        Found = Selection.Find.Execute
        If Found Then
            ' ERASE Markdown Subsection Delimitters
            Selection.Text = ""
            Selection.Style = ActiveDocument.Styles("Heading 2")
            Selection.MoveDown
            ' ERASE EMPTY LINE AFTER SUBSECTION
            ' Expand the Selection to the entire line
            Selection.Expand Unit:=wdLine
            ' Get the text of the selected line
            selectedLine = Selection.Text
            ' Check if the selected line is empty
                    ' Initialize the flag to track if non-ASCII characters are found
            hasNonASCII = False
            ' Loop through each character in the line
            For i = 1 To Len(selectedLine)
                ' Get the ASCII code of the character
                charCode = Asc(Mid(selectedLine, i, 1))
                
                ' Check if the character is outside the range of printable ASCII characters
                If charCode < 32 Or charCode > 126 Then
                    hasNonASCII = True
                    Exit For
                End If
            Next i
            ' If non-ASCII characters are found, delete the entire line
            If hasNonASCII Then
                Selection.Delete
            End If
            Selection.MoveDown
        End If
    Loop While Found
    
    ' Move the selection to the top of the page (Page 1)
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    
    ' subsubsections
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "^p#### "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Do
        Found = Selection.Find.Execute
        If Found Then
            ' ERASE Markdown Subsection Delimitters
            Selection.Text = ""
            Selection.Style = ActiveDocument.Styles("Heading 3")
            Selection.MoveDown
            ' ERASE EMPTY LINE AFTER SUBSECTION
            ' Expand the Selection to the entire line
            Selection.Expand Unit:=wdLine
            ' Get the text of the selected line
            selectedLine = Selection.Text
            ' Check if the selected line is empty
                    ' Initialize the flag to track if non-ASCII characters are found
            hasNonASCII = False
            ' Loop through each character in the line
            For i = 1 To Len(selectedLine)
                ' Get the ASCII code of the character
                charCode = Asc(Mid(selectedLine, i, 1))
                
                ' Check if the character is outside the range of printable ASCII characters
                If charCode < 32 Or charCode > 126 Then
                    hasNonASCII = True
                    Exit For
                End If
            Next i
            ' If non-ASCII characters are found, delete the entire line
            If hasNonASCII Then
                Selection.Delete
            End If
            Selection.MoveDown
        End If
    Loop While Found
    
    ' Move the selection to the top of the page (Page 1)
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    
    ' bold font
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "[*][*]*[*][*]"
        .Replacement.Text = ""
        .Forward = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchWildcards = True
    End With
    Do
        Found = Selection.Find.Execute
        If Found Then
            ' ERASE Markdown Delimitters
            Selection.Text = Mid(Selection.Text, 3, Len(Selection.Text) - 4)
            Selection.Font.Bold = True
            Selection.MoveRight
        End If
    Loop While Found
End Sub


Function clean_matrix_block_equation(inputText As String) As String
    Dim startTag As String
    Dim endTag As String
    Dim startIdx As Long
    Dim endIdx As Long
    Dim selectedText As String
    Dim cleanedText As String
    
    ' Define the start and end tags
    startTag = "\begin{align*}" & vbCr
    endTag = "\end{align*}" & vbCr
    
    ' Find the position of the start and end tags
    startIdx = InStr(inputText, startTag)
    endIdx = InStr(inputText, endTag)
    
    ' Check if both start and end tags are found
    If startIdx > 0 And endIdx > 0 Then
        ' Extract the text between start and end tags
        selectedText = Mid(inputText, startIdx + Len(startTag), endIdx - (startIdx + Len(startTag)))
        
        ' Remove "\\\" from the end of each line
        selectedText = Replace(selectedText, "\\" & vbCr, vbCr)
        
        ' Replace "\begin{bmatrix}" with "\left(\begin{matrix}"
        selectedText = Replace(selectedText, "\begin{bmatrix}", "\left(\begin{matrix}")
        
        ' Replace "\end{bmatrix}" with "\end{matrix}\right)"
        selectedText = Replace(selectedText, "\end{bmatrix}", "\end{matrix}\right)")
        
        ' Combine cleaned text with original text
        cleanedText = Mid(inputText, 1, startIdx - 1) & selectedText & Mid(inputText, endIdx + Len(endTag))
    Else
        ' If start and/or end tags are not found, return the original input text
        cleanedText = inputText
    End If
    
    ' Return the cleaned text
    clean_matrix_block_equation = cleanedText
End Function
