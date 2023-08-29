Attribute VB_Name = "Lesson_Formatter"
' TODO: Format tables.

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
    
    ' Format Page Breaks:
    ReplaceDashesWithPageBreak
    
    ' Format Bullet Points
    Call ConvertToBulletPoints
    
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
            'Selection.Text = clean_matrix_block_equation(Selection.Text)
            Selection.Text = clean_matrix_inline_equation(Selection.Text)
            Selection.Text = clean_aligned_block_equation(Selection.Text)
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
            delete_empty_line
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
            Selection.Text = Left(Selection.Text, 1)
            Selection.MoveRight
            Selection.Style = ActiveDocument.Styles("Heading 1")
            Selection.MoveDown
            delete_empty_line
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
            Selection.Text = Left(Selection.Text, 1)
            Selection.MoveRight
            Selection.Style = ActiveDocument.Styles("Heading 2")
            Selection.MoveDown
            delete_empty_line
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
            Selection.Text = Left(Selection.Text, 1)
            Selection.MoveRight
            Selection.Style = ActiveDocument.Styles("Heading 3")
            Selection.MoveDown
            delete_empty_line
        End If
    Loop While Found
    
    ' Move the selection to the top of the page (Page 1)
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    
    ' subsubsubsections
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "^p##### "
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
            Selection.Text = Left(Selection.Text, 1)
            Selection.MoveRight
            Selection.Style = ActiveDocument.Styles("Heading 4")
            Selection.MoveDown
            delete_empty_line
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
    ' TODO - FIX ME
    Dim startTag As String
    Dim endTag As String
    Dim startIdx As Long
    Dim endIdx As Long
    Dim selectedText As String
    Dim cleanedText As String
    
    ' Define the start and end tags
    startTag = "\begin{bmatrix}" & vbCr
    endTag = "\end{bmatrix}" & vbCr
    
    ' Find the position of the start and end tags
    startIdx = InStr(inputText, startTag)
    endIdx = InStr(inputText, endTag)
    
    ' Check if both start and end tags are found
    If startIdx > 0 And endIdx > 0 Then
        ' Extract the text between start and end tags
        selectedText = Mid(inputText, startIdx, endIdx + Len(endTag) - startIdx)
        
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


Function clean_matrix_inline_equation(inputText As String) As String
    Dim cleanedText As String
    ' Replace "\begin{bmatrix}" with "\left(\begin{matrix}"
    cleanedText = Replace(inputText, "\begin{bmatrix}", "\left(\begin{matrix}")
    ' Replace "\end{bmatrix}" with "\end{matrix}\right)"
    clean_matrix_inline_equation = Replace(cleanedText, "\end{bmatrix}", "\end{matrix}\right)")
End Function


Function clean_aligned_block_equation(inputText As String) As String
    Dim startTag As String
    Dim endTag As String
    Dim startIdx As Long
    Dim endIdx As Long
    Dim selectedText As String
    Dim cleanedText As String
    
    ' Define the start and end tags
    startTag1 = "\begin{align*}" & vbCr
    endTag1 = "\end{align*}" & vbCr
    
    ' Find the position of the start and end tags
    startIdx1 = InStr(inputText, startTag1)
    endIdx1 = InStr(inputText, endTag1)
    
    ' Define the start and end tags
    startTag2 = "\begin{aligned}" & vbCr
    endTag2 = "\end{aligned}" & vbCr
    
    ' Find the position of the start and end tags
    startIdx2 = InStr(inputText, startTag2)
    endIdx2 = InStr(inputText, endTag2)
    
    If startIdx1 > 0 And endIdx1 > 0 Then
        startIdx = startIdx1
        endIdx = endIdx1
        startTag = startTag1
        endTag = endTag1
    Else
        startIdx = startIdx2
        endIdx = endIdx2
        startTag = startTag2
        endTag = endTag2
    End If
      
    ' Check if both start and end tags are found
    If startIdx > 0 And endIdx > 0 Then
        ' Extract the text between start and end tags
        selectedText = Mid(inputText, startIdx + Len(startTag), endIdx - (startIdx + Len(startTag)))
        
        ' Remove "\\\" from the end of each line
        selectedText = Replace(selectedText, "\\" & vbCr, vbCr)
        
        ' Combine cleaned text with original text
        cleanedText = Mid(inputText, 1, startIdx - 1) & selectedText & Mid(inputText, endIdx + Len(endTag))
    Else
        ' If start and/or end tags are not found, return the original input text
        cleanedText = inputText
    End If
    
    ' Return the cleaned text
    clean_aligned_block_equation = cleanedText
End Function


Sub ConvertToBulletPoints_OLD()
    Dim para As Paragraph
    Dim rng As Range
    
    ' Loop through each paragraph in the active document
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        
        ' Check if the paragraph starts with "- "
        If Left(rng.Text, 2) = "- " Then
            ' Convert the paragraph to a bullet point
            rng.ListFormat.ApplyBulletDefault
        End If
    Next para
End Sub


Sub ConvertToBulletPoints()
    'Note to call ReplaceDashesWithPageBreak() before this to prevent creation of random bullets
    Dim para As Paragraph
    Dim rng As Range
    Dim lineIndent As String
    Dim bulletIndent As Integer
    
    ' Loop through each paragraph in the active document
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        
        ' Get the text of the paragraph
        Dim paraText As String
        paraText = rng.Text
        
        ' Check if the paragraph starts with "-"
        If Left(paraText, 1) = "-" Then
            ' Get the indentation before the bullet
            Dim i As Integer
            For i = 1 To Len(paraText)
                If Mid(paraText, i, 1) = "-" Then
                    lineIndent = Left(paraText, i - 1)
                    Exit For
                End If
            Next i
            
            ' Determine the number of tabs or spaces in the indentation
            bulletIndent = 0
            For i = 1 To Len(lineIndent)
                If Mid(lineIndent, i, 1) = vbTab Then
                    bulletIndent = bulletIndent + 1
                ElseIf Mid(lineIndent, i, 1) = " " Then
                    bulletIndent = bulletIndent + 1
                    If i < Len(lineIndent) And Mid(lineIndent, i + 1, 1) = " " Then
                        ' Treat two spaces as one level of indentation
                        bulletIndent = bulletIndent + 1
                        i = i + 1
                    End If
                Else
                    Exit For
                End If
            Next i
            
            ' Convert the paragraph to a bullet point
            rng.ListFormat.ApplyBulletDefault
            ' Indent the bullet point based on indentation level
            rng.ParagraphFormat.LeftIndent = bulletIndent * CentimetersToPoints(0.5) ' Adjust the multiplier as needed
        End If
    Next para
End Sub


Sub ReplaceDashesWithPageBreak()
    Dim para As Paragraph
    Dim rng As Range
    
    ' Loop through each paragraph in the active document
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        
        ' Check if the paragraph starts with "---"
        If Left(rng.Text, 3) = "---" Then
            ' Replace the line with a page break
            rng.Text = ""
            rng.InsertBreak Type:=wdPageBreak
        End If
    Next para
End Sub

Private Sub delete_empty_line()
    ' Expand the Selection to the entire line
    Selection.Expand Unit:=wdLine
    
    ' Get the text of the selected line
    Dim selectedLine As String
    selectedLine = Selection.Text
    
    ' Trim the line text to remove leading and trailing spaces
    Dim lineText As String
    lineText = Trim(selectedLine)
    
    ' Check if the selected line is empty
    If selectedLine = vbCr Then
        Selection.Delete
    Else
        Selection.MoveLeft
    End If
End Sub
