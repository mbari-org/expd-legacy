<%      
   Function filteredText(strDirty)
   'Purpose:  Filters the input string to remove offensive characters
   'Inputs:   The ascii string to be filtered
   'Outputs:  Returns the sanitized string modified as follows:
   '          1) non-printing control characters are converted to space
   '          2) whitespace (tabs, spaces etc) are converted to space and
   '             collapsed to a single instance if needed
   '          3) backslash "\" converted to forwardslash "/"
   '          4) double " and single ' quotes to backtic 1 and collapsed
   '             to a single instance as needed.
   'Note1:    Requires VBScript 5.5 to be loaded on the webserver.
   'Note2:    The effect is cummulative - ie if a control char is next
   '          to a space or tab, the control character is converted to
   '          a space by the first step, then collapsed to a single space
   '          by the second.

   Dim rgx1
   Dim rgx2
   Dim rgx3
   Dim rgx4
   
   Set rgx1 = New RegExp
   Set rgx2 = New RegExp
   Set rgx3 = New RegExp
   Set rgx4 = New RegExp
   
   Dim i
   Dim strTmp, strClean
   
   ' pre-processing trim
   strClean = Trim(strDirty)
   
   'set up for regular expression objects - one per test
   'control characters 1-31
   strTmp = "["
   For i = 1 To 31
      If Chr(i) = vbCr Then
         strTmp = strTmp & "\r"
      ElseIf Chr(i) = vbLf Then
         strTmp = strTmp & "\r"
      Else
         strTmp = strTmp & Chr(i)
      End If
   Next
   strTmp = strTmp & "]+"
   
   With rgx1
         .Pattern = strTmp 'control characters
         .IgnoreCase = True
         .Global = True
   End With
   With rgx2
         .Pattern = "[ \t]+" 'whitespace with collapse
         .IgnoreCase = True
         .Global = True
   End With   
   With rgx3
         .Pattern = "[\\]" 'backslash
         .IgnoreCase = True
         .Global = True
   End With   
   With rgx4
         .Pattern = "[" & Chr(34) & Chr(39) & "]+" 'single and double quotes
         .IgnoreCase = True
         .Global = True
   End With   
   
   'perform each test - the output of one test is the input
   'to the other so the effect is cumulative.
    
   ' control characters to spaces
      
   strClean = rgx1.Replace(strClean, " ")
    
   'multiple whitespace to single space

   strClean = rgx2.Replace(strClean," ")

   'backslash \ to forwardslash /

   strClean = rgx3.Replace(strClean,"/")   

   'single or double quotes to backtic

   strClean = rgx4.Replace(strClean,"`")

   ' cleanup
   Set rgx1 = nothing
   Set rgx2 = nothing
   Set rgx3 = nothing
   Set rgx4 = nothing

   ' post processing trim
   strClean = Trim(strClean)

   filteredText = strClean
   
   End Function 'filteredText

   '------------------------------------------------------------------

   Function filteredHeader(strDirty)
   'Purpose:  Filters the input string to remove offensive characters and spaces.
   '          Used to clean up the column headers in case the user, or a program
   '          that the user uses to modify the waypoints adds extraneous spaces
   '          or characters.
   'Inputs:   The ascii string to be filtered
   'Outputs:  Returns the sanitized string modified as follows:
   '          1) non-printing control characters are converted to nothing (empty string)
   '          2) whitespace (tabs, spaces etc) are converted to nothing
   '          3) slashes "/" "\" and dashes "-" are converted to nothing
   '          4) double " and single ' quotes and backticks are converted to nothing
   'Note1:    Requires VBScript 5.5 to be loaded on the webserver.

   Dim rgx1
   Dim rgx2
   Dim rgx3
   Dim rgx4
   
   Set rgx1 = New RegExp
   Set rgx2 = New RegExp
   Set rgx3 = New RegExp
   Set rgx4 = New RegExp
   
   Dim i
   Dim strTmp, strClean
   
   ' pre-processing trim
   strClean = Trim(strDirty)
   
   'set up for regular expression objects - one per test
   'control characters 1-31
   strTmp = "["
   For i = 1 To 31
      If Chr(i) = vbCr Then
         strTmp = strTmp & "\r"
      ElseIf Chr(i) = vbLf Then
         strTmp = strTmp & "\r"
      Else
         strTmp = strTmp & Chr(i)
      End If
   Next
   strTmp = strTmp & "]+"
   
   With rgx1
         .Pattern = strTmp 'control characters
         .IgnoreCase = True
         .Global = True
   End With
   With rgx2
         .Pattern = "[ \t]+" 'whitespace with collapse
         .IgnoreCase = True
         .Global = True
   End With   
   With rgx3
         .Pattern = "[\\/-]" 'slashes and dashes
         .IgnoreCase = True
         .Global = True
   End With   
   With rgx4
         .Pattern = "[" & Chr(34) & Chr(39) & "`" & "]+" 'single & double quotes and backticks
         .IgnoreCase = True
         .Global = True
   End With   
   
   'perform each test - the output of one test is the input
   'to the other so the effect is cumulative.
    
   ' control characters to nothing
      
   strClean = rgx1.Replace(strClean, "")

   'slashes and dashes to nothing

   strClean = rgx3.Replace(strClean, "")

   'single or double quotes and backticks to nothing

   strClean = rgx4.Replace(strClean, "")

   'multiple whitespace to nothing

   strClean = rgx2.Replace(strClean, "")

   ' cleanup
   Set rgx1 = nothing
   Set rgx2 = nothing
   Set rgx3 = nothing
   Set rgx4 = nothing

   ' post processing trim
   strClean = Trim(strClean)

   filteredHeader = strClean
   
   End Function 'filteredHeader

   '------------------------------------------------------------------

   Function reformatNumber(strInput, intScale)
   'Purpose:  Format numbers for readability
   'Inputs:   The ascii string to be filtered
   '          The scale (places to the right of the decimal point)
   'Outputs:  Returns the reformated number, or the untouched string (if
   '          the input was non-numeric)

   Dim datum 'the working value

   Select Case VarType(strInput)
      Case 2, 3 'integer / long
         If (strInput<>0) Then
            datum = FormatNumber(strInput, 0, -1, 0, 0)
         End If			
      Case 4, 5 'single / double
         datum = FormatNumber(strInput, intScale, -1, 0, 0)
      Case 8 'strings that are numeric
         If IsNumeric(strInput) Then
            datum = FormatNumber(strInput, intScale, -1, 0, 0)
         Else
            datum = strInput
         End If
      Case 14 'decimal
         datum = FormatNumber(strInput, intScale, -1, 0, 0)
      Case Else
         datum = strInput
      End Select

   reformatNumber = datum

   End Function 'reformatNumber
%>