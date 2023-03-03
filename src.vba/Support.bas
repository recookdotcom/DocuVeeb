Option Explicit



''
' Parses the declaration section for any group comments attached to the module.
'
' In a VBA module, the declaration section includes all lines prior to the first function definition.
' The section ends with the last variable or constant declaration.
' <H4>Group Description Layout</H4>
' <P>A group description follows a similar structure as for a function description, with a brief description
' followed by a blank comment line and an optional full description.
' The description must begin with a line containing only a pair of single quotes, as with a function description.
' However, the description must also end with another line containing only a pair of single quotes.
' The brief description must include the @groupnote tag, as in the example.
' If the group comment block is not formatted properly, the group notes are not captured.
' <P>Group descriptions are currently implemented only for classes.
' <P>To be interpreted, group descriptions must be in the "Declarations" section of a module.
' If the module contains no statements that would imply a Declarations section, the section
' can be forced by the use of a nonsense declaration, such as:<br><br>
'
' <code>&nbsp;&nbsp;'' <br>
' &nbsp;&nbsp;' @groupnote This class handles the grouping of declarations within a project. <br>
' &nbsp;&nbsp;' <br>
' &nbsp;&nbsp;' This class blah blah blah... <br>
' &nbsp;&nbsp;''<br>
' &nbsp;&nbsp;Private DeclarationEnd = 0  ' Dummy statement to end the declaration section
' </code><br><br>
' A group description for a class must appear in the class module.
' @param An array containing the entire declaration section.
' @return No return value.
' @author Dr. Richard Cook
' @date 11/24/22
'Public Function DescriptionParse(ByRef arrDeclarationSection() As String)
Public Function DescriptionParse(ByRef arrDeclarationSection As Variant)
    Dim nLineCur As Integer, szTag As String, nStart As Integer, nStop As Integer ', bNewTag As Boolean
    Dim szLineCur As String, nPosition As Integer, nParamCur As Integer, szDescription As String
    Dim bTagSection As Boolean, nItem As Integer
    
    nStart = LBound(arrDeclarationSection)
    nStop = UBound(arrDeclarationSection)
    szDescription = ""
    szTag = ""

    ' Loop through array of input comment lines
    For nLineCur = nStart To nStop
        ' Set the tags for the implicit tag lines
        ' The input lines following these implicit tags will be handled according to these tags
        If arrDeclarationSection(nLineCur) = "" Then
            ' Do nothing with blank lines
        ElseIf Left(arrDeclarationSection(nLineCur), 6) = "Option" Then
'            ' Add to list of module options #### Do list as array? ####
'            If InStr(arrDeclarationSection(nLineCur), "Option Private Module") = 1 Then
'                bOptionPrivate = True
'            ElseIf InStr(arrDeclarationSection(nLineCur), "Option Base") = 1 Then
'                nOptionBase = Val(Right(arrDeclarationSection(nLineCur), 2))
'            ElseIf InStr(arrDeclarationSection(nLineCur), "Option Compare") = 1 Then
'                szOptionCompare = Mid(arrDeclarationSection(nLineCur), Len("Option Compare") + 2)
'            End If
        ElseIf arrDeclarationSection(nLineCur) = "''" Then
            bTagSection = False ' End of group note section
        Else
            ' Remove the leading comment mark, following space
            szLineCur = Trim(Mid(arrDeclarationSection(nLineCur), 2))
            
            ' Look for the groupnote tag
            ' Tag remains in effect until next tag encountered
            If Left(szLineCur, 10) = "@groupnote" Then
                ' We've now entered the tagged group note section
                bTagSection = True
                ' A space delimits the @tag from its contents
                nPosition = InStr(szLineCur, " ")
                If nPosition > 0 Then
                    'szTag = Left(szLineCur, nPosition - 1)
                    szLineCur = Mid(szLineCur, nPosition + 1)
                Else
                    'szTag = szLineCur
                    szLineCur = ""
                End If
            ElseIf Left(szLineCur, 6) = "@group" Then
                szTag = Mid(szLineCur, 8)
                szLineCur = ""
            ElseIf Left(arrDeclarationSection(nLineCur), 3) = "' @" Then
                szLineCur = ""
            End If  ' @
            szLineCur = Trim(szLineCur)
        End If  ' nLineCur contents
        
        If bTagSection = True Then
            szDescription = szDescription & " " & szLineCur
        End If
    Next nLineCur

    If szDescription <> "" Then
        If szTag <> "" Then
            szDescription = szTag & "|" & szDescription
        End If
    End If
    DescriptionParse = szDescription

End Function  ' DescriptionParse()



''
' Sorts a single dimension array using a standard bubble sort.
'
' Sorts the array in ascending order.
' <p>Bubble sort is slow with large arrays, but the coding is clear and easy to maintain.
' <p>The function Array2DBubbleSort() sorts a 2-D array.
' @param The array to be sorted.
' @param <i>This parameter is ignored</i>
' @return Always returns True.
' @group Arrays
' @see Array2DBubbleSort
' @author Dr. Richard Cook
Public Function ArrayBubbleSort(arrToSort() As Variant, Optional nSortOrder As Integer = 0)

    Dim i As Integer, j As Integer, vntTemp As Variant, nLast As Long
    Dim inc As Integer, rc As Integer, nFirst As Long, vntOutput As Variant

    nFirst = LBound(arrToSort)
    nLast = UBound(arrToSort) ' - nFirst + 1

    For i = nFirst To nLast - 1
        For j = i + 1 To nLast
            If arrToSort(i) > arrToSort(j) Then
                vntTemp = arrToSort(i)
                arrToSort(i) = arrToSort(j)
                arrToSort(j) = vntTemp
            End If
        Next
    Next
    
    ArrayBubbleSort = True

End Function  ' ArrayBubbleSort()


''
' Sorts a 2-dimension array according to the 1st dimension.
'
' The function assumes the input is a rectangular array.
' (The 2nd dimension may contain any number of elements, but must be of fixed size.)
' Bubble sort is slow with large arrays, but with 2 dimensions the coding is clear and easy to maintain.
' @param The array to be sorted.
' @param <i>This parameter is ignored</i>
' @return Always returns True.
' @group Arrays
' @see ArrayBubbleSort
Public Function Array2DBubbleSort(arrToSort() As Variant, Optional nSortOrder As Integer = 0)

    Dim i As Integer, j As Integer, k As Integer, vntTemp As Variant
    Dim inc As Integer, rc As Integer, vntOutput As Variant, arrTemp() As Variant
    Dim nFirst1 As Long, nFirst2 As Long, nLast1 As Long, nLast2 As Long, nSize1 As Long, nSize2 As Long
    
'    If ArrayDimensions(arrToSort) <> 2 Then
'        ' Error message
'        vntOutput = CVErr(xlErrValue)
'        Exit Function
'    End If

    nFirst1 = LBound(arrToSort, 1)
    nLast1 = UBound(arrToSort, 1)
    'nSize1 = nLast1 - nFirst1 + 1
    nFirst2 = LBound(arrToSort, 2)
    nLast2 = UBound(arrToSort, 2)
    'nSize2 = nLast2 - nFirst2 + 1
    
    ReDim arrTemp(0, nFirst2 To nLast2) ' Holds all elements of the row being swapped

    For i = nFirst1 To nLast1 - 1
        For j = i + 1 To nLast1
            If arrToSort(i, 1) > arrToSort(j, 1) Then
                ' Swap every item individually in the row
                For k = nFirst2 To nLast2
                    arrTemp(0, k) = arrToSort(i, k)
                    arrToSort(i, k) = arrToSort(j, k)
                    arrToSort(j, k) = arrTemp(0, k)
                Next
            End If
        Next
    Next
    
    Array2DBubbleSort = True

End Function  ' Array2DBubbleSort()


''
' Capture the substring between the 2 specified delimiters.
'
' The substring starts after the first instance of the open delimiter, and ends before the
' first instance of the close delimiter AFTER the open (ignores earlier instances, if any).
' Delimiters may be one or more characters.
' <p>If the delimiters should be part of the substring, the calling code can reattach them.
' @param The main string from which to extract the substring.
' @param The character(s) denoting the start of the substring.
' @param The character(s) denoting the end of the substring.
' @return Returns the that portion of the main string between the 2 delimiters, not including those delimiters.
' If the delimiters are not found in the input string, returns xlErrNA.
' @group Strings
' @worksheet Yes
' @author Dr. Richard Cook
Function SubstringBetween(ByVal szInput As String, ByVal szOpen As String, _
    ByVal szClose As String)

    Dim nStart As Integer, nEnd As Integer, varOutput As Variant
    Dim szWorking As String, nLengthOpen As Integer
    
    ' Look for open and close delimiters, in that order
    nLengthOpen = Len(szOpen)
    nStart = InStr(szInput, szOpen)
    nEnd = InStr(nStart + nLengthOpen + 1, szInput, szClose)
    
    If nStart = 0 Or nEnd = 0 Then
        ' Error if either not found
        varOutput = CVErr(xlErrNA)
    Else
        ' Delimiters found, get the substring between them
        ' Begin after the end of the start delimiter
        szWorking = Left(szInput, nEnd - 1)
        varOutput = Mid(szWorking, nStart + nLengthOpen, Len(szWorking))
    End If
    
    SubstringBetween = varOutput
End Function


Function SubstringCount(ByVal szInput As String, ByVal szSubstring As String)
' Counts the occurences of the substring in the input string

    Dim nCount As Long, nNext As Long
    nCount = 0
    nNext = 1
    
    While nNext > 0
        nNext = InStr(nNext, szInput, szSubstring)
        If nNext > 0 Then
            nNext = nNext + 1
            nCount = nCount + 1
        End If
    Wend
    
    SubstringCount = nCount
    
End Function


