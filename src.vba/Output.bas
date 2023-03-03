Option Explicit



''
' Creates the default help page for each individual procedure (the topic page).
'
' Each topic page is a stand-alone html page, spelling out all of the information in the
' documentation comments and parsed from the procedure declaration.
' @param The declaration object for the procedure to be documented.
' @param The path to the topic file folder.
' @return No return value.
' @group Output_HTML
' @author Dr. Richard Cook
Function HtmlProcFileCreate(ByRef colGroups As Collection, declCurrent As ClsDeclaration, ByVal szPathOut As String)
    Dim fsoCurrent As FileSystemObject, txtOutput As TextStream, vntRc As Variant, vntCurrent As Variant
    Dim szHtmlOpen As String, szHtmlClose As String, prmCurrent As ClsParameter, szTest As String, szFile As String
    Dim bFormatted As Boolean, arrFacts() As String, vntGroup As Variant, szOutput As String, szExtension As String
    Dim txtInput As TextStream
    
    szHtmlOpen = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbNewLine & "<HTML><HEAD>"
    szHtmlOpen = szHtmlOpen & "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Format.css"">"
    szHtmlClose = "</BODY></HTML>"
    
    Set fsoCurrent = CreateObject("Scripting.FileSystemObject")
    szExtension = ThisWorkbook.Names("txtHtmlExtension").RefersToRange.Cells(1, 1).Text
    Set txtOutput = fsoCurrent.CreateTextFile(szPathOut & declCurrent.Name & szExtension) ' ".htm")
    
    If declCurrent.BriefDescription <> "" Then
        bFormatted = True
    End If
    
    txtOutput.WriteLine (szHtmlOpen)
    txtOutput.WriteLine ("<TITLE>" & declCurrent.ProcName & "</TITLE>")
    txtOutput.WriteLine ("</HEAD>")
    txtOutput.WriteLine ("<BODY>")
    txtOutput.WriteBlankLines (2)
    txtOutput.WriteLine ("<H1>")
    If declCurrent.IsClass = True Then
        txtOutput.WriteLine (declCurrent.Module & ":&nbsp;&nbsp;")
    End If
    txtOutput.WriteLine (declCurrent.ProcName)
    txtOutput.WriteLine ("</H1>")
    txtOutput.WriteBlankLines (1)
    
    ' Output from here on depends on whether input comment block follows formatting/tag conventions
    If bFormatted = True Then
        txtOutput.Write ("<P>" & declCurrent.BriefDescription)
        If declCurrent.Procedure = "Property" Then
            'txtOutput.Write ("  (" & declCurrent.ReadWrite & ")" & vbNewLine)
            txtOutput.Write (vbNewLine & "  (" & declCurrent.ReadWrite & "; ") ' Scope:
            If InStr(declCurrent.ReadWrite, "Read") > 0 Then
                txtOutput.Write (declCurrent.ScopeRead)
            End If
            If InStr(declCurrent.ReadWrite, "/") > 0 Then
                txtOutput.Write ("/")
            End If
            If InStr(declCurrent.ReadWrite, "Write") > 0 Then
                txtOutput.Write (declCurrent.ScopeWrite)
            End If
            txtOutput.Write (")" & vbNewLine)
        Else
            txtOutput.WriteLine (" (Scope: " & declCurrent.Scope & ")")
        End If
        
        ' Deprecated code is presented with a warning in bold red
        If declCurrent.Deprecated <> "" Then
            szOutput = declCurrent.Deprecated
            txtOutput.WriteLine ("<P><b><font color=""red"">Deprecated</font></b> " & szOutput) ' declCurrent.Deprecated)
        End If
        
        txtOutput.WriteBlankLines (2)
    End If
    
    txtOutput.WriteLine ("<H4>Notes</H4>")
    'txtOutput.WriteLine (declCurrent.Description)
    'szTest = Replace(declCurrent.Description, "<EOL>", "<BR>")
    txtOutput.WriteLine ("<P>" & Replace(declCurrent.Description, "<EOL>", "<BR>"))
    
    If declCurrent.Procedure = "Function" And declCurrent.Worksheet <> "" Then
        If declCurrent.Worksheet = "Yes" Then
            szTest = "Available in worksheets."
        ElseIf declCurrent.Worksheet = "No" Then
            szTest = "Not available in worksheets."
        Else
            szTest = "Available in worksheets? " & declCurrent.Worksheet
        End If
        txtOutput.WriteLine ("<P>" & szTest)
    End If
        
    txtOutput.WriteLine ("<H4>Declaration</H4><CODE>")
    If declCurrent.StatementRead <> "" Then
        txtOutput.WriteLine (declCurrent.StatementRead & "<BR>")
    End If
    If declCurrent.StatementWrite <> "" Then
        txtOutput.WriteLine (declCurrent.StatementWrite & "<BR>")
    End If
    If declCurrent.StatementRead = "" And declCurrent.StatementWrite = "" Then
        txtOutput.WriteLine (declCurrent.Declaration)
    End If
    txtOutput.WriteLine ("</CODE>")
    txtOutput.WriteBlankLines (1)
    
    ' List of parameters
    Call HtmlParameters(declCurrent, txtOutput)
    
    If declCurrent.Procedure = "Function" Then
        txtOutput.WriteLine ("<H4>Return Value</H4>")
        txtOutput.WriteLine ("<P>(" & declCurrent.ReturnType & ") ")
        If bFormatted = True Then
            txtOutput.WriteLine (declCurrent.Returns)
        End If
        txtOutput.WriteBlankLines (1)
    End If
    
    
    If declCurrent.Example <> "" Then
'        szOutput = ThisWorkbook.Names("pathProject").RefersToRange.Cells(1, 1).Text
'        szOutput = szOutput & declCurrent.Example
        
        On Error Resume Next
        szOutput = fsoCurrent.GetDrive(declCurrent.Example)
        If szOutput = "" Then
            szOutput = ThisWorkbook.Names("pathProject").RefersToRange.Cells(1, 1).Text
            szOutput = szOutput & declCurrent.Example
        Else
            szOutput = declCurrent.Example
        End If
        
        If fsoCurrent.FileExists(szOutput) Then
            ' If example is a path, insert the target file
            Set txtInput = fsoCurrent.OpenTextFile(szOutput, ForReading)
            szOutput = txtInput.ReadAll '(szOutput)
            txtOutput.WriteLine (szOutput)
            On Error GoTo 0
        Else
            ' If example is text with links, include just the statement, not the target
            txtOutput.WriteLine ("<H4>Example</H4>")
            txtOutput.WriteLine ("<P>" & declCurrent.Example)
                
        End If
    End If
    
    
    ' Link(s) to related functions
    If declCurrent.FunctionGroup <> "" Then
        txtOutput.WriteLine ("<H4>Related </H4>")
        
        ' List specified links before the group, since they are likely more relevant
        ' May be multiple links; must be split. Literal links should be added without interpretation.
        If declCurrent.SeeAlso <> "" Then
            'szOutput = LinkInterpret(declCurrent.SeeAlso, colGroups, declCurrent)
            'szOutput = declCurrent.SeeAlso
            
            arrFacts = Split(declCurrent.SeeAlso, ";")
            txtOutput.WriteLine ("<P>")
            For Each vntCurrent In arrFacts
                'txtOutput.WriteLine ("<P>")
                'txtOutput.WriteLine (vntCurrent & "<br>")
                txtOutput.WriteLine ("<P>" & vntCurrent)
            Next vntCurrent
            'If Not IsError(szOutput) Then
'                txtOutput.WriteLine ("<P>")
'                txtOutput.WriteLine (szOutput & "<br>")
            'End If
        End If
        
        arrFacts = Split(declCurrent.FunctionGroup, ";")
        txtOutput.WriteLine ("<P>")
        For Each vntGroup In arrFacts
            txtOutput.Write ("<a href=""../Groups/")
            txtOutput.Write (Replace(vntGroup, " ", "_"))
            txtOutput.Write (szExtension & """>" & vntGroup & "</a>" & "<br>")
            'txtOutput.Write (".htm"">" & vntGroup & "</a>" & "<br>")
        
        Next vntGroup
        txtOutput.WriteLine ("<P>")
    End If
    
    ' Table of contents link
    szFile = ThisWorkbook.Names("fileContents").RefersToRange.Cells(1, 1).Text '&
'        ThisWorkbook.Names("txtHtmlExtension").RefersToRange.Cells(1, 1).Text
    txtOutput.Write ("<a href=""../" & szFile)
    txtOutput.Write (szExtension & """>Table of Contents</a>" & "<br>")
    txtOutput.WriteLine ("")
    
    
    txtOutput.WriteLine ("<H4>Other Info</H4>")
    txtOutput.WriteLine ("<P><B>Module</B>: " & declCurrent.Module & "<BR>")
    If bFormatted = True Then
        If declCurrent.Author <> "" Then
            txtOutput.WriteLine ("<B>Author</B>: " & declCurrent.Author & "<BR>")
        End If
        If declCurrent.Created <> "" Then
            txtOutput.WriteLine ("<B>Created</B>: " & declCurrent.Created & "<BR>")
        End If
    End If
    
    ' Document updated date @docdate (not yet implemented)
    If declCurrent.DocumentedDate <> "" Then
        txtOutput.WriteLine ("<P class=""small"">This page last updated " & declCurrent.DocumentedDate)
    End If

    'txtOutput.WriteLine ("<P class=""right"">Documentation generated " & Date & ", using DocuVeeb v. 1.10.08")
    txtOutput.WriteLine ("<P class=""right small"">Generated " & Date & ", using DocuVeeb " & VersionId())
    
    txtOutput.WriteLine (szHtmlClose)
    txtOutput.Close
    
End Function  ' HtmlProcFileCreate()



''
' Creates the XML name-value pair HTML Help expects.
'
' @param Human-readable code name; underscores are replaced with spaces.
' @param Number of indent levels for the first line of the group (the "<LI>" entry).
' Following lines in the object are indented an extra level beyond the first.
' @return Returns the complete text of the name-value pair in XML, ready to be written to the Help contents file.
' @group Output_Windows_Help
Function HelpContentNameLocalObject(ByVal szName As String, ByVal szNameFull As String, ByVal nLevels As Integer)

    Dim szOutput As String, szTabs As String, szExtension As String
    
    ' Set the desired number of tabs to indent the first line
    ' Subsequent lines are indented an extra tab beyond the first
    szTabs = String(nLevels, vbTab)
    szExtension = ThisWorkbook.Names("txtHtmlExtension").RefersToRange.Cells(1, 1).Text
    
    ' Write the XML for the name-value pair
    szOutput = szTabs & "<LI> <OBJECT type=""text/sitemap"">"
    szOutput = szOutput & vbNewLine & szTabs & vbTab & "<param name=""Name"" value=""" & Replace(szName, "_", " ") & """>"
    'szOutput = szOutput & vbNewLine & szTabs & vbTab & "<param name=""Local"" value=""" & szNameFull & ".htm"">"
    szOutput = szOutput & vbNewLine & szTabs & vbTab & "<param name=""Local"" value=""" & szNameFull & szExtension & """>"
    szOutput = szOutput & vbNewLine & szTabs & vbTab & "</OBJECT>"
    
    HelpContentNameLocalObject = szOutput
    
End Function  ' HelpContentNameLocalObject()


''
' Creates an html link from display text and a filename, without extension.
'
' Can be used for any combination of name and associated file.
' <P>This function is called to generate links in the default HTML contents file.
' @param The text to be displayed in the link.
' @param The link target filename, without the ".htm" extension.
' A path may be included with the filename.
' @return Returns the HTML code for the link, from the opening "<a href=...>" to the closing "</a>".
' @group Output_HTML
' @author Dr. Richard Cook
Function HtmlNameLink(ByVal szDisplayText As String, ByVal szFilename As String)

    Dim szOutput As String, szExtension As String
    
    ' Write the html link
    szExtension = ThisWorkbook.Names("txtHtmlExtension").RefersToRange.Cells(1, 1).Text
    'szOutput = "<a href=""" & szFilename & ".htm"" >"
    szOutput = "<a href=""" & szFilename & szExtension & """ >"
    szOutput = szOutput & szDisplayText & "</a>"
    
    HtmlNameLink = szOutput
    
End Function  ' HtmlNameLink()


''
' Generates the "Parameters" section in the default HTML output for the procedure's Topic page.
'
' The section includes the title and a table of the parameters in the input declaration.
' The table includes a header row.
' If the declaration includes a ParamArray, a footnote explaining a ParamArray follows the table.
' @param The declaration object whose parameters are to be described.
' @param The topic file being written.
' @return No return value.
' @group Output_HTML
' @author Dr. Richard Cook
Function HtmlParameters(ByRef declCurrent As ClsDeclaration, txtOutput As TextStream)

    Dim nParamCount As Integer, szOutput As String, vntCurrent As Variant, bHasParamArray As Boolean
    Dim szAsterisk As String
    
    nParamCount = declCurrent.Parameters.Count
    bHasParamArray = False
    
    If nParamCount > 0 Then
        ' Start table; Print headers, if any
        txtOutput.WriteLine ("<H4>Parameters </H4>")
        txtOutput.WriteLine ("<TABLE border=1>")  ' <caption>ASR</caption>

        ' Write the header line for the parameter table
        txtOutput.WriteLine ("<TR>")
        txtOutput.WriteLine ("<TH><p>Name</TH>")
        txtOutput.WriteLine ("<TH><p>Type</TH>")
        txtOutput.WriteLine ("<TH><p>Required</TH>")
        txtOutput.WriteLine ("<TH><p>Passed</TH>")
        txtOutput.WriteLine ("<TH><p>Remarks</TH>")
        txtOutput.WriteLine ("</TR>")
        
        For Each vntCurrent In declCurrent.Parameters
            If vntCurrent.IsParamArray = True Then
                bHasParamArray = True
                szAsterisk = "*"
            Else
                szAsterisk = ""
            End If
            
            ' Name, type, req'd/optional, passing, remarks
            txtOutput.WriteLine ("<TD><p>" & vntCurrent.Name & szAsterisk & "</TD>")
            txtOutput.WriteLine ("<TD><p>" & vntCurrent.VariableType & "</TD>")  ' VariableType
            
            If vntCurrent.OptionalInput = True Then
                szOutput = "Optional"
            Else
                szOutput = "Required"
            End If
            txtOutput.WriteLine ("<TD><p>" & szOutput & "</TD>")
            
            If vntCurrent.PassByValue = True Then
                szOutput = "ByVal"
            Else
                szOutput = "ByRef"
            End If
            txtOutput.WriteLine ("<TD><p>" & szOutput & "</TD>")  ' PassByValue
            
            If vntCurrent.Description = "" Then
                szOutput = "&nbsp;"
            Else
                szOutput = vntCurrent.Description
                If vntCurrent.DefaultValue <> "" Then
                    szOutput = szOutput & " Default value: " & vntCurrent.DefaultValue
                End If
            End If
            txtOutput.WriteLine ("<TD><p>" & szOutput & "</TD>")
            txtOutput.WriteLine ("</TR>")
        Next vntCurrent
        
        txtOutput.WriteLine ("</TABLE>" & vbNewLine)
        
        If bHasParamArray = True Then
            ' Write footnote explaining ParamArray
            txtOutput.WriteLine ("<P class=""small"">&nbsp;&nbsp;* Parameter is a ParamArray, an indefinite number of parameters,")
            txtOutput.WriteLine ("from none to as many as needed. An actual array is not required.</p>")
        End If
    End If

End Function  ' HtmlParameters()


''
' Creates the default Table of Contents file for a Windows HTML Help project.
'
' Each group appears as a chapter in the Windows Help file, according to the @group tags.
' With classes, the class itself forms the chapter, and any @group tags are ignored.
' @param The set of groups to be written.
' @param The filename with full path for the output group file.
' @return No return value. #### Should return success/failure information. ####
' @group Output_Windows_Help
' @author Dr. Richard Cook
Function HelpContentFileCreate(ByRef colGroups As Collection, ByVal szPathOut As String)
    Dim fsoCurrent As FileSystemObject, txtOutput As TextStream, vntRc As Variant, nCurrentGrp As Integer
    Dim szHtmlOpen As String, szHtmlClose As String, grpCurrent As ClsGroup, szOutput As String
    Dim bFormatted As Boolean, declCurrent As ClsDeclaration, arrGroupNames() As Variant
    Dim nLastGroup As Integer, nLastProc As Integer, arrProcNames() As Variant, nCurrentProc As Integer
    Dim bIncludePrivate As Boolean
    
    szHtmlOpen = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbNewLine & "<HTML>" & vbNewLine & "<HEAD>"
    szHtmlOpen = szHtmlOpen & "<meta name=""GENERATOR"" content=""AutoDocumenter"">" & vbNewLine & "<!-- Sitemap 1.0 -->"
    szHtmlOpen = szHtmlOpen & "</HEAD><BODY>" & vbNewLine & "<OBJECT type=""text/site properties"">"
    szHtmlOpen = szHtmlOpen & vbNewLine & vbTab & "<param name=""Auto Generated"" value=""Yes"">" & vbNewLine & "</OBJECT>"

    szHtmlClose = "</BODY></HTML>"
    
    Set fsoCurrent = CreateObject("Scripting.FileSystemObject")
    Set txtOutput = fsoCurrent.CreateTextFile(szPathOut)
    
    ' Determine whether to include non-public declarations (Private, Friend or Project)
    bIncludePrivate = ThisWorkbook.Names("optionIncludePrivate").RefersToRange.Cells(1, 1).Text

    txtOutput.WriteLine (szHtmlOpen)
    
    vntRc = CollectionNames(colGroups, arrGroupNames)
    nLastGroup = colGroups.Count
    
    'For Each grpCurrent In colGroups
    For nCurrentGrp = 1 To nLastGroup
        Set grpCurrent = colGroups(arrGroupNames(nCurrentGrp, 2))
        nLastProc = grpCurrent.Members.Count

        ' Write the group open by starting a new UL (unordered list)
        txtOutput.WriteLine ("<UL>")
        szOutput = HelpContentNameLocalObject(grpCurrent.Name, "Groups\" & grpCurrent.Code, 1)
        txtOutput.WriteLine (szOutput & vbNewLine)
        txtOutput.WriteLine (vbTab & "<UL>")
        
        vntRc = CollectionNames(grpCurrent.Members, arrProcNames)
        ' Loop through the group's member procedures
        'For Each declCurrent In grpCurrent.Members
        For nCurrentProc = 1 To nLastProc
            ' Add each declaration as a list item
            Set declCurrent = grpCurrent.Members(arrProcNames(nCurrentProc, 2))
            'szOutput = HelpContentNameLocalObject(declCurrent.Name, 2)
            If declCurrent.Scope = "Public" Or bIncludePrivate = True _
                    Or declCurrent.ScopeRead = "Public" Then
                szOutput = HelpContentNameLocalObject(declCurrent.ProcName, "Topics\" & declCurrent.Name, 2)
                txtOutput.WriteLine (szOutput & vbNewLine)
            End If
        
        Next 'declCurrent
        
        ' Close the group
        txtOutput.WriteLine (vbTab & "</UL>")
        txtOutput.WriteLine ("</UL>")
        
    Next nCurrentGrp
    
    
    txtOutput.WriteLine (szHtmlClose)
    txtOutput.Close

End Function  ' HelpContentFileCreate


''
' Creates the project file for a Windows HTML Help project.
'
' Each group appears as a chapter in the Windows Help file, according to the @group tags.
' With classes, the class itself forms the chapter, and any @group tags are ignored.
' @param The set of groups to be written.
' @param The filename with full path for the output group file.
' @return No return value. #### Should return success/failure information. ####
' @group Output_Windows_Help
' @author Dr. Richard Cook
Function HelpProjectFileCreate(ByRef colGroups As Collection, ByVal szPathOut As String)
    Dim fsoCurrent As FileSystemObject, txtOutput As TextStream, vntRc As Variant, nCurrentGrp As Integer
    Dim szHtmlOpen As String, szHtmlClose As String, grpCurrent As ClsGroup, szOutput As String
    Dim bFormatted As Boolean, declCurrent As ClsDeclaration, arrGroupNames() As Variant
    Dim nLastGroup As Integer, nLastProc As Integer, arrProcNames() As Variant, nCurrentProc As Integer
    Dim rngRow As Range, szExtension As String, bIncludePrivate As Boolean
    
    Set fsoCurrent = CreateObject("Scripting.FileSystemObject")
    Set txtOutput = fsoCurrent.CreateTextFile(szPathOut)
    szExtension = ThisWorkbook.Names("txtHtmlExtension").RefersToRange.Cells(1, 1).Text
    
    ' Determine whether to include non-public declarations (Private, Friend or Project)
    bIncludePrivate = ThisWorkbook.Names("optionIncludePrivate").RefersToRange.Cells(1, 1).Text
    
    ' Write opening section "[OPTIONS]", using the parameter page definitions
    txtOutput.WriteLine ("[OPTIONS]")
    For Each rngRow In ThisWorkbook.Names("txtHhpOptions").RefersToRange.Rows
        txtOutput.WriteLine (rngRow.Cells(1, 1) & "=" & rngRow.Cells(1, 2))
    Next rngRow
    
    vntRc = CollectionNames(colGroups, arrGroupNames)
    nLastGroup = colGroups.Count
    
    ' Write the list of files in the "[FILES]" section
    txtOutput.WriteBlankLines (1)
    txtOutput.WriteLine ("[FILES]")
    
    For nCurrentGrp = 1 To nLastGroup
        Set grpCurrent = colGroups(arrGroupNames(nCurrentGrp, 2))
        nLastProc = grpCurrent.Members.Count

        vntRc = CollectionNames(grpCurrent.Members, arrProcNames)
        ' Loop through the group's member procedures
        For nCurrentProc = 1 To nLastProc
            ' Add each declaration as a list item
            Set declCurrent = grpCurrent.Members(arrProcNames(nCurrentProc, 2))
            'szOutput = HelpContentNameLocalObject(declCurrent.Name, 2)
            'If declCurrent.Scope <> "Private" Then
            If declCurrent.Scope = "Public" Or bIncludePrivate = True _
                    Or declCurrent.ScopeRead = "Public" Then
                'szOutput = HelpContentNameLocalObject(declCurrent.ProcName, "Topics\" & declCurrent.Name, 2)
                txtOutput.WriteLine ("Topics\" & declCurrent.Name & szExtension)
            End If
        
        Next 'declCurrent
        
    Next nCurrentGrp
    
    txtOutput.Close

End Function  ' HelpProjectFileCreate


''
' Creates the default HTML Table of Contents file for a web-style project.
'
' Groups are listed in alphabetical order, with all procedures in the group listed in
' a table for the group as a whole. The procedures are also alphabetized.
' With classes, the class itself forms the group, and any @group tags are ignored.
' @param The set of groups to be written.
' @param The filename with full path for the output group file.
' @param Text to be used in the <TITLE> tab for the web page.
' @return No return value. #### Should return success/failure information. ####
' @group Output_HTML
' @author Dr. Richard Cook
Function HtmlContentFileCreate(ByRef colGroups As Collection, ByVal szPathOut As String, _
    Optional ByVal szTitle As String)
    Dim fsoCurrent As FileSystemObject, txtOutput As TextStream, vntRc As Variant, nCurrentGrp As Integer
    Dim szHtmlOpen As String, szHtmlClose As String, grpCurrent As ClsGroup, szOutput As String
    Dim bFormatted As Boolean, declCurrent As ClsDeclaration, arrGroupNames() As Variant
    Dim nLastGroup As Integer, nLastProc As Integer, arrProcNames() As Variant, nCurrentProc As Integer
    Dim bIncludePrivate As Boolean
    
    szHtmlOpen = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbNewLine & "<HTML><HEAD>"
    szHtmlOpen = szHtmlOpen & vbNewLine & "<meta name=""GENERATOR"" content=""AutoDocumenter"">" & vbNewLine
    szHtmlOpen = szHtmlOpen & "<link rel=""stylesheet"" type=""text/css"" href=""Include/Format.css"">" & vbNewLine
    szHtmlOpen = szHtmlOpen & "</HEAD><BODY>"

    szHtmlClose = "</BODY></HTML>"
    
    Set fsoCurrent = CreateObject("Scripting.FileSystemObject")
    Set txtOutput = fsoCurrent.CreateTextFile(szPathOut) ' & declCurrent.Name & ".htm")
    
    ' Determine whether to include non-public declarations (Private, Friend or Project)
    bIncludePrivate = ThisWorkbook.Names("optionIncludePrivate").RefersToRange.Cells(1, 1).Text
    
    txtOutput.WriteLine (szHtmlOpen)
    
    vntRc = CollectionNames(colGroups, arrGroupNames)
    nLastGroup = colGroups.Count
    
    'ProjectDisplayName
    txtOutput.WriteLine ("<H1>" & szTitle & "</H1>")
    
    'For Each grpCurrent In colGroups
    For nCurrentGrp = 1 To nLastGroup
        Set grpCurrent = colGroups(arrGroupNames(nCurrentGrp, 2))
        nLastProc = grpCurrent.Members.Count

        ' Write the group open by starting a new table, with group name as introduction ####
        txtOutput.WriteBlankLines (1)
        txtOutput.WriteLine ("<P>")
        szOutput = HtmlNameLink(grpCurrent.Name, "Groups\" & grpCurrent.Code)
        txtOutput.WriteLine (szOutput)
        txtOutput.WriteLine ("<table border=1>")
        
        vntRc = CollectionNames(grpCurrent.Members, arrProcNames)
        ' Loop through the group's member procedures
        For nCurrentProc = 1 To nLastProc
            ' Add each declaration as a list item
            Set declCurrent = grpCurrent.Members(arrProcNames(nCurrentProc, 2))
            'szOutput = HelpContentNameLocalObject(declCurrent.Name, 2)
            'If declCurrent.Scope <> "Private" Then
            If declCurrent.Scope = "Public" Or bIncludePrivate = True _
                    Or declCurrent.ScopeRead = "Public" Then
                txtOutput.Write ("<TR>")
                txtOutput.Write ("<TD><P>")
                szOutput = HtmlNameLink(declCurrent.ProcName, "Topics\" & declCurrent.Name)
                txtOutput.Write (szOutput)
                txtOutput.Write ("<TD><P>")
                txtOutput.Write (declCurrent.BriefDescription)
                txtOutput.Write ("</TR>")
            End If
        
        Next nCurrentProc
        
        ' Close the group
        txtOutput.WriteLine ("</table>")
        
    Next nCurrentGrp
    
    txtOutput.WriteLine ("<P class=""right small"">Generated " & Date & ", using DocuVeeb " & VersionId())
    
    txtOutput.WriteLine (szHtmlClose)
    txtOutput.Close

End Function  ' HtmlContentFileCreate


''
' Creates a separate HTML file for each group, listing the procedures within that group.
'
' @param The list of groups, containing all declarations that will be included.
' @param The path to the output files; filenames will be added, based on the group name.
' @return No return value. #### Should return success/failure information. ####
' @group Output_HTML
' @author Dr. Richard Cook
Function HtmlGroupFilesCreate(ByRef colGroups As Collection, ByVal szPathOut As String)
    Dim fsoCurrent As FileSystemObject, txtOutput As TextStream, vntRc As Variant, nCurrentGrp As Integer
    Dim szHtmlOpen As String, szHtmlClose As String, grpCurrent As ClsGroup, szOutput As String, szFile As String
    Dim bFormatted As Boolean, declCurrent As ClsDeclaration, arrGroupNames() As Variant, szExtension As String
    Dim nLastGroup As Integer, nLastProc As Integer, arrProcNames() As Variant, nCurrentProc As Integer
    Dim bIncludePrivate As Boolean
    
    szHtmlOpen = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbNewLine & "<HTML><HEAD>"
    szHtmlOpen = szHtmlOpen & "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Format.css"">"
    
    szHtmlClose = "</BODY></HTML>"
    szExtension = ThisWorkbook.Names("txtHtmlExtension").RefersToRange.Cells(1, 1).Text
    
    ' Determine whether to include non-public declarations (Private, Friend or Project)
    bIncludePrivate = ThisWorkbook.Names("optionIncludePrivate").RefersToRange.Cells(1, 1).Text
    
    Set fsoCurrent = CreateObject("Scripting.FileSystemObject")
    
    vntRc = CollectionNames(colGroups, arrGroupNames)
    nLastGroup = colGroups.Count
    
    'For Each grpCurrent In colGroups
    For nCurrentGrp = 1 To nLastGroup
        Set grpCurrent = colGroups(arrGroupNames(nCurrentGrp, 2))
        nLastProc = grpCurrent.Members.Count

        Set txtOutput = fsoCurrent.CreateTextFile(szPathOut & grpCurrent.Code & szExtension)  ' & declCurrent.Name)
        
        txtOutput.WriteLine (szHtmlOpen)
        txtOutput.WriteLine ("<TITLE>" & grpCurrent.Name & "</TITLE>")
        txtOutput.WriteLine ("</HEAD>")
        txtOutput.WriteLine ("<BODY>")
        txtOutput.WriteBlankLines (2)
        txtOutput.WriteLine ("<H1>")
        If grpCurrent.ModuleType = "Class" Then
            txtOutput.WriteLine ("Class" & ":&nbsp;&nbsp;")
        End If
        txtOutput.WriteLine (grpCurrent.Name)
        txtOutput.WriteLine ("</H1>")
        txtOutput.WriteBlankLines (1)
        
        If grpCurrent.Description <> "" Then
            txtOutput.WriteLine ("<P>" & grpCurrent.Description & "</P>")
        End If
        
        ' Begin table, w/ header row
        ' Write the header line
        txtOutput.WriteLine ("<TABLE border=1>")
        txtOutput.WriteLine ("<TR>")
        txtOutput.WriteLine ("<TH><p>Name</TH>")
        txtOutput.WriteLine ("<TH><p>Description</TH>")
        txtOutput.WriteLine ("</TR>")
            
        vntRc = CollectionNames(grpCurrent.Members, arrProcNames)
        ' Loop through the group's member procedures
        'For Each declCurrent In grpCurrent.Members
        For nCurrentProc = 1 To nLastProc
            ' Add each declaration as a link, with brief description
            Set declCurrent = grpCurrent.Members(arrProcNames(nCurrentProc, 2))
            'If declCurrent.Scope <> "Private" Then
            If declCurrent.Scope = "Public" Or bIncludePrivate = True _
                    Or declCurrent.ScopeRead = "Public" Then
                txtOutput.WriteLine ("<TR>")
                txtOutput.Write ("<TD><p>")
                txtOutput.Write ("<a href=""../Topics/")
                txtOutput.Write (declCurrent.Name)
                'txtOutput.Write (".htm"">" & declCurrent.ProcName & "</a>")
                txtOutput.Write (szExtension & """>" & declCurrent.ProcName & "</a>")
                txtOutput.Write ("</TD>")
            
                ' Only formatted comments provide a brief description; leave blank for unformatted comments
                If declCurrent.BriefDescription <> "" Then
                    szOutput = declCurrent.BriefDescription
                Else
                    szOutput = "&nbsp;"
                End If
                txtOutput.WriteLine ("<TD><p>" & szOutput & "</TD>")
            End If
            
            ' End the table row
            txtOutput.WriteLine ("</TR>")
            
        Next nCurrentProc
        
        ' Close the group
        txtOutput.WriteLine ("</TABLE><br><br>" & vbNewLine) ' <caption>ASR</caption>

    
        ' Table of contents link
        szFile = ThisWorkbook.Names("fileContents").RefersToRange.Cells(1, 1).Text '&
    '        ThisWorkbook.Names("txtHtmlExtension").RefersToRange.Cells(1, 1).Text
        txtOutput.Write ("<P><a href=""../" & szFile)
        txtOutput.Write (szExtension & """>Table of Contents</a>") '& "<br>")
        txtOutput.WriteLine ("")
        
        txtOutput.WriteLine ("<P class=""right small"">Generated " & Date & ", using DocuVeeb " & VersionId())
        
        txtOutput.WriteLine (szHtmlClose)
        txtOutput.Close
    
    Next nCurrentGrp
    
End Function  'HtmlGroupFilesCreate()


''
' Determines whether the named item exists in the collection.
'
' Cycles through items in the collection looking for the specified name. The comparison is case sensitive.
' <p>Has been tested with the Worksheets and Names (ranges) collections; has been used but not tested extensively in other collections.
' @param The name of the collection.
' @param The name of the item being tested for existence.
' @return Returns True if the named item is found in the collection, or False if it is not.
' @group Collections
' @author Dr. Richard Cook
' @date 4/11/13
Function ItemExists(varCollection As Variant, szName As String)

    Dim bFound As Boolean, varCurrent As Variant
    
    bFound = False
    
    For Each varCurrent In varCollection
        If varCurrent.Name = szName Then
            bFound = True
            Exit For
        End If
    Next
    
    ItemExists = bFound

End Function


''
' Determines the read/write status of a property in a VBA class.
'
' Properties may be created as read only, read/write, or (rarely) write only.
' @param The name of the class module.
' @param The name of the property being tested.
' @return Returns one of the following strings, as appropriate:<br>
' "Read Only"<br>
' "Write Only"<br>
' "Read/Write"<br>
' xlErrNA on error<br>
' @group Classes
' @author Dr. Richard Cook
' @date 10/16/2015
Function PropertyIsReadWrite(cmModuleCur As CodeModule, szProcName As String) As Variant   ' vbext_pk_Get, Set, Let
    Dim nLineDeclare As Long, nOutput As Integer, szOutput As String, vntOutput As Variant
    nOutput = 0
    szOutput = ""
    
    ' Accessing a non-existent property will generate an error.
    ' Attempt to access each type and see whether it generates an error.
    On Error Resume Next
    Err.Clear
    nLineDeclare = cmModuleCur.ProcBodyLine(szProcName, vbext_pk_Get)
    If Err.Number = 0 Then
        nOutput = nOutput + 1
    End If
    
    Err.Clear
    nLineDeclare = cmModuleCur.ProcBodyLine(szProcName, vbext_pk_Set)
    If Err.Number = 0 Then
        nOutput = nOutput + 2
    Else
        Err.Clear
        nLineDeclare = cmModuleCur.ProcBodyLine(szProcName, vbext_pk_Let)
        If Err.Number = 0 Then
            nOutput = nOutput + 2
        End If
    End If
    On Error GoTo 0
    
    Select Case nOutput
        Case 0:
            vntOutput = CVErr(xlErrNA)
        Case 1:
            vntOutput = "Read Only"
        Case 2:
            vntOutput = "Write Only"
        Case 3:
            vntOutput = "Read/Write"
        Case Else:
            vntOutput = CVErr(xlErrNA)
    End Select
    
    PropertyIsReadWrite = vntOutput

End Function  ' PropertyIsReadWrite()


'                        ' Write the function/macro name, etc., to the CodeList sheet
'                        wkshOutput.Cells(nRowCurOutput, 1).Value = declCurrent.ProcName
'                        wkshOutput.Cells(nRowCurOutput, 2).Value = declCurrent.Declaration
'                        wkshOutput.Cells(nRowCurOutput, 3).Value = declCurrent.Description
'                        wkshOutput.Cells(nRowCurOutput, 4).Value = declCurrent.Module
'                        wkshOutput.Cells(nRowCurOutput, 6).Value = declCurrent.Scope
'
'                        If declCurrent.Macro = True Then
'                            wkshOutput.Cells(nRowCurOutput, 5).Value = "Macro"
'                        Else
'                            wkshOutput.Cells(nRowCurOutput, 5).Value = "Function"
'                        End If


' calcThisVersion
Function VersionId()
    Dim vntVersion As String, rngVersion As Range
    
    vntVersion = ThisWorkbook.Names("calcThisVersion").RefersToRange.Text
    
    VersionId = vntVersion

End Function

