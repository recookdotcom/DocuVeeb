Option Explicit


' ClsEnvironmentNew() in Utility (PMO)
' Set reference to PMO Public Library
' Properties: capture both declarations if read/write
'   Type is return for Get, param for Let/Set
'   Param types are all for Get, all but last for Let/Set
'   Return is return for Get, param for Let/Set
'   If Let/Set is parsed, return is set to variant by default, since no type specified


' #### If sub: no return value to help file ####
' #### But can be used to describe the sub's explected outcome

''
' @groupnote The Strings function provide a limited number of common actions useful in parsing strings.
' @group Strings
''

''
' @groupnote The Paths function provide a limited number of common actions useful in handling paths.
' @group Paths
''

Private nNonsense As Integer ' Forces above to be a declaration section, to avoid inclusion in DocuVeeb output


''
' Executes the documentation gathering and default output routines.
'
' This routine is the main program that runs all others.
' @example As a test, see the example in {@link ProcGetModule}
' @group DocuVeeb
' @author Dr. Richard Cook
Sub DocuVeebMain()

    Dim vbpProjectCur As VBIDE.VBProject, vbcComponentCur As VBComponent, cmModuleCur As CodeModule
    Dim nLineStart As Long, szOutput As String, szProcName As String, nLineNext As Long, nLineCnt As Integer
    Dim nColumnCur As Long, nRowCurOutput As Long, arrFacts() As String
    Dim szLineCur As String, nLineDeclare As Long, declCurrent As ClsDeclaration, vntRc As Variant
    Dim colDeclarations As Collection, bFormatted As Boolean, nCountLines As Integer, szDescription As String
    Dim nPosition As Integer, colGroups As Collection, vntCurrent As Variant, bFound As Boolean
    Dim nItem As Integer, grpCurrent As ClsGroup, colNewItem As Collection, szPath As String
    Dim wkbkThis As Workbook, szFile As String, vntGroup As Variant, bClassModule As Boolean
    Dim nProcKind As Long, szReadWrite As String, szDeclareName As String, szGroupName As String, szGroupDescription As String
    Dim arrDeclarationSection As Variant, bOptionPrivateModule As Boolean, bIncludePrivate As Boolean
    'Dim arrDeclarationSection() As String, bOptionPrivateModule As Boolean, bIncludePrivate As Boolean
    
    Set colDeclarations = New Collection
    Set colGroups = New Collection
    
    ' Capture the target project from the input list (fileVbaProject)
    szFile = ThisWorkbook.Names("fileVbaProject").RefersToRange.Text
    
    ' If no project listed, use the active workbook
    ' Expand to mulitple projects with .Project
    If szFile = "" Or szFile = "Active" Then
        Set vbpProjectCur = ActiveWorkbook.VBProject
    ElseIf szFile = "All Open" Then
        '
    Else
        nPosition = Len(szFile)
        For Each wkbkThis In Workbooks
            If Left(wkbkThis.Name, nPosition) = szFile Then
                ' Found the target workbook
                Set vbpProjectCur = wkbkThis.VBProject
                bFound = True
                Exit For
            End If
        Next wkbkThis
        If bFound = False Then
            Call MsgBox("Open workbook and rerun this macro.", vbOKOnly, "Cannot find " & szFile & " workbook")
            Exit Sub
        End If
    End If
    
    nRowCurOutput = 2
    ' Loop through the VB components (modules)
    For Each vbcComponentCur In vbpProjectCur.VBComponents
        ' Ignore modules whose name ends with underscore (the signal to be ignored)
        If Right(vbcComponentCur.Name, 1) <> "_" Then
            Set cmModuleCur = vbcComponentCur.CodeModule
            
            If vbcComponentCur.Name = "DeclareSection" Then
                nRowCurOutput = nRowCurOutput  ' #### DEBUG; allows breakpoint ####
            End If
            
            ' Standard/class modules, no forms, etc.
            If vbcComponentCur.Type = vbext_ct_StdModule Or vbcComponentCur.Type = vbext_ct_ClassModule Then
                If vbcComponentCur.Type = vbext_ct_ClassModule Then
                    bClassModule = True
                Else
                    bClassModule = False
                End If
                
                bOptionPrivateModule = False
                
                ' Start line will advance to the start of each proc
                nLineStart = cmModuleCur.CountOfDeclarationLines + 1
                
                ' Capture all lines in the Declaration Section for later processing
                ' Do line by line in order to handle Option Private Module statement
                ' Private Module affects scope declaration for individual procs
                If nLineStart > 1 Then
                    ReDim arrDeclarationSection(1 To nLineStart - 1)
                    For nLineCnt = 1 To nLineStart - 1
                        arrDeclarationSection(nLineCnt) = Trim(cmModuleCur.Lines(nLineCnt, 1))
                        If InStr(arrDeclarationSection(nLineCnt), "Option Private Module") = 1 Then
                            bOptionPrivateModule = True
                        End If
                    Next nLineCnt
                Else
                    ' No Declaration Section in this module
                    Erase arrDeclarationSection
                End If

                
                ' ####
                ' # Post-Declare section; individual functions/subs
                ' ####
                
                ' Process declare lines along with procs, from start
                ' If declare section, don't use built-in statements:
                ' ProcOfLine, ProcBodyLine, ProcCountLines
                
                
                ' nLineStart is the first blank line or proc line after the prior declaration
                Do Until nLineStart >= cmModuleCur.CountOfLines
                    ' Determine name of next procedure
                    szProcName = cmModuleCur.ProcOfLine(nLineStart, nProcKind)  ' vbext_pk_Get, Set, Let
                    
                    If bClassModule = True Then
                        ' Class names sort w/i the class; use class prefix to group together
                        szDeclareName = vbcComponentCur.Name & "." & szProcName
                    Else
                        ' Other names sort by name overall; use module suffix to avoid grouping by module
                        szDeclareName = szProcName & "." & vbcComponentCur.Name
                    End If
                    
                    
                    If ItemExists(colDeclarations, szDeclareName) = False Then
                        ' New procedure declaration unless this is the 2nd statement of an already-declared property
                        Set declCurrent = New ClsDeclaration
                        declCurrent.ProcName = szProcName
                        declCurrent.Name = szDeclareName
                        declCurrent.Project = vbpProjectCur.Name
                        declCurrent.IsClass = bClassModule
                        declCurrent.IsOptionPrivate = bOptionPrivateModule
                        If nProcKind <> vbext_pk_Proc Then
                            ' Class member properties: determine read/write status, other properties
                            Call declCurrent.PropertyFillReadWrite(cmModuleCur)
                        End If
                        bFormatted = False
                        
                        ' Look for the declaration of the next procedure
                        ' (ProcStartLine: 1st line after prior proc) ProcBodyLine gives line number for decl.
                        nLineDeclare = cmModuleCur.ProcBodyLine(szProcName, nProcKind)
                        
                        ' Set array to capture the pre-declaration comments
                        nCountLines = nLineDeclare - nLineStart
                        If nCountLines < 0 Then
                            Exit Do
                        ElseIf nCountLines > 0 Then
                            ReDim arrFacts(1 To nCountLines)
                            
                            ' DESCRIPTION, PRE-DECLARATION
                            ' Capture all leading comments (preceeding proc declaration)
                            'szOutput = cmModuleCur.Lines(nLineStart, nLineDeclare - nLineStart)
                            nLineNext = nLineStart
                            nLineCnt = 1
                            szDescription = ""
                            szLineCur = ""
                            While nLineNext <= cmModuleCur.CountOfLines _
                                    And (cmModuleCur.Lines(nLineNext, 1) = "" _
                                    Or Left(Trim(cmModuleCur.Lines(nLineNext, 1)), 1) = "'")
                                szLineCur = Trim(cmModuleCur.Lines(nLineNext, 1))
                                arrFacts(nLineCnt) = szLineCur
                                If Left(szLineCur, 1) = "'" Then
                                    ' Look for start of formatted block
                                    ' If found, remove all comments preceeding the block
                                    If szLineCur = "''" And nLineCnt > 0 Then
                                        If arrFacts(nLineCnt - 1) = "" Then
                                            bFormatted = True
                                            szDescription = ""
                                        End If
                                    End If
                                    ' Remove the comment mark and leading spaces before text following comment mark
                                    szLineCur = Trim(Mid(szLineCur, 2))
                                End If
                                If szLineCur <> "" Then
                                    If Left(szLineCur, 1) = "@" Then
                                        bFormatted = True
                                    End If
                                    If szDescription <> "" Then
                                        szLineCur = "<EOL>" & vbNewLine & szLineCur
                                    End If
                                    szDescription = szDescription & szLineCur
                                End If
                                nLineNext = nLineNext + 1
                                nLineCnt = nLineCnt + 1
                            Wend
                        End If
                        
                        ' PROCEDURE DECLARATION
                        ' Capture multi-line declarations, denoted by trailing underscore
                        ' Retains internal underscores, drops only the trailing underscore
                        szOutput = DeclarationCapture(cmModuleCur, nLineDeclare, nLineCnt)
                        
                        ' Capture the complete declaration
                        vntRc = declCurrent.ProcStatementParse(szOutput)
                        
                        ' Put description parts into the declaration object, now that parameter list has been formed,
                        ' which instantiates the object.
                        ' If the block is not formatted, captures all leading comments and those following the declaration,
                        ' up to the first line of code.
                        If bFormatted = True Then
                            vntRc = declCurrent.CommentsParse(arrFacts) ' ##########
                        Else
                            declCurrent.Description = szDescription
                        End If
                        
                        ' DESCRIPTION, POST-DECLARATION
                        ' Capture comments following proc declaration; stop at first non-blank, non-comment line
                        ' Skip post-decl comments if the pre-decl comment block is formatted
                        If bFormatted = False Then
                            nLineNext = nLineDeclare + nLineCnt
                            nLineCnt = 1
                            szOutput = ""
                            szLineCur = ""
                            While nLineNext <= cmModuleCur.CountOfLines _
                                    And (cmModuleCur.Lines(nLineNext, 1) = "" _
                                    Or Left(Trim(cmModuleCur.Lines(nLineNext, 1)), 1) = "'")
                                szLineCur = Trim(cmModuleCur.Lines(nLineNext + nLineCnt - 1, 1))
                                If szLineCur <> "" Then
                                    If Left(szLineCur, 1) = "'" Then
                                        szLineCur = Trim(Mid(szLineCur, 2))
                                    End If
                                    szOutput = szOutput & szLineCur & "<EOL>" & vbNewLine  ' "<EOL>" "<EOP>" ?
                                    'szOutput = szOutput & szLineCur & vbNewLine
                                End If
                                nLineNext = nLineNext + 1
                            Wend
                            declCurrent.Description = declCurrent.Description & szOutput
                        End If
                        
                        ' Add each proc declaration to the collection
                        colDeclarations.Add Item:=declCurrent, Key:=declCurrent.Name '& "1"
                        declCurrent.Module = vbcComponentCur.Name
                        
                        ' ####
                        ' Establish the group(s), creating if necessary
                        ' ####
                        
                        If bClassModule = True Then   ' Use declCurrent.IsClass, to pass declCurrent param
                            ' For classes, ignore any tagged group and use the class name
                            declCurrent.FunctionGroup = vbcComponentCur.Name
                        Else
                            ' Remove spaces after semicolon delimiting tagged groups
                            declCurrent.FunctionGroup = Replace(declCurrent.FunctionGroup, "; ", ";")
                            If declCurrent.FunctionGroup = "" Then
                                declCurrent.FunctionGroup = "Miscellaneous"
                            End If
                        End If
                        
                        If declCurrent.FunctionGroup <> "" Then
                            arrFacts = Split(declCurrent.FunctionGroup, ";")
                            For Each vntGroup In arrFacts
                                bFound = False
                                For Each vntCurrent In colGroups  ' #### need colGroups ####
                                    If vntCurrent.Code = vntGroup Then
                                        bFound = True
                                        
                                        Set grpCurrent = vntCurrent
                                        Exit For
                                    End If
                                Next vntCurrent
            
                                ' Add the group to the collection if new
                                If bFound = False Then
                                    ' Define the new group name as a new collection
                                    Set grpCurrent = New ClsGroup ' #####
                                    grpCurrent.Code = vntGroup
                                    If vbcComponentCur.Type = vbext_ct_StdModule Then
                                        grpCurrent.ModuleType = "Standard"
                                    ElseIf vbcComponentCur.Type = vbext_ct_ClassModule Then
                                        grpCurrent.ModuleType = "Class"
                                    End If
            
                                    colGroups.Add Item:=grpCurrent, Key:=grpCurrent.Name
                                End If
            
                                ' Add the procedure to the group collection
                                grpCurrent.Members.Add Item:=declCurrent
                            Next vntGroup
                        End If  ' declCurrent.FunctionGroup
                    Else
                        ' Procedure already declared; this must be the other side of a class property (Read/Write)
                        nLineDeclare = cmModuleCur.ProcBodyLine(szProcName, nProcKind)
                        szOutput = DeclarationCapture(cmModuleCur, nLineDeclare, nLineCnt)
                        Call declCurrent.DeclarationScope(szOutput)
                    End If  ' ItemExists() (new declaration)
                    
                        ' #### ????
                    
                    ' Skips remaining lines in function
                    ' Remove this; look at every line
                    ' Replace with flag for start of next proc: ProcStartLine + ProcCountLines
                    ' Capture comments up to first Dim or other statement
    '                    nLineStart = nLineStart + cmModuleCur.ProcCountLines _
    '                      (cmModuleCur.ProcOfLine(nLineStart, vbext_pk_Proc), vbext_pk_Proc)
                    nLineStart = nLineStart + cmModuleCur.ProcCountLines _
                      (cmModuleCur.ProcOfLine(nLineStart, vbext_pk_Proc), nProcKind)
                    nRowCurOutput = nRowCurOutput + 1
                    
                Loop  ' Loop until end of module
                
                ' Process the group/module declaration statements in arrDeclarationSection
                ' Parse the statements in a separate function
                'Call DescriptionParse(arrDeclarationSection)
                
                
                ' Add description to the group
                ' colGroups(grpCurrent.Name)
                szGroupName = ""
                szGroupDescription = ""
                If ItemExists(colGroups, grpCurrent.Name) Then
                    If cmModuleCur.CountOfDeclarationLines > 0 Then   ' Avoid error w/ modules that have no declaration setion
                        szGroupDescription = DescriptionParse(arrDeclarationSection)
                        If grpCurrent.ModuleType = "Class" Then
                            'grpCurrent.DescriptionParse (arrDeclarationSection)
                            grpCurrent.Description = szGroupDescription
                        Else
                            nPosition = InStr(szGroupDescription, "|")
                            If nPosition > 0 Then
                                szGroupName = Left(szGroupDescription, nPosition - 1)
                                
                                For Each vntCurrent In colGroups  ' #### need colGroups ####
                                    If vntCurrent.Code = szGroupName Then
                                        'colGroups(szGroupName).Description = Mid(szGroupDescription, nPosition + 1)
                                        vntCurrent.Description = Trim(Mid(szGroupDescription, nPosition + 1))
                                        Exit For
                                    End If
                                Next vntCurrent
                                
                                'colGroups(szGroupName).Description = Mid(szGroupDescription, nPosition + 1)
                            End If
                        End If
                    End If
                End If
                
                
            End If ' Standard/class code module
        End If  ' Skip module name ending in "_"
    Next vbcComponentCur  ' Next module
    
    
    ' ####
    ' #### RESOLVE LINKS ####
    ' ####
    
    ' Loop through all declarations, or flag them as they are created?
    'LinkInterpret()  LinksResolve()
    ' collection: declCurrent.props(i)?
    ' define collection to include these props
    For Each declCurrent In colDeclarations
        vntRc = declCurrent.LinksResolve(colGroups)
    Next declCurrent
    
    ' ####
    ' #### OUTPUT ####
    ' ####
    
    ' Output path for all files #### Need to test existence, create if needed ####
    Set wkbkThis = ThisWorkbook
    szPath = wkbkThis.Names("pathProject").RefersToRange.Cells(1, 1).Text
    
    Call SubfoldersCreate(szPath)
    
    ' HTML Table of Contents file in main output folder
    szFile = szPath & wkbkThis.Names("fileContents").RefersToRange.Cells(1, 1).Text & _
            wkbkThis.Names("txtHtmlExtension").RefersToRange.Cells(1, 1).Text
    szOutput = wkbkThis.Names("ProjectDisplayName").RefersToRange.Cells(1, 1).Text
    vntRc = HtmlContentFileCreate(colGroups, szFile, szOutput)
    
    ' Help File Table of Contents file in main output folder
    szFile = szPath & wkbkThis.Names("fileContents").RefersToRange.Cells(1, 1).Text & ".hhc"
    vntRc = HelpContentFileCreate(colGroups, szFile)
    szFile = szPath & wkbkThis.Names("ProjectName").RefersToRange.Cells(1, 1).Text & ".hhp"
    vntRc = HelpProjectFileCreate(colGroups, szFile)
    
    ' Determine whether to include non-public declarations (Private, Friend or Project)
    bIncludePrivate = wkbkThis.Names("optionIncludePrivate").RefersToRange.Cells(1, 1).Text
    
    ' Write output files to Projects\Topics\ subfolder: colDeclarations
    szFile = szPath & wkbkThis.Names("pathTopicFiles").RefersToRange.Cells(1, 1).Text
    For Each declCurrent In colDeclarations
        If declCurrent.Scope = "Public" Or bIncludePrivate = True _
            Or declCurrent.ScopeRead = "Public" Then
            vntRc = HtmlProcFileCreate(colGroups, declCurrent, szFile)  ' szFile is path only; filename will be appended
        End If
    Next declCurrent
    
    ' Write module to register functions: public, worksheets
    szFile = szPath & wkbkThis.Names("pathTopicFiles").RefersToRange.Cells(1, 1).Text
'    For Each declCurrent In colDeclarations
'        If declCurrent.Scope = "Public" And declCurrent.Worksheet = "Yes" _
'                And declCurrent.Procedure = "Function" Then
'            vntRc = UdfRegister(declCurrent)  ' szFile is path only; filename will be appended
'        End If
'    Next declCurrent
    
    szFile = szPath & wkbkThis.Names("pathGroupFiles").RefersToRange.Cells(1, 1).Text
    vntRc = HtmlGroupFilesCreate(colGroups, szFile)
    
End Sub  ' DocuVeebMain


Function UdfRegister(declCurrent As ClsDeclaration)

    Dim arParamDesc() As String, colParameters As Collection, vntCurrent As Variant
    Dim nCurrent As Integer, nLower As Integer, nUpper As Integer, nCountParam As Integer
    Dim szStatement As String

    ' Create array of parameter descriptions
    nCountParam = colParameters.Count
    ReDim arParamDesc(0 To nCountParam - 1)
    
    For Each vntCurrent In colParameters
        ' Read this declaration's info
        nCurrent = vntCurrent.Item
        arParamDesc(nCurrent) = vntCurrent.Description
        
        ' Write this declaration's info to an input file for the registration function
        
    Next vntCurrent
    
    ' Register the function
    ' Create a statement to be written to a module, to be executed on the user's PC
    Application.MacroOptions Macro:=declCurrent.Name, _
        Description:=declCurrent.BriefDescription, _
        Category:=declCurrent.FunctionGroup, _
        ArgumentDescriptions:=arParamDesc

    szStatement = "Application.MacroOptions Macro:=" & declCurrent.Name & ", _" & vbNewLine
    szStatement = szStatement & "Description:=" & declCurrent.BriefDescription & ", _" & vbNewLine
    szStatement = szStatement & "Category:=" & declCurrent.FunctionGroup & ", _" & vbNewLine
'    szStatement = szStatement & "ArgumentDescriptions:=" & arParamDesc  ' #### ??? ####
    szStatement = szStatement & vbNewLine & vbNewLine

    UdfRegister = szStatement

End Function  ' UdfRegister()



''
' Converts a tagged link to a valid HTML hyperlink.
'
' The link text should be of the form "project.module#procedure display text",
' or the class equivalent "project.class#member display text".
' Display text is optional; it should follow the link text, and be separated from it by a space.
' Spaces within the display text are accepted.
' <P>In the input string, the "#procedure" is required; the "project" is optional.
' The ".class" is required for classes, but ".module" is optional for non-class modules.
' The optional "project" and "project.module" are assumed to point to the current
' project if missing. That is, if the input is "#ProcName", the link will point to
' the ProcName() procedure in the current project. If multiple procedures in a project share the same ProcName,
' the .module is required.
' <P>The function is designed to convert links in tags such as @see and @deprecated,
' and in inline tags such as {@link}.
' <P>Note that links may also be included directly in the comments. For those links, the HTML code
' must be spelled out, from the "<a ...>" to the "</a>" HTML tags.
' @param The text to be converted, stripped of the tag and curly braces.
' @param The list of groups, to determine whether a target is a class or standard module.
' @param The declaration currently being handled; used for project and module default values, if allowed.
' @return Returns the HTML code for a hyperlink.
' Returns xlErrNA if the target procedure cannot be found.
' @group Tags
' @author Dr. Richard Cook
Function LinkInterpret(ByVal szInput As String, ByRef colGroups As Collection, _
        ByRef declCurrent As ClsDeclaration) As Variant

    Dim szLink As String, szDisplay As String, nPositionHash As Integer, nPositionDot As Integer, grpCurrent As ClsGroup
    Dim szProject As String, szModule As String, szProcedure As String, vntOutput As Variant, bLinkIsModule As Boolean
    Dim szExtension As String
    
    ' Separate the input into link and optional display text, delimited by the first space
    nPositionHash = InStr(szInput, " ")
    If nPositionHash > 0 Then
        szLink = Left(szInput, nPositionHash - 1)
        szDisplay = Mid(szInput, nPositionHash + 1)
    Else
        ' No explicit display text; entire input is link
        szLink = szInput
    End If
    
    ' Parse the link text according to specs for the module/class: project.module#procedure
    nPositionHash = InStr(szLink, "#")
    nPositionDot = InStr(szLink, ".")
    If nPositionDot = 0 Then
        ' procedure
        ' #procedure
        szProcedure = Mid(szLink, nPositionHash + 1)
        vntOutput = ProcGetModule(colGroups, szProcedure)  ' #### Must be variant vntRc ####
        If Not IsError(vntOutput) Then
            szModule = ProcGetModule(colGroups, szProcedure)  ' #### Follow up if vntOutput is an error ####
        End If
        szProject = ""
    Else
        ' nPositionDot > 0
        If nPositionHash > 0 Then
            ' .module#procedure
            ' project.module#procedure
            szProcedure = Mid(szLink, nPositionHash + 1)
            szModule = Mid(szLink, nPositionDot + 1, nPositionHash - 1)
            szProject = Left(szLink, nPositionDot - 1)
        Else
            ' nPositionHash = 0
            ' .module
            ' project.module
            szProcedure = ""
            szModule = Mid(szLink, nPositionDot + 1)
            szProject = Left(szLink, nPositionDot - 1)
        End If
    End If
    
    
    ' No explicit display text; display the link
    If szDisplay = "" Then
        If szProcedure <> "" Then
            szDisplay = szProcedure
            bLinkIsModule = False
        Else
            szDisplay = szModule
            bLinkIsModule = True
        End If
    End If
    
    ' ModuleType determines whether class.proc.htm or proc.module.htm
    ' Default form: Proc.Module.htm; reverse it only for classes
    szExtension = ThisWorkbook.Names("txtHtmlExtension").RefersToRange.Cells(1, 1).Text
    
    If szProcedure <> "" Then
        ' Avoid adding leading dot to module filename (if no procedure listed)
        szLink = szProcedure & "."
    Else
        szLink = ""
    End If
    szLink = szLink & szModule & szExtension
    
    For Each grpCurrent In colGroups
        If grpCurrent.Name = szModule And grpCurrent.ModuleType = "Class" Then
            szLink = szModule & "." & szProcedure & szExtension
        End If
    Next grpCurrent
    
    If bLinkIsModule = True Then
        ' For class w/o proc, eliminate double dot
        szLink = Replace(szLink, "..", ".")
        szLink = "../Groups/" & szLink
    End If
    
    ' Skip null input string  ' szInput = "" Or
    If szProcedure = "" And szProject = "" And szModule = "" Then
        vntOutput = CVErr(xlErrNA)
    Else
        ' Build output hyperlink  #### call HtmlNameLink()? ####
        vntOutput = "<a href=""" & szLink & """>" & szDisplay & "</a>" ' #### Complete link ####
    End If
        
    LinkInterpret = vntOutput
    '
End Function  ' LinkInterpret()


Function LinkInterpretOld(ByVal szInput As String, ByRef colGroups As Collection, _
        ByRef declCurrent As ClsDeclaration) As Variant

    Dim szLink As String, szDisplay As String, nPositionHash As Integer, nPositionDot As Integer, grpCurrent As ClsGroup
    Dim szProject As String, szModule As String, szProcedure As String, vntOutput As Variant
    
    ' Separate the input into link and optional display text, delimited by the first space
    nPositionHash = InStr(szInput, " ")
    If nPositionHash > 0 Then
        szLink = Left(szInput, nPositionHash - 1)
        szDisplay = Mid(szInput, nPositionHash + 1)
    Else
        ' No explicit display text; entire input is link
        szLink = szInput
    End If
    
    ' Parse the link text according to specs for the module/class
    nPositionHash = InStr(szLink, "#")
    nPositionDot = InStr(szLink, ".")
    If nPositionHash > 0 Then
        szProcedure = Mid(szLink, nPositionHash + 1)
        szModule = Left(szLink, nPositionHash - 1)
        
        ' Text to left of "#" is either module or project.module
        If szModule <> "" Then
            nPositionDot = InStr(szModule, ".")
            If nPositionDot > 0 Then
                ' Dot means both project and module are explicit
                szProject = Left(szModule, nPositionDot - 1)
                szModule = Mid(szModule, nPositionDot + 1)
            Else
                ' No dot means the project is stated, module is implied
                szProject = szModule
                'szProject = declCurrent.Project
                szModule = ProcGetModule(colGroups, szProcedure)
            End If
        Else
            ' No text before the # means the project and module are both implicit
            szProject = declCurrent.Project
            szModule = ProcGetModule(colGroups, szProcedure)
            'szModule = declCurrent.Module
        End If
    
    ElseIf szLink <> "" Then
        ' No project or project.module; entire link is procedure
        szProcedure = szLink
        szModule = ProcGetModule(colGroups, szProcedure)
        'szProject = declCurrent.Project
    End If
    
    ' No explicit display text; display the link
    If szDisplay = "" Then
        szDisplay = szProcedure
    End If
    
    ' ModuleType determines whether class.proc.htm or proc.module.htm
    ' Default form: Proc.Module.htm; reverse it only for classes
    szLink = szProcedure & "." & szModule & ".htm"
    For Each grpCurrent In colGroups
        If grpCurrent.Name = szModule And grpCurrent.ModuleType = "Class" Then
            szLink = szModule & "." & szProcedure & ".htm"
        End If
    Next grpCurrent
    
    ' Skip null input string  ' szInput = "" Or
    If szProcedure = "" And szProject = "" And szModule = "" Then
        vntOutput = CVErr(xlErrNA)
    Else
        ' Build output hyperlink  #### call HtmlNameLink()? ####
        vntOutput = "<a href=""" & szLink & """>" & szDisplay & "</a>" ' #### Complete link ####
    End If
        
    LinkInterpretOld = vntOutput
    '
End Function  ' LinkInterpret()



''
' Retrieves the module where a procedure is defined, when only the procedure name is known.
'
' Searches a set of Declaration objects for one that matches the input procedure name.
' <P>If a project contains multiple functions with the same name, the method stops looking at the first match,
' which might not be the correct module.
' (That is one more reason to avoid creating multiple functions with the same name within a project.)
' @param The set of groups and/or classes to be searched.
' @param The (unqualified) name of the procedure (the Declaration object's ProcName property).
' @return Returns the module name if the procedure is found.<br>
' Returns xlErrNA if the module cannot be found.
' @group Classes
' @author Dr. Richard Cook
' @date 11/17/15
Function ProcGetModule(ByRef colGroups As Collection, ByVal szProcName As String)

    Dim declCurrent As ClsDeclaration, bFound As Boolean, grpCurrent As ClsGroup, vntRc As Variant
    
    vntRc = CVErr(xlErrNA)
    bFound = False
    
    For Each grpCurrent In colGroups
        For Each declCurrent In grpCurrent.Members
            If szProcName = declCurrent.ProcName Then
                vntRc = declCurrent.Module
                bFound = True
                Exit For
            End If
        Next declCurrent
        If bFound = True Then
            Exit For
        End If
    Next grpCurrent
    
    ProcGetModule = vntRc
    
End Function  ' ProcGetModule()



''
' Finds in-line tags, and has them interpreted.
'
' Currently handles {@link} tags only.
' See the documentation for {@link LinkInterpret} for the proper format of @link tags.
'
' @param The text to be searched for in-line tags.
' @param The list of groups, to determine whether a target is a class or standard module.
' @param The declaration currently being handled; used for project and module default values, if allowed.
' @return Returns the input text with in-line tags converted to text for output.
' Returns xlErrNA if the target procedure cannot be found.
' @group Tags
' @author Dr. Richard Cook
Function LinksFindInText(ByVal szInput As String, ByRef colGroups As Collection, _
        ByRef declCurrent As ClsDeclaration) As Variant
        
    Dim szTag As String, vntOutput As String, nPosition As Integer, szLink As String, szRemaining As String
    Dim nLengthIn As Integer, vntRc As Variant, nLengthTag As Integer
    Dim szCodeTag As String, nPositionCode As Integer, nPositionClose As Integer
    
    szTag = "{@link "
    nLengthTag = Len(szTag)
    vntOutput = ""
    szRemaining = szInput
    
    Do
        ' Look for inline tags
        'bLiteral = False
        nPosition = InStr(szRemaining, szTag)
        If nPosition > 0 Then
            ' Tag found; is it in a section to be handled literally?
            ' #### This assumes only 1 pair of <code> tags; loop through search for <code> tags ####
            szCodeTag = "<code>"
            nPositionCode = InStr(szRemaining, szCodeTag)
            If nPositionCode > 0 And nPositionCode < nPosition Then
                ' Act on <code> tag even if it has no matching </code> tag
                szCodeTag = "</code>"
                nPositionClose = InStr(szRemaining, szCodeTag)
                If nPositionClose > nPositionCode Then
                    ' Treat the inline tag literally
                    vntOutput = vntOutput & Left(szRemaining, nPositionClose + Len(szCodeTag))
                    szRemaining = Mid(szRemaining, nPositionClose + Len(szCodeTag) + 1)
                Else
                    ' No closing </code> tag, assume rest of text is <code>
                    ' That makes the output more obviously erroneous, and more likely to be discovered
                    vntOutput = vntOutput & szRemaining
                    szRemaining = ""
                End If
            Else
                ' capture text up to tag
                vntOutput = vntOutput & Left(szRemaining, nPosition - 1)
                ' Capture tag text and interpret it
                szLink = SubstringBetween(szRemaining, szTag, "}")
                nLengthIn = Len(szLink)
                vntRc = LinkInterpret(Trim(szLink), colGroups, declCurrent)
                If IsError(vntRc) Then
                    vntOutput = vntRc
                    Exit Do
                End If
                szLink = vntRc
                ' Build output from input and interpreted tag
                vntOutput = vntOutput & szLink
                szRemaining = Mid(szRemaining, nPosition + nLengthIn + nLengthTag + 1) ' Skip closing }
            End If
        Else
            ' No more inline tags
            vntOutput = vntOutput & szRemaining
            Exit Do
        End If
    Loop
    
    LinksFindInText = vntOutput
End Function  ' LinksFindInText


''
' Captures the full text of the declaration statement.
'
' Reads both single- and multi-line declaration statements. With multi-line statements,
' the output is how the statement would appear if written on a single line (trailing
' underscores and line breaks are removed).
' <P>If a comment follows the declaration on the last line, the comment is removed.
' @param The module being read.
' @param The line number of the target declaration. The line number can be determined by
' the VBA CodeModule property ProcBodyLine.
' @param An output variable that will indicate the number of lines read in the declaration.
' The number of lines read is useful when reading a module line by line.
' @return Returns the text of the declaration.<br>
' Returns xlErrNA if the input line is not the first line of a declaration.
' @group VBA_Statements
' @author Dr. Richard Cook
' @date 11/14/15
Function DeclarationCapture(ByVal cmModuleCur As CodeModule, ByVal nLineDeclare As Integer, _
        Optional ByRef nLinesRead As Integer) As Variant
    Dim nLineCnt As Integer, szProcName As String, szLineCur As String, nLineNext As Integer
    Dim nPosition As Integer, vntOutput As Variant, nProcKind As Long
    
    nLinesRead = 0
    ' Ensure the input line is a declaration
    szProcName = cmModuleCur.ProcOfLine(nLineDeclare, nProcKind)  ' vbext_pk_Get, Set, Let
    nPosition = cmModuleCur.ProcBodyLine(szProcName, nProcKind)
    If nPosition <> nLineDeclare Then
        vntOutput = CVErr(xlErrNA)
    Else
        ' PROCEDURE DECLARATION
        ' Capture multi-line declarations, denoted by trailing underscore
        ' Retain internal underscores, drop only the trailing underscore
        nLineNext = nLineDeclare
        vntOutput = ""
        szLineCur = ""
        While Right(Trim(cmModuleCur.Lines(nLineNext + nLinesRead, 1)), 1) = "_"
            szLineCur = Trim(cmModuleCur.Lines(nLineNext + nLinesRead, 1))
            szLineCur = Left(szLineCur, Len(szLineCur) - 1) ' Drop trailing underscore
            vntOutput = vntOutput & szLineCur
            nLinesRead = nLinesRead + 1
        Wend
        ' Capture the last line of the proc declaration, w/o trailing underscore
        vntOutput = vntOutput & Trim(cmModuleCur.Lines(nLineNext + nLinesRead, 1))
        nPosition = InStr(vntOutput, "'")
        ' Drop any trailing comment (can appear only on last line of declaration)
        If nPosition > 0 Then
            vntOutput = Left(vntOutput, nPosition - 1)
        End If
        'vntOutput = szOutput
        nLinesRead = nLinesRead + 1
    End If
        
    DeclarationCapture = vntOutput

End Function  ' DeclarationCapture()


''
' Determines the number of folder/subfolder levels in the path.
'
' <P>The count is dependent on whether the path is absolute or relative.
' With an absolute path, the root folder, e.g., "C:\", has 1 level. Below the root, the level is determined by the number of folders and subfolders in the path.
' For example, "C:\Windows\" has 2 levels, while "C:\Users\Username\Documents\" has 4 levels.
' With relative paths, the number of levels is relative to the first folder in the path, which is level 1.
' For example, "..\Folder\Subfolder" has 3 levels, regardless of how far down the tree Folder\ is located.
' <P>The function does not test whether the path is valid. It is up to the user to determine the path's validity.
' Also, the path need not actually exist.
' @param The path to be analyzed.
' @return Returns the number of levels in the input path. On error, returns xlErrRef (#REF!).
' @group Paths
' @worksheet Yes
' @author Dr. Richard Cook
' @date 9/14/2021
' @version 2.10.4
Function PathLevels(szPath As String) As Variant
    Dim vntRc As Variant, folPath As Scripting.Folder, fsoCurrent As Scripting.FileSystemObject
    Dim nLevels As Integer
    
    On Error GoTo ErrorHandler
    vntRc = CVErr(xlErrRef)
    
    Set fsoCurrent = CreateObject("Scripting.FileSystemObject")
    
    nLevels = 0
    
    ' Count the number of parent folders to determine the level
    While szPath <> ""
        szPath = fsoCurrent.GetParentFolderName(szPath)
        nLevels = nLevels + 1
    Wend

    vntRc = nLevels
    
    Set fsoCurrent = Nothing
    
ErrorHandler:
    PathLevels = vntRc
End Function


Sub TestStub()
    Dim vntRc As Variant, declCurrent As ClsDeclaration, szFile As String, folPath As Scripting.Folder, fsoCurrent As Scripting.FileSystemObject
    Dim nLevels As Integer
    
    Set declCurrent = New ClsDeclaration
    declCurrent.Module = "ClsDeclaration"
    declCurrent.Project = "Parser"
    
'    fsoCurrent.
    szFile = ThisWorkbook.Names("pathTopicFiles").RefersToRange.Cells(1, 1).Text
    szFile = "C:\Windows"
    vntRc = PathLevels(szFile)

    'declCurrent.
'    vntRc = LinkInterpret("Parser.ClsDeclaration#ProcName Link Test", declCurrent)
'    vntRc = PathCreate("x:\test\")
'    Err.Clear
'    vntRc = PathCreate("Richard\NoFolder\NoSub\")
    Err.Clear
    vntRc = PathCreate("c:\Richard\NoFolder\NoSub\")
    'vntRc = SubfoldersCreate("c:\Richard\testingxxx\")
    

End Sub


''
' Creates an entire path at once.
'
' Creates all folders that do not exist in the input path.
' Creates multiple levels in a single command.
' Starts at the first folder that does exist, and creates the remaining folders to complete the path.
' <P>The input path must be fully specified, including the drive, or the function will fail.
' @param The full path, including drive.
' @return Returns 0 on success, or an error code indicating failure.
' @group File_System
' @author Dr. Richard Cook
Function PathCreate(ByVal szPathInput As String) As Variant
    Dim fsoCurrent As FileSystemObject, szPathParent As String, drvInput As Drive
    Dim arrFolders() As String, szFolder As String, nCountFldrs As Integer, szDrive As String
    Dim szPathNew As String, nCurrent As Integer, vntRc As Variant

    Set fsoCurrent = CreateObject("Scripting.FileSystemObject")
    Err.Clear
    On Error GoTo Bailout
    
    ' If the project path alread exists, set RC and quit
    If fsoCurrent.FolderExists(szPathInput) Then
'        Err.Number = 58
'        Err.Description = "Folder already exists."
'        Err.Source = "PathCreate"
        Err.Raise (58)
    Else
        ' Ensure drive is valid
        szDrive = fsoCurrent.GetDriveName(szPathInput)
        'vntRc = fsoCurrent.DriveExists(szDrive)
        If fsoCurrent.DriveExists(szDrive) Then
            ' If drive is valid, build the path
            ' Step up the path to find first valid, existing ancestor
            nCountFldrs = 0
            szPathParent = szPathInput
            
            Do
                szPathParent = fsoCurrent.GetParentFolderName(szPathParent)
                nCountFldrs = nCountFldrs + 1
            Loop Until fsoCurrent.FolderExists(szPathParent) ' #### Assumes valid path; needs escape route if not
            
            ' Create each subfolder from that ancestor
            szPathNew = Mid(szPathInput, Len(szPathParent) + 2)
            arrFolders = Split(szPathNew, "\")
            
            'If fsoCurrent.FolderExists(szPathParent) Then
                szPathNew = szPathParent
                For nCurrent = 0 To nCountFldrs - 1
                    szPathNew = szPathNew & "\" & arrFolders(nCurrent)
                    fsoCurrent.CreateFolder (szPathNew)
                Next nCurrent
            'End If
            
            vntRc = 0
        Else
            ' Set error code
            Err.Number = 76
            Err.Description = "Path not found."
            Err.Source = "PathCreate"
            vntRc = Err
        End If  ' Drive Exists
        
    End If
    
    PathCreate = vntRc
    Exit Function
    
Bailout:
    Set vntRc = Err
    Set PathCreate = vntRc
End Function  ' PathCreate


Function SubfoldersCreate(ByVal szPath As String)

    Dim fsoCurrent As FileSystemObject, szPathParent As String, folNew As Folder, szPathNew As String

    Set fsoCurrent = CreateObject("Scripting.FileSystemObject")
    
    ' If the project path doesn't exist, create it
    If Not fsoCurrent.FolderExists(szPath) Then
        szPathParent = fsoCurrent.GetParentFolderName(szPath)
        If Not fsoCurrent.FolderExists(szPathParent) Then
        Else
            Set folNew = fsoCurrent.CreateFolder(szPath)  ' Error 76 bad drive, no parent; 58 already exists
        End If
    End If
    
        If fsoCurrent.FolderExists(szPath) Then
            szPathNew = szPath & "Include\"
            If Not fsoCurrent.FolderExists(szPathNew) Then fsoCurrent.CreateFolder (szPathNew)
            szPathNew = szPath & "Groups\"
            If Not fsoCurrent.FolderExists(szPathNew) Then fsoCurrent.CreateFolder (szPathNew)
            szPathNew = szPath & "Topics\"
            If Not fsoCurrent.FolderExists(szPathNew) Then fsoCurrent.CreateFolder (szPathNew)
        End If
    
End Function



Function HtmlFileOpen(ByVal szPathOut As String, ByVal szTopic As String)
    Dim fsoCurrent As FileSystemObject, txtOutput As TextStream, vntRc As Variant, vntCurrent As Variant
    Dim szHtmlOpen As String, szHtmlClose As String, prmCurrent As ClsParameter, szTest As String
    Dim bFormatted As Boolean, szExtension As String
    
    szHtmlOpen = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbNewLine & "<HTML><HEAD>"
    szHtmlOpen = szHtmlOpen & "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Format.css"">"
    szHtmlClose = "</BODY></HTML>"
    
    szExtension = ThisWorkbook.Names("txtHtmlExtension").RefersToRange.Cells(1, 1).Text
    Set fsoCurrent = CreateObject("Scripting.FileSystemObject")
    Set txtOutput = fsoCurrent.CreateTextFile(szPathOut & szTopic & szExtension)
    
    ' Write header, open body
    txtOutput.WriteLine (szHtmlOpen)
    txtOutput.WriteLine ("<TITLE>" & szTopic & "</TITLE>")
    txtOutput.WriteLine ("</HEAD>")
    txtOutput.WriteLine ("<BODY>")
    txtOutput.WriteBlankLines (2)
    txtOutput.WriteLine ("<H1>" & szTopic & "</H1>")
    txtOutput.WriteBlankLines (2)
    
    HtmlFileOpen = txtOutput
    
    txtOutput.WriteLine (szHtmlClose)
    txtOutput.Close
    
End Function  ' HtmlFileOpen()



''
' Creates an array of the names of items in a collection.
'
' The array can then be used to handle the items in alphabetic order,
' without the performance hit of sorting the collection itself.
' <p>Note that each item in the collection must have the name property defined, or an error will be generated.
' @param The input collection whose names are desired.
' @param The output array; must be declared dynamically.
' @return Returns the number of items in the collection.
' @group Collections
' @author Dr. Richard Cook
Public Function CollectionNames(ByRef colInput As Collection, ByRef arrNames() As Variant)
    Dim lngLast As Long, lngCurrent As Long, vntRc As Variant
    
    lngLast = colInput.Count
    
    If lngLast > 0 Then
        ReDim arrNames(1 To lngLast, 1 To 2)
        
        For lngCurrent = 1 To lngLast
            arrNames(lngCurrent, 1) = colInput(lngCurrent).Name
            arrNames(lngCurrent, 2) = lngCurrent
        Next lngCurrent
        
        vntRc = Array2DBubbleSort(arrNames)
    End If
    
    CollectionNames = lngLast

End Function  ' CollectionNames()



