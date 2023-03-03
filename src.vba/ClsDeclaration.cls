Option Explicit


Private szName As String  ' Fully qualified: class.name or name.module
Private szProcName As String  ' Actual name of procedure/property
Private szDeclaration As String  ' The full declaration statement
Private szStatementRead As String  ' Full statement of Property Get
Private szStatementWrite As String  ' Full statement of Property Let/Set
Private bMacro As Boolean  ' Macro, Function
Private szProcedure As String  ' Macro, Function, Property, Enum, Constant
Private szBrief As String  ' 1-sentence description
Private szDescription As String  ' Full description
Private szModule As String
Private bModuleIsOptionPrivate As Boolean  ' Whether Option Private Module is set (limits Public scope to Project scope)
Private szProject As String  ' VBA project; not yet in use
Private bIsClass As Boolean ' Class vs. standard module
Private szScope As String ' Private, Public, Friend  (Option Private Module - "Project"?)
Private szScopeRead As String ' Applies only to class member properties
Private szScopeWrite As String ' Applies only to class member properties
Private bStatic As Boolean
Private szReturnType As String
Private szFunctionGroup As String  ' Group of related functions/topics
Private szAuthor As String
Private szDate As String
Private szDocDate As String
Private szReturn As String
Private szExample As String
Private szDeprecated As String
Private szSee As String
Private szWorksheet As String
Private szReadWrite As String
Private lngId As Long

Private arrParameters() As ClsParameter
Private colParameters As Collection




''
' Parses the comments attached to a procedure/property declaration, if the comments are formatted according to the standard.
'
' Includes the comment block, with tags, preceeding the declaration.
' <p>If the comment block is not formatted, this function is not called.
' In that case the comment block is stored as the Description, without interpretation.
' @param An array containing the comment lines from the formatted block.
' @return No return value.
' @author Dr. Richard Cook
Friend Function CommentsParse(ByRef arrFacts() As String)

    Dim nCurrent As Integer, szTag As String, nStart As Integer, nStop As Integer, bNewTag As Boolean
    Dim szLineCur As String, nPosition As Integer, nParamCur As Integer
    Dim bTagSection As Boolean, nItem As Integer
    
    nStart = LBound(arrFacts)
    nStop = UBound(arrFacts)
    bNewTag = False
    bTagSection = False
    
    ' Loop through array of input comment lines
    For nCurrent = nStart To nStop
    
        ' Set the tags for the implicit tag lines
        ' The input lines following these implicit tags will be handled according to these tags
        If arrFacts(nCurrent) = "" Then
            ' Do nothing with blank lines
        ElseIf Left(arrFacts(nCurrent), 2) = "''" Then
            szTag = "@brief"
        ElseIf arrFacts(nCurrent) = "'" Then
            ' First empty comment line indicates start of full description; others may be in any section
            ' Once the explicit tag section has begun, an empty comment line has no meaning
            If bTagSection = False Then
                szTag = "@description"
            End If
        Else
            bNewTag = False
            ' Remove the leading comment mark
            szLineCur = Trim(Mid(arrFacts(nCurrent), 2))
            
            ' Get the tag (implicit for Brief and Full Descriptions)
            ' Tag remains in effect until next tag encountered
            If Left(szLineCur, 1) = "@" Then
                ' We've now entered the explicit tag section
                bTagSection = True
                bNewTag = True
                ' A space delimits the @tag from its contents
                nPosition = InStr(szLineCur, " ")
                If nPosition > 0 Then
                    szTag = Left(szLineCur, nPosition - 1)
                    szLineCur = Mid(szLineCur, nPosition)
                Else
                    szTag = szLineCur
                    szLineCur = ""
                End If
            End If
            szLineCur = Trim(szLineCur)
            
            ' Use switch/case to handle known tags; default case for unknown tags
            ' Set class member according to switch/case
            ' If new tag, start w/ clean entry; otherwise, append to prior
            Select Case szTag
                Case "@author"
                    If szAuthor <> "" Then
                        szAuthor = szAuthor & ";"
                    End If
                    szAuthor = szAuthor & szLineCur
                Case "@date"
                    szDate = szLineCur
                Case "@docdate"
                    szDocDate = szLineCur
                Case "@brief" ' implicit
                    szBrief = szBrief & " " & szLineCur
                Case "@description" ' implicit
                    szDescription = szDescription & " " & szLineCur
                Case "@return"
                    szReturn = szReturn & " " & szLineCur
                Case "@param"
                    If bNewTag = True Then
                        nParamCur = nParamCur + 1
                    End If
                    ' Avoid error if user supplies more parameter tags than parameters
                    ' Parameter object will not exist for extra parameters
                    On Error Resume Next
                        Parameters(nParamCur).Description = Parameters(nParamCur).Description & " " & szLineCur
                    On Error GoTo 0
                Case "@group"
                    If szFunctionGroup <> "" Then
                        szFunctionGroup = szFunctionGroup & ";"
                    End If
                    szFunctionGroup = szFunctionGroup & szLineCur
                Case "@see"
                    If szSee <> "" Then
                        szSee = szSee & ";"
                    End If
                    szSee = szSee & szLineCur
                Case "@example"
'                    If szExample <> "" Then
'                        szExample = szExample & ";"
'                    End If
'                    szExample = szExample & szLineCur
                    szExample = szLineCur
                Case "@deprecated"
                    ' Single line only
                    szDeprecated = szLineCur
                Case "@worksheet"
                    ' Usually Yes/No, but may include other comments. Interpret Yes/No on output, not here.
                    szWorksheet = szLineCur
                Case Else
            End Select
        End If
        
    Next nCurrent
    
    CommentsParse = 0

End Function  ' CommentsParse()



''
' Parses the procedure/property declaration into individual components.
'
' Any information that can be derived from the declaration should not be included in comments.
' @param The complete declaration as a single string, with continuation underscores removed
' from a multi-line declaration.
' @return No return value.
' @author Dr. Richard Cook
Friend Function ProcStatementParse(ByVal szStatement As String)

    Dim szWorkPre As String, szWorkPost As String, szTarget As String, nPosition As Integer, bFound As Boolean
    Dim szParmList As String, bErrorFlag As Boolean, vntRc As Variant, nPointer As Integer, nNext As Integer
    
    ' Scope: Private, Public, Project
    Call DeclarationScope(szStatement)
    
    If szDeclaration <> "" Then
        ' Declaration already captured; this must be the matching declaration of a property (Get vs. Let)
        ' Capture this one as the alternative declaration, do nothing else with it
'        szDeclarationAlt = Trim(szStatement)
    Else
        szDeclaration = Trim(szStatement)
        
        ' Split the declaration into leading, parm list, trailing
        ' Separate the return value from the statement, following closing paren
        nPosition = 1
        nPointer = 0
        While nPosition > 0
            nPosition = InStr(nPosition, szStatement, ")")
            If nPosition > 0 Then
                nPosition = nPosition + 1
                'nCount = nCount + 1
                nPointer = nPosition
            End If
            'nPointer = nPointer + nPosition
        Wend
        
        'nPosition = InStr(szStatement, ")")
        szParmList = Left(szStatement, nPointer - 2)
        szWorkPost = " " & Mid(szStatement, nPointer)
        
        ' Split the remainder (before closing paren) into Pre and ParmList
        nPosition = InStr(szParmList, "(")
        szWorkPre = " " & Left(szParmList, nPosition - 1) ' Leading space for finding whole word at start
        szParmList = Mid(szParmList, nPosition + 1)
        
        ' Return type after closing paren
        szTarget = " As "
        If InStr(szWorkPost, szTarget) > 0 Then
            szReturnType = Trim(Mid(szWorkPost, Len(szTarget) + 1))
        Else
            szReturnType = "Variant"  ' Default if none declared
        End If
        
'        ' Scope: Private, Public, Project
'        Call DeclarationScope(szStatement)
'        szTarget = " Private "
'        If InStr(szWorkPre, szTarget) > 0 Then
'            szScope = Trim(szTarget)
'        Else
'            szTarget = " Friend "
'            If InStr(szWorkPre, szTarget) > 0 Then
'                szScope = Trim(szTarget)
'            Else
'                ' Option Private Module limits Public declarations to the current project
'                If bModuleIsOptionPrivate = True Then
'                    szScope = "Project"
'                Else
'                    szScope = "Public"
'                End If
'            End If
'        End If
        
        ' Static or not
        szTarget = " Static "
        nPosition = InStr(szWorkPre, szTarget)
        If nPosition > 0 Then
            bStatic = True
        Else
            bStatic = False
        End If
        
        ' Type of procedure: Macro vs. Function
        szTarget = " Function "
        nPosition = InStr(szWorkPre, szTarget)
        If nPosition > 0 Then
            bMacro = False
            szProcedure = "Function"
        Else
            szTarget = " Sub "
            nPosition = InStr(szWorkPre, szTarget)
            If nPosition > 0 Then
                bMacro = True
                szProcedure = "Macro"
            End If
    '    Else
    '        szTarget = " Enum "
    '        nPosition = InStr(szWorkPre, szTarget)
    '        If nPosition > 0 Then
    '            bMacro = True
    '            szProcedure = "Enum"
    '        End If
    '    Else
    '        szTarget = " Const "
    '        nPosition = InStr(szWorkPre, szTarget)
    '        If nPosition > 0 Then
    '            bMacro = True
    '            szProcedure = "Constant"
    '        End If
    '    Else
    '        bErrorFlag = True  ' #### Provide info on type of error ####
        End If
        
        ' Proc name ' #### Capture directly with built-in statement ####
        
        ' Parm list
        vntRc = ParmListParse(szParmList)
    End If  ' szDeclaration exists
    
    ProcStatementParse = bErrorFlag
    
End Function  ' ProcStatementParse()



''
' Parses statements in the declaration section into individual components.
'
' Built from ProcStatementParse, to handle the initial declaration section.
' Not yet in use.
' Enums must be captured and presented in full (they do not include line continuation underscores).
' Any information that can be derived from the declaration should not be included in comments.
' @param The complete declaration as a single string, with continuation underscores removed
' from a multi-line declaration.
' @return No return value.
' @author Dr. Richard Cook
Friend Function DeclarationSectionParse(ByVal szStatement As String)

    Dim szWorkPre As String, szWorkPost As String, szTarget As String, nPosition As Integer, bFound As Boolean
    Dim szParmList As String, bErrorFlag As Boolean, vntRc As Variant, nPointer As Integer, nNext As Integer
    
    If szDeclaration <> "" Then
        ' Declaration already captured; this must be the matching declaration of a property (Get vs. Let)
        ' Capture this one as the alternative declaration, do nothing else with it
'        szDeclarationAlt = Trim(szStatement)
    Else
        szDeclaration = Trim(szStatement)
        
'        ' Split the declaration into leading, parm list, trailing
'        ' Separate the return value from the statement, following closing paren
'        nPosition = 1
'        nPointer = 0
'        While nPosition > 0
'            nPosition = InStr(nPosition, szStatement, ")")
'            If nPosition > 0 Then
'                nPosition = nPosition + 1
'
'                nPointer = nPosition
'            End If
'            'nPointer = nPointer + nPosition
'        Wend
'
'        szParmList = Left(szStatement, nPointer - 2)
'        szWorkPost = " " & Mid(szStatement, nPointer)
'
'        ' Split the remainder (before closing paren) into Pre and ParmList
'        nPosition = InStr(szParmList, "(")
'        szWorkPre = " " & Left(szParmList, nPosition - 1) ' Leading space for finding whole word at start
'        szParmList = Mid(szParmList, nPosition + 1)
        
        ' Return type after closing paren: will hold for global constants and variables
        szTarget = " As "
        If InStr(szWorkPost, szTarget) > 0 Then
            szReturnType = Trim(Mid(szWorkPost, Len(szTarget) + 1))
        Else
            szReturnType = "Variant"  ' Default if none declared
        End If
        
        Call DeclarationScope(szStatement)
        
'        ' Private vs. Public
'        szTarget = " Private "
'        If InStr(szWorkPre, szTarget) > 0 Then
'            szScope = Trim(szTarget)
'        Else
'            szTarget = " Friend "
'            If InStr(szWorkPre, szTarget) > 0 Then
'                szScope = Trim(szTarget)
'            Else
'                szScope = "Public"
'            End If
'        End If
        
        ' Static or not
        szTarget = " Static "
        nPosition = InStr(szWorkPre, szTarget)
        If nPosition > 0 Then
            bStatic = True
        Else
            bStatic = False
        End If
        
        ' Type of
        szTarget = " Enum "
        nPosition = InStr(szWorkPre, szTarget)
        If nPosition > 0 Then
            bMacro = False
            szProcedure = "Enum"
        End If

        szTarget = " Const "
        nPosition = InStr(szWorkPre, szTarget)
        If nPosition > 0 Then
            bMacro = True
            szProcedure = "Constant"
        End If

'            szTarget = " Sub "
'            nPosition = InStr(szWorkPre, szTarget)
'            If nPosition > 0 Then
'                bMacro = True
'                szProcedure = "Macro"
'            End If
    '    Else
    '        szTarget = " Enum "
    '        nPosition = InStr(szWorkPre, szTarget)
    '        If nPosition > 0 Then
    '            bMacro = True
    '            szProcedure = "Enum"
    '        End If
    '    Else
    '        szTarget = " Const "
    '        nPosition = InStr(szWorkPre, szTarget)
    '        If nPosition > 0 Then
    '            bMacro = True
    '            szProcedure = "Constant"
    '        End If
    '    Else
    '        bErrorFlag = True  ' #### Provide info on type of error ####
    '    End If
        
        ' Proc name ' #### Capture directly with built-in statement ####
        
        ' Parm list
        vntRc = ParmListParse(szParmList)
    End If  ' szDeclaration exists
    
    DeclarationSectionParse = bErrorFlag
    
End Function  ' DeclarationSectionParse()


''
' Parses the parameter list in a procedure declaration into individual components.
'
' Any information that can be derived from the declaration should not be included in comments.
' @param The parameter list from the declaration (all text within the parentheses).
' @return No return value.
' @author Dr. Richard Cook
Friend Function ParmListParse(ByVal szStatement As String)

    Dim bErrorFlag As Boolean, nPosition As Integer, prmTemp As ClsParameter
    Dim nStart As Integer, nCount As Integer, arrStatement() As String, nCurrent As Integer

    Set colParameters = New Collection
    
    If szStatement = "" Then
        nCount = 0
    Else
        ' List of parameters from initial input: szParmList
        ' Count parameters
        nStart = 0
        nCount = 1
        Do
            ' Parameter count is comma count + 1
            nStart = InStr(nStart + 1, szStatement, ",")
            If nStart = 0 Then Exit Do
            nCount = nCount + 1
        Loop
        
        arrStatement = Split(szStatement, ",")
        ReDim arrParameters(0 To nCount - 1)
        
        For nCurrent = 0 To nCount - 1
            ' Call the Parameters class to parse statement into Parameter objects
'            Set arrParameters(nCurrent) = New ClsParameter
'            arrParameters(nCurrent).DeclarationParse (szStatement)
'            If prmTemp Is Nothing Then
'                Set prmTemp = New ClsParameter
'            End If
            If arrParameters(nCurrent) Is Nothing Then
                Set arrParameters(nCurrent) = New ClsParameter
            End If
            arrParameters(nCurrent).DeclarationParse (arrStatement(nCurrent))
            colParameters.Add Item:=arrParameters(nCurrent), Key:=CStr(nCurrent)
        Next nCurrent
        '
        
        ' Passing (do for each param)
        
    End If

    ParmListParse = bErrorFlag
    
End Function  ' ParmListParse()


''
' The name of either a procedure or a property, as used in the VBA code.
'
' <p>Note that procedure names are unique within a module, but are not necessarily unique within a VBA project.
' The Name property (the fully qualified name for procedure/property) is unique across an entire project,
' and is thus used to identify the Declaration object.
' Proc Name
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get ProcName() As String
    ProcName = szProcName
End Property

Friend Property Let ProcName(szInput As String)
    szProcName = szInput
End Property


''
' The fully qualified name for procedure/property.
'
' Takes the format class.proc within a class module, or proc.module in a standard module.
' Fully qualified names are unique within a VBA project.
' They also allow sorting in a meaningful way.
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get Name() As String
    Name = szName
End Property

Friend Property Let Name(szInput As String)
    szName = szInput
End Property


''
' The brief description gives a broad overview of the procedure or property.
'
' It is typically a single sentence that ignores the details, which are presented in the Description property.
' The brief description is usually the only summary presented with the procedure or property name within a list.
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get BriefDescription() As String
    BriefDescription = szBrief
End Property


''
' The full description of the procedure or property.
'
' The Description should include all relevant detail needed to work with the procedure or property.
' It should not include information provided within the other tags or derived from the declaration itself.
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get Description() As String
    Description = szDescription
End Property

Friend Property Let Description(szInput As String)
    szDescription = szInput
End Property


''
' Function Group to which the procedure belongs.
'
' For class members, the class is set as the group, ignoring any @group tag.
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get FunctionGroup() As String
    FunctionGroup = szFunctionGroup
End Property

Friend Property Let FunctionGroup(szInput As String)
    szFunctionGroup = szInput
End Property


''
' An external HTML file containing 1 or more examples of the procedure.
'
' The tag should be the path to the example file, relative or absolute.
' Recommended practice is to store example files in the Include\ folder.
' @author Dr. Richard Cook
' @date 8/13/16
Public Property Get Example() As String
    Example = szExample
End Property

Friend Property Let Example(szInput As String)
    szExample = szInput
End Property


''
' Specifies whether a procedure is a macro (True) or a function (False).
'
' @deprecated The {@link Procedure} property describes the declaration as a Macro, Function, Property, or Enum.
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get Macro() As Boolean
    Macro = bMacro
End Property




''
' Determines the type of procedure: Macro, Function, Property, or Enum.
'
' Note that Enums are not currently supported in the documenter.
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get Procedure() As String
    Procedure = szProcedure
End Property

'Friend Property Let Procedure(szInput As String)
'    szProcedure = szInput
'End Property


''
' Retrieves a function's return values as described in the @return tag.
'
' The return type is parsed from the declaration, and should not be repeated in the @return tag.
' @date 8/24/15
' @author Dr. Richard Cook
Public Property Get Returns() As String
    Returns = szReturn
End Property


''
' Retrieves a function's return type, parsed from the procedure declaration.
'
' The return type should not be included in the @return tag, unless different types
' are returned under different circumstances.
' @author Dr. Richard Cook
' @date 9/11/15
Public Property Get ReturnType() As String
    ReturnType = szReturnType
End Property


''
' Gets the date the procedure was last updated, specified in the @docdate tag.
'
' The @docdate entry should be updated whenever the documentation section is edited.
' @author Dr. Richard Cook
' @date 4/1/2016
' @docdate 4/1/2016
Public Property Get DocumentedDate() As String
    DocumentedDate = szDocDate
End Property


''
' Gets the date the procedure was created, specified in the @date tag.
'
' @author Dr. Richard Cook
' @date 9/11/15
Public Property Get Created() As String
    Created = szDate
End Property


''
' Retrieves the scope of a procedure: Private, Public, Friend or Project.
'
' "Project" scope appears to be the non-class equivalent of Friend; the procedure is declared
' Public, but can be accessed only within the current project.
' @author Dr. Richard Cook
' @date 9/11/15
Public Property Get Scope() As String
    Scope = szScope
End Property


''
' Retrieves the read scope of a class property: Private, Public, or Friend.
'
' @author Dr. Richard Cook
' @date 9/11/15
Public Property Get ScopeRead() As String
    ScopeRead = szScopeRead
End Property


''
' Retrieves the write scope of a class property: Private, Public, or Friend.
'
' @author Dr. Richard Cook
' @date 9/11/15
Public Property Get ScopeWrite() As String
    ScopeWrite = szScopeWrite
End Property


''
' Retrieves any See Also links (tagged as @see).
'
' @author Dr. Richard Cook
' @date 10/17/15
Public Property Get SeeAlso() As String
    SeeAlso = szSee
End Property


''
' Determines whether a function can be used in a worksheet cell.
'
' The value is read from the @worksheet tag.
' @param "Yes" or "No", or a statement to be output to the documentation.
' In the default output, "Yes" or "No" by itself results in a generic statement. Any other statement
' will be output verbatim, with no interpretation.
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get Worksheet() As String
    Worksheet = szWorksheet
End Property


''
' Determines whether the procedure is declared as static.
'
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get IsStatic() As Boolean
    IsStatic = bStatic
End Property


''
' Whether the procedure/property is deprecated.
'
' The procedure/property is noted as deprecated if the deprecated member is non-empty
' (the @deprecated tag is in place AND has a comment attached to it).
' <p>The @deprecated tag should refer to the preferred code to use, preferably with an in-line {@link} tag. For example:<br><br>
' <code>' @deprecated As of Release 2.0.0, replaced by {@link OpWorkbookReleaseDate}.</code><br>
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get Deprecated() As String
    Deprecated = szDeprecated
End Property


''
' Creator of the procedure.
'
' May include multiple authors, in which case the primary author should be listed first.
' Multiple entries may be listed on one line, separated with semicolons, or each on its own line.
' For example, either of these forms is acceptable:<br><br>
' <code>' @author Dr. Richard Cook; Noah Count</code><br><br>
' Or:<br><br>
' <code>
' ' @author Dr. Richard Cook<br>
' ' Noah Count<br>
' </code>
' @return The list of author(s) is stored as a string, with semicolons between authors.
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get Author() As String
    Author = szAuthor
End Property


''
' Determines whether a property can be read and written.
' @return Returns "Read only", "Write only", or "Read/write", or "" if it cannot be determined.
' @author Dr. Richard Cook
' @date 10/4/15
Public Property Get ReadWrite() As String
    ReadWrite = szReadWrite
End Property

Friend Property Let ReadWrite(szInput As String)
    szReadWrite = szInput
End Property


''
' Module in which the procedure is defined
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get Module() As String
    Module = szModule
End Property

Friend Property Let Module(szInput As String)
    szModule = szInput
End Property


''
' Project in which the procedure's module is defined.
'
' Allows looping through multiple projects without name collisions.
' @author Dr. Richard Cook
' @date 10/30/15
Public Property Get Project() As String
    Project = szProject
End Property

Friend Property Let Project(szInput As String)
    szProject = szInput
End Property


''
' Whether the procedure is a class member.
'
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get IsClass() As Boolean
    IsClass = bIsClass
End Property

Friend Property Let IsClass(bInput As Boolean)
    bIsClass = bInput
End Property


''
' Whether the procedure is in a module with Option Private Module.
'
' @author Dr. Richard Cook
' @date 12/28/15
Public Property Get IsOptionPrivate() As Boolean
    IsOptionPrivate = bModuleIsOptionPrivate
End Property

Friend Property Let IsOptionPrivate(bInput As Boolean)
    bModuleIsOptionPrivate = bInput
End Property


''
' Captures the complete declaration as stated in the module.
'
' Note that multi-line statements are reduced to a single line, with trailing underscores removed.
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get Declaration() As String
    Declaration = szDeclaration
End Property


'''
'' Complete declaration statement for the alternative declaration of a property (Get/Let).
''
'' @author Dr. Richard Cook
'' @date 10/13/15
'Public Property Get DeclarationAlt() As String
'    DeclarationAlt = szDeclarationAlt
'End Property
'
''Friend Property Let Declaration(szInput As String)
''    szDeclaration = szInput
''End Property


''
' Complete declaration statement for the read declaration of a property (Get).
'
' @author Dr. Richard Cook
' @date 10/17/15
Public Property Get StatementRead() As String
    StatementRead = szStatementRead
End Property


''
' Complete declaration statement for the write declaration of a property (Let/Set).
'
' @author Dr. Richard Cook
' @date 10/17/15
Public Property Get StatementWrite() As String
    StatementWrite = szStatementWrite
End Property


''
' Retrieves a list of the parameters declared for the procedure.
'
' @author Dr. Richard Cook
' @date 8/24/15
Public Property Get Parameters() As Collection
    Set Parameters = colParameters
End Property


''
' Resets all class properties to their default values.
'
' @author Dr. Richard Cook
' @date 8/24/15
Public Function Clear()
    szName = ""
    szDescription = ""
    szDeclaration = ""
    szScope = ""
    szModule = ""
    szFunctionGroup = ""
    bStatic = False
    bMacro = False
    Erase arrParameters
    Set colParameters = Nothing
End Function




''
' Fills several class members that track the read/write status of a property.
'
' Captures the declaration statement(s) for a property, setting them as StatementRead and/or StatementWrite.
' Also captures a string describing the status, such as "Read/Write", and sets the string as the property ReadWrite.
' <P>The status is determined by whether the property has a Get() and/or either a Let() or Set() statement.
' @param The name of the class module.
' @param The name of the property being tested.
' @return Returns one of the following strings, as appropriate:<br>
' "Read Only"<br>
' "Write Only"<br>
' "Read/Write"<br>
' @author Dr. Richard Cook
' @date 10/16/2015
Friend Function PropertyFillReadWrite(cmModuleCur As CodeModule) As String   ' vbext_pk_Get, Set, Let
'Function PropertyFillReadWrite(cmModuleCur As CodeModule, szProcName As String) As String   ' vbext_pk_Get, Set, Let
    ' #### Change/eliminate the parameters to fit the class ####
    Dim nLineDeclare As Long, nOutput As Integer, szOutput As String, szProcName As String
    nOutput = 0
    szOutput = ""
    
    szProcName = Me.ProcName
    
    ' Accessing a non-existent property will generate an error.
    ' Attempt to access each type and see whether it generates an error.
    On Error Resume Next
    Err.Clear
    nLineDeclare = cmModuleCur.ProcBodyLine(szProcName, vbext_pk_Get)
    If Err.Number = 0 Then
        szStatementRead = DeclarationCapture(cmModuleCur, nLineDeclare)
        nOutput = nOutput + 1
    End If
    
    Err.Clear
    nLineDeclare = cmModuleCur.ProcBodyLine(szProcName, vbext_pk_Set)
    If Err.Number = 0 Then
        szStatementWrite = DeclarationCapture(cmModuleCur, nLineDeclare)
        nOutput = nOutput + 2
    Else
        Err.Clear
        nLineDeclare = cmModuleCur.ProcBodyLine(szProcName, vbext_pk_Let)
        If Err.Number = 0 Then
            szStatementWrite = DeclarationCapture(cmModuleCur, nLineDeclare)
            nOutput = nOutput + 2
        End If
    End If
    On Error GoTo 0
    
    ' If at least one read/write declaration, we know we have a property, not a procedure
    If nOutput > 0 Then
        szProcedure = "Property"
    End If
    
    Select Case nOutput
        Case 1:
            szOutput = "Read Only"
        Case 2:
            szOutput = "Write Only"
        Case 3:
            szOutput = "Read/Write"
        Case Else:
            szOutput = ""
    End Select
    
    szReadWrite = szOutput
    PropertyFillReadWrite = szOutput

End Function  ' PropertyFillReadWrite()


''
' Determines the scope of the given declaration statement.
'
' Sets the appropriate scope (read/write/all) for this declaration.
' <P>The module type and the option statements must already be set for this declaration before calling this procedure.
' @param The complete declaration statement, including all lines for multi-line declarations.
' Allows parsing different scope levels for read/write properties.
' @return Returns the scope: "Public", "Private", "Friend", or "Project".
' @author Dr. Richard Cook
' @date 1/17/2016
Friend Function DeclarationScope(ByVal szStatement As String)
    Dim szTarget As String, bClass As Boolean, bModuleIsOptionPrivate As Boolean, szScopeTemp As String
    Dim szGetSet As String, bFound As Boolean
    
    bClass = Me.IsClass
    bModuleIsOptionPrivate = Me.IsOptionPrivate
    
    ' Add leading space to allow simple search for complete word
    szStatement = " " & szStatement
    
    ' First look for explicit scope statement
    szTarget = " Private "
    If InStr(szStatement, szTarget) > 0 Then
        szScopeTemp = Trim(szTarget)
    Else
        szTarget = " Friend "
        If InStr(szStatement, szTarget) > 0 Then
            szScopeTemp = Trim(szTarget)
        Else
            szTarget = " Public "
            If InStr(szStatement, szTarget) > 0 Then
                ' Explicit "Public" statement
                If bModuleIsOptionPrivate = True Then
                    ' Option Private Module limits Public declarations to the current project
                    ' Not allowed in class module; this must be standard module
                    szScopeTemp = "Project"
                Else
                    szScopeTemp = "Public"
                End If
            Else
                ' Default, nothing declared; scope depends on module type
                If bClass = True Then
                    szScopeTemp = "Private"
                Else
                    szScopeTemp = "Public"
                End If
            End If
        End If
    End If
    
    ' In a class, set the read/write scope for properties
    If IsClass = True Then
        szTarget = " Get "
        If InStr(szStatement, szTarget) > 0 Then
            szScopeRead = szScopeTemp
            bFound = True
        End If
        szTarget = " Let "
        If InStr(szStatement, szTarget) > 0 Then
            szScopeWrite = szScopeTemp
            bFound = True
        End If
        szTarget = " Set "
        If InStr(szStatement, szTarget) > 0 Then
            szScopeWrite = szScopeTemp
            bFound = True
        End If
    End If
    
    ' If no Get, Let or Set found, declaration is not a property, or module is not a class
    If bFound = False Then
        szScope = szScopeTemp
    End If
    
    DeclarationScope = szScopeTemp

End Function  ' DeclarationScope()


''
' Finds and resolves inline tags in each ClsDeclaration class member that allows them.
'
' The function cannot be called until all declarations in all modules have been parsed.
' Currently @link tags are the only inline tags supported.
' <P>See {@link LinkInterpret} for a discussion of the links and how they are interpreted.
' @param The collection of all groups of declarations.
' @return The current form returns 0 in all cases.
' @see {@link .Tags}
' @author Dr. Richard Cook
' @date 3/22/16
Friend Function LinksResolve(ByRef colGroups As Collection) ', ByRef declCurrent As ClsDeclaration) As Variant
    Dim vntCurrent As Variant
    
    szBrief = LinksFindInText(szBrief, colGroups, Me)
    szDescription = LinksFindInText(szDescription, colGroups, Me)
    szReturn = LinksFindInText(szReturn, colGroups, Me)
    szExample = LinksFindInText(szExample, colGroups, Me)
    szDeprecated = LinksFindInText(szDeprecated, colGroups, Me)
    szWorksheet = LinksFindInText(szWorksheet, colGroups, Me)
    
    ' Handle the See Also statement depending on its contents
    If szSee <> "" Then
        If Left(szSee, 1) = """" And Right(szSee, 1) = """" Then
            ' If in double quotes, strip the quotes and do not interpret the text
            szSee = Mid(szSee, 2, Len(szSee) - 2)
        Else
            ' Not quoted; convert links
            szSee = LinksFindInText(szSee, colGroups, Me)
        End If
    End If
    
    For Each vntCurrent In colParameters
        vntCurrent.Description = LinksFindInText(vntCurrent.Description, colGroups, Me)
    Next vntCurrent

    LinksResolve = 0
    
End Function

