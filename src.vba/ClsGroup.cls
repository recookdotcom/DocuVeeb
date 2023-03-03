Option Explicit


''
' @groupnote This class handles the grouping of declarations within a project.
'
' <P>Groups are essential to the code parser, because the collection of groups
' includes all information gleaned from the code and the comments.
' Passing the collection of groups to a subroutine passes all information available
' for processing the set of declarations.
' <P>A group can be a formal set of declarations, such as a class, or it can
' be an arbitrary set of declarations that fit together based on any
' criteria of the coder's choice. Declarations are grouped by the
' use of the @group tag in the declaration's comments. A new group is automatically
' established whenever a new @group tag is parsed.
' <P>Note that class members are automatically grouped as the class,
' and any @group tag is ignored within a class.
' <P>Any declaration that is not in a class module and does not have an explicit @group tag
' is assigned the group "Miscellaneous".
' <P>In the default output, the declarations are presented as belonging to the groups.
' They are listed as a set of related procedures, in such a way as to encourage browsing
' the docunmentation for other members of the group. For example, a set of functions
' for handling strings could be grouped.
' <P>A group description follows the same structure as for a declaration, with a brief description
' followed by a blank comment line and an optional full description.
' To be properly interpreted, group descriptions belong in the "Declarations" section of a module.
' If the module contains no statements that would imply a Declarations section, the section
' can be forced by the use of a nonsense declaration, such as:<br><br>
' <code>Private DeclarationEnd = 0</code><br><br>
' A group description for a class should appear in the class module.
''


Private szName As String  ' Human readable, w/ spaces
Private szCode As String  ' Same, w/ underscores instead of spaces
Private szModuleType As String ' Class vs. Standard (code)
Private colMembers As Collection  ' Declaration objects (ClsDeclaration)
Private szBrief As String  ' Optional Class/Group-level brief (one-line) description
Private szDescription As String  ' Optional Class/Group-level description




Private Sub Class_Initialize()
    Set colMembers = New Collection
End Sub

Private Sub Class_Terminate()
    Set colMembers = Nothing
End Sub


''
' Denotes whether the module is a Standard or Class module.
'
' Other module types are not interpreted in this tool.
' @return "Standard" or "Class" as appropriate.
' @author Dr. Richard Cook
Public Property Get ModuleType() As String
    ModuleType = szModuleType
End Property

Public Property Let ModuleType(szInput As String)
    szModuleType = szInput
End Property


''
' Retrieves a human-readable form of the name of the function group, or of the Class module.
'
' Underscores in the name are replaced with spaces.
' <p>In the case of class modules, the @group tag is ignored.
' @return The name of the group or class.
' @author Dr. Richard Cook
Public Property Get Name() As String
    Name = szName
End Property


''
' Retrieves a description of the function group or the Class.
'
' <P>A module-level description in a standard module applies to the group specified in a @group tag.
' In a class module a module-level description applies to the class.
' In a class module, any @group tag is ignored.
' <P>In the default output, the module-level description appears in the contents file with the
' procedures/members of the group or class.
' @return The description for the group or class.
' @author Dr. Richard Cook
Public Property Get Description() As String
    Description = szDescription
End Property

Friend Property Let Description(szIinput As String)
    szDescription = szIinput
End Property


''
' Retrieves the code name of the function group, or of the Class module.
'
' The code name is machine readable, with underscores in place of spaces.
' @return The code name of the group or class.
' @author Dr. Richard Cook
Public Property Get Code() As String
    Code = szCode
End Property

Public Property Let Code(szInput As String)
    szCode = szInput
    szName = Replace(szInput, "_", " ")
End Property


''
' Retrieves a list of procedures in a defined function group, or of class members in a class.
'
' In a class, the members are the public properties and methods of the class.
' For a function group, members are defined by the user by the @group tag in the documentation.
' @return The list of members of the group or class.
' @author Dr. Richard Cook
Public Property Get Members() As Collection
    Set Members = colMembers
End Property

Public Property Let Members(colInput As Collection)
    Set colMembers = colInput
End Property


