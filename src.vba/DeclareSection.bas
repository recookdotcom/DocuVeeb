Option Private Module  ' test
Option Explicit


''
' Provides a set of declaration-section declarations for parsing.
'
' This comment block is intended to apply at the module level, which is applicable for class modules.
' When a module-level comment block appears in a class module, it is automatically applied to the class.
' <P>This can also be used for function-group documentation, where the group can be
' independent of the module. In that case, the associated group is determined by the @group tag.
' <P>When used for function-group documentation, the @group tag must include a single group,
' and at least one declaration must list the function group for it to appear in the default output.
' If a separate group output file exists in the Include\ folder, that file will override any
' module-level comments for that function group in the default output.
' @group Testing
''


''
' Provides a set of declaration-section variables.
' @group Testing
Public testPublicVar As Integer
Global testGlobalVar As Integer ' (superseded by Public; in std modules only)
Private testPrivateVar As Integer ' (module-level)
Dim testDimVar As Integer ' (same as private, but procedure level by convention)


''
' Provides a set of declaration-section constants.
' @group Testing
Public Const testPublicConst = 1
Global Const testGlobalConst As Long = 10 ' (superseded by Public; in std modules only)
Private Const testPrivateConst = "Testing" ' (module-level)
'Dim Const testDimConst = -1 ' Dim const is not allowed


''
' Tests an implicitly public enum declaration.
'
' Also tests using the @param tag with enums.
' @param First parameter
' @param Second parameter
' @param Third parameter
' @group Testing
Enum testEnumPublic
    TestEnumPublic1 = 1
    TestEnumPublic2 = 2
    TestEnumPublic3 = 3
End Enum


''
' Tests an explicitly private enum declaration.
'
' Also uses inline comments for an enum (comments following each statement).
' @group Testing
Private Enum testEnumPrivate
    TestEnumPrivate1 = 10  ' Inline comment 10
    TestEnumPrivate2 = 20  ' Inline comment 20
    TestEnumPrivate3 = 30  ' Inline comment 30
End Enum


''
' Provides a procedure declaration to end the declaration section.
' @group Testing
Private Function EndDeclaration()
    EndDeclaration = True
End Function
