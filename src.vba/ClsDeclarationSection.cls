Option Explicit


''
' @groupnote Class for parsing and interpreting the declare section (at the start of each module).
'
' Not yet in use.
''


Private szModuleScope As String
Private nOptionBase As Integer
Private szOptionCompare As String
Private bOptionPrivate As Boolean
Private szGroupNote As String



Public Property Get Scope() As String
    Scope = szModuleScope
End Property


Public Property Get Compare() As String
    Compare = szOptionCompare
End Property


Public Property Get IsPrivate() As Boolean
    IsPrivate = bOptionPrivate
End Property


Public Property Get OptionBase() As Integer
    OptionBase = nOptionBase
End Property





