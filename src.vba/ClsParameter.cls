Option Explicit


Private bByVal As Boolean
Private bOptional As Boolean
Private bArray As Boolean
Private szParameterName As String
Private szParamDescription As String
Private szVarType As String
Private szDefaultValue As String
'Private vntDefault As Variant
Private bParamArray As Boolean


''
' Parses each parameter of the procedure into individual components.
'
' The components are stored in the relevant class properties.
' <p>Any information that can be derived from the declaration should not be included in the comments or @ tags.
' @param The complete declaration of a single parameter (everything between the commas).
' @return Output is an object with all values that can be derived from the declaration.
' @author Dr. Richard Cook
Public Function DeclarationParse(ByVal szInput As String) As ClsParameter

    Dim nPointer As Integer, bFound As Boolean, nPosition As Integer, szTarget As String
    
    ' Add leading space to allow match at start of statement
    ' Search for target w/ leading & trailing space to avoid match with part of parm name
    szInput = " " & szInput
    nPointer = 1
    
    ' Optional parameter
    szDefaultValue = ""
    szTarget = " Optional "
    nPosition = InStr(szInput, szTarget)
    If nPosition > 0 Then
        bOptional = True
        nPointer = nPosition + Len(szTarget)
        
        ' Optional param default value
        If bOptional = True Then
            szTarget = " = "
            nPosition = InStr(szInput, szTarget)
            If nPosition > 0 Then
                szDefaultValue = Mid(szInput, nPosition + Len(szTarget))
                szDefaultValue = Trim(szDefaultValue)
            End If
        End If
    Else
        bOptional = False
    End If
    
    ' Passing by ref or by value
    szTarget = " ByVal "
    nPosition = InStr(szInput, szTarget)
    If nPosition > 0 Then
        bByVal = True
        bFound = True
    Else
        bByVal = False
        szTarget = " ByRef "
        nPosition = InStr(szInput, szTarget)
        If nPosition > 0 Then
            bFound = True
        End If
    End If
    
    If bFound = True Then
        nPointer = nPosition + Len(szTarget)
    End If
    
    ' Optional parameter array (last parameter in list)
    szTarget = " ParamArray "
    nPosition = InStr(szInput, szTarget)
    If nPosition > 0 Then
        bParamArray = True
        nPointer = nPosition + Len(szTarget)
    Else
        bParamArray = False
    End If
    
    szParameterName = Trim(Mid(szInput, nPointer))
    nPosition = InStr(szParameterName, " ")
    If nPosition > 0 Then
        ' Cut off parm name at 1st space
        szParameterName = Left(szParameterName, nPosition - 1)
    End If

    ' Parameter variable type
    szTarget = " As "
    nPosition = InStr(szInput, szTarget)
    If nPosition > 0 Then
        szVarType = Trim(Mid(szInput, nPosition + Len(szTarget)))
        nPosition = InStr(szVarType, " ")
        If nPosition > 0 Then
            ' Cut off return type after 1st space (default value might exist)
            szVarType = Left(szVarType, nPosition - 1)
        End If
    Else
        szVarType = "Variant"  ' Default if none declared
    End If
    
End Function  ' DeclarationParse()


''
' Reset all class properties to their default values.
' @date 8/24/15
' @author Dr. Richard Cook
Public Function Clear()
    szParameterName = ""
    szVarType = ""
    bByVal = False
    bOptional = False
    bArray = False
    szDefaultValue = ""
End Function


''
' The name of the parameter.
'
' The parameter name is taken as stated in the declaration.
' @date 8/24/15
' @author Dr. Richard Cook
Public Property Get Name() As String
    Name = szParameterName
End Property

Friend Property Let Name(szInput As String)
    szParameterName = szInput
End Property


''
' The parameter's description, separate from any information in the procedure declaration.
'
' The parameter description is read from the @param tag.
' @date 8/24/15
' @author Dr. Richard Cook
Public Property Get Description() As String
    Description = szParamDescription
End Property

Friend Property Let Description(szInput As String)
    szParamDescription = szInput
End Property


''
' Parameter's variable type.
'
' The parameter type is derived from the declaration; it is "variant" if not declared explicitly.
' It should not be included in the @param tag.
' @date 8/24/15
' @author Dr. Richard Cook
Public Property Get VariableType() As String
    VariableType = szVarType
End Property

Friend Property Let VariableType(szInput As String)
    szVarType = szInput
End Property


''
' Whether the parameter is passed ByVal or ByRef.
'
' The parameter passing is derived from the declaration; it is "ByRef" by default, if not declared explicitly.
' True: ByVal; False: ByRef
' @date 8/24/15
' @author Dr. Richard Cook
Public Property Get PassByValue() As Boolean
    PassByValue = bByVal
End Property

Friend Property Let PassByValue(bInput As Boolean)
    bByVal = bInput
End Property


''
' Whether the parameter is an array.
'
' It is derived from the declaration and should not be stated in the @param tag.
' @date 8/24/15
' @author Dr. Richard Cook
Public Property Get IsArray() As Boolean
    IsArray = bArray
End Property

Friend Property Let IsArray(bInput As Boolean)
    bArray = bInput
'    If InStr(szInput, "()") > 0 Then
'        bArray = True
'    Else
'        bArray = False
'    End If
End Property


''
' Whether the parameter is declared as ParamArray.
'
' It is derived from the declaration and should not be stated in the @param tag.
' <P>A ParamArray is a variable list of parameters, all of which are optional (the parameter
' list may contain no parameters or as many as desired).
' @date 11/5/15
' @author Dr. Richard Cook
Public Property Get IsParamArray() As Boolean
    IsParamArray = bParamArray
End Property

Friend Property Let IsParamArray(bInput As Boolean)
    bParamArray = bInput
End Property


''
' Whether the parameter is optional.
'
' It is derived from the declaration and should not be stated in the @param tag.
' @date 8/24/15
' @author Dr. Richard Cook
Public Property Get OptionalInput() As Boolean
    OptionalInput = bOptional
End Property

Friend Property Let OptionalInput(bInput As Boolean)
    bOptional = bInput
End Property


''
' The default value of an optional parameter, if declared.
'
' It is derived from the declaration and should not be stated in the @param tag.
' The default value property is null if no default value is declared.
' @date 8/24/15
' @author Dr. Richard Cook
Public Property Get DefaultValue() As String
    If bOptional = False Then
        DefaultValue = ""
    Else
        DefaultValue = szDefaultValue
    End If
End Property

Friend Property Let DefaultValue(szInput As String)
    szDefaultValue = szInput
End Property

