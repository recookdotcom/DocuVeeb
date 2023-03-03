


''
' Tests documentation of Enums.
'
' Has no value outside of test.
' @param Single
' @param Paired
' @param Threesome
' @group Testing
' @author Dr. Richard E. Cook
' @date 9/14/15
Public Enum Status
    pmoSingle = 1
    pmoPaired = 2
    pmoThreesome = 3
End Enum


''
' Exports all modules in the active workbook to text files.
'
' Output files can be compared across versions to identify all changes in a new release.
' @group DocuVeeb
' @author Dr. Richard Cook
Sub ButtonCodeExport()
    Dim szPathOut As String
    
    'szPathOut = "C:\Richard\Consulting\State\Excel\Libraries\Compare\1_1_7\"
    
    If szPathOut = "" Then
        szPathOut = ThisWorkbook.Names("pathExportOutput").RefersToRange.Text
    End If
    
    If szPathOut <> "" Then
        CodeModuleExport szPathOut
    Else
        Call MsgBox("Enter an output path on the Exporter page", vbOKOnly, "No Path Specified")
    End If

End Sub


''
' Not currently called by any code w/i this project
Sub CodeExportAllModules()
    Dim szPathOut As String, vntRc As Variant  ' , wkbkThis As Workbook
    
    vntRc = ThisWorkbook.Names("pathExportOutput").RefersToRange.Text
    If VarType(vntRc) = vbString Then
        szPathOut = vntRc
        CodeModuleExport szPathOut
    End If
    
End Sub


' List all procs in a single module as a macro
' Call repeatedly for an entire workbook
' Parse the actual declaration, for sub vs. function, and the list of parameters

' Extend to parse each line, listing all function/sub calls
' Can build the entire program flow from that list
' Need list of user-defined functions (might be different library) since Call keyword not used consistently
' Function calls are also non-standard


Sub CodeModuleStub()
    ' Run declaration parser, put list into array
    ' Then run overall flow, look at each code line for any array elements

    Dim vbcComponentCur As VBComponent, wkbkCurrent As Workbook, vbccList As VBComponents
    Dim nLineCount As Long, nLineCurrent As Long, szNameFunction As String
    
    Set wkbkCurrent = ThisWorkbook
    
    Set vbccList = wkbkCurrent.VBProject.VBComponents
    
    For Each vbcComponentCur In vbccList
        If vbcComponentCur.Type = vbext_ct_StdModule Or vbcComponentCur.Type = vbext_ct_ClassModule Then
            nLineCount = vbcComponentCur.CodeModule.CountOfLines
            
            For nLineCurrent = 1 To nLineCount
                szNameFunction = vbcComponentCur.CodeModule.ProcOfLine(nLineCurrent, vbext_pk_Proc)
            Next
            
        End If
    Next
    
End Sub





Sub ProceduresList()
' Based on code from Walkenbach

    Dim vbpProjectCur As VBIDE.VBProject, vbcComponentCur As VBComponent, cmModuleCur As CodeModule
    Dim nStartLine As Long, szOutput As String, szProcName As String, nLineNext As Long, nLineCnt As Integer
    Dim nColumnCur As Long, nRowCur As Long, wkshOutput As Worksheet, arrFacts() As String
    Dim szLineCur As String
    
   ' Use the active workbook
    Set vbpProjectCur = ActiveWorkbook.VBProject
    
    Set wkshOutput = ActiveWorkbook.Worksheets("CodeList")
    
    nRowCur = 1
   ' Loop through the VB components
    For Each vbcComponentCur In vbpProjectCur.VBComponents
        Set cmModuleCur = vbcComponentCur.CodeModule
'        If vbcComponentCur.Name = "MacrosEval" Then
'            nColumnCur = 1
'        End If
        If vbcComponentCur.Type = vbext_ct_StdModule Then
            'szOutput = szOutput & vbNewLine
            nStartLine = cmModuleCur.CountOfDeclarationLines + 1
            
            Do Until nStartLine >= cmModuleCur.CountOfLines
                ' Determine name of next procedure
                ' nStartLine is the first blank line after the prior declaration
                szOutput = cmModuleCur.ProcOfLine(nStartLine, vbext_pk_Proc)
                
                ' Write the declaration to the CodeList sheet
                wkshOutput.Cells(nRowCur, 1).Value = szOutput
                
                ' Look for the declaration of the next procedure
                ' Skip over blank lines and comments
                nLineNext = nStartLine
                While nLineNext <= cmModuleCur.CountOfLines _
                        And (cmModuleCur.Lines(nLineNext, 1) = "" _
                        Or Left(Trim(cmModuleCur.Lines(nLineNext, 1)), 1) = "'")
                    nLineNext = nLineNext + 1
                Wend
                
                ' Capture multi-line declarations, denoted by trailing underscore
                ' #### NOTE: This version will stop at the first underscore, NOT the EOL underscore
                nLineCnt = 1
                szOutput = ""
                szLineCur = ""
                While Right(Trim(cmModuleCur.Lines(nLineNext + nLineCnt - 1, 1)), 1) = "_"
                    szLineCur = Trim(cmModuleCur.Lines(nLineNext + nLineCnt - 1, 1))
                    'szLineCur = Left(szLineCur, InStr(szLineCur, "_") - 1)
                    szLineCur = Left(szLineCur, Len(szLineCur) - 1) ' Drop trailing underscore
                    szOutput = szOutput & szLineCur
                    nLineCnt = nLineCnt + 1
                Wend
                'If nLineCnt = 1 Then
                    szOutput = szOutput & Trim(cmModuleCur.Lines(nLineNext + nLineCnt - 1, 1))
                'End If
                
                ' Complete declaration
                wkshOutput.Cells(nRowCur, 2).Value = szOutput
                ' Module name
                wkshOutput.Cells(nRowCur, 3).Value = vbcComponentCur.Name
                ' Function or macro
                If InStr(szOutput, "Function") Then
                    wkshOutput.Cells(nRowCur, 4).Value = "Function"
                Else
                    wkshOutput.Cells(nRowCur, 4).Value = "Macro"
                End If
                
                nStartLine = nStartLine + cmModuleCur.ProcCountLines _
                  (cmModuleCur.ProcOfLine(nStartLine, vbext_pk_Proc), vbext_pk_Proc)
                nRowCur = nRowCur + 1
            Loop
            
        End If ' Standard code module
    Next vbcComponentCur
    
End Sub


''
' Exports all modules in the active workbook to *.bas files.
'
' Called by several of the export routines in this module
Sub CodeModuleExport(szPathOut As String)

    Dim vbcComponentCur As VBComponent, wkbkCurrent As Workbook, vbccList As VBComponents
    Dim nLineCount As Long, nLineCurrent As Long, szNameFunction As String, szNameModule As String
    
    Set wkbkCurrent = ActiveWorkbook
    
    Set vbccList = wkbkCurrent.VBProject.VBComponents
    
    For Each vbcComponentCur In vbccList
        If vbcComponentCur.Type = vbext_ct_StdModule Or vbcComponentCur.Type = vbext_ct_ClassModule Then
            ' Export each code module and class definition
            vbcComponentCur.Export Filename:=szPathOut & vbcComponentCur.Name & ".bas"
            
        End If
    Next
    
End Sub









