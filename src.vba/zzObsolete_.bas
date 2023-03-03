Option Explicit




' Set pmoCurrentRun = New pmoClsEnvironment


Sub stub()
    Dim vntRc As Variant, decCurrent As ClsDeclaration, Test As Integer
    Set decCurrent = New ClsDeclaration
    
    vntRc = decCurrent.ParmListParse("szButtonText As String, szNameMacroNew As String")
End Sub

Sub ProceduresListWithComments()
' Based on code from Walkenbach

    Dim vbpProjectCur As VBIDE.VBProject, vbcComponentCur As VBComponent, cmModuleCur As CodeModule
    Dim nLineStart As Long, szOutput As String, szProcName As String, nLineNext As Long, nLineCnt As Integer
    Dim nColumnCur As Long, nRowCurOutput As Long, wkshOutput As Worksheet, arrFacts() As String
    Dim szLineCur As String, nLineDeclare As Long, declCurrent As ClsDeclaration, vntRc As Variant
    Dim colDeclarations As Collection
    
    Set colDeclarations = New Collection
    
    ' Use the active workbook
    Set vbpProjectCur = ActiveWorkbook.VBProject
    ' Define the output page
    Set wkshOutput = ActiveWorkbook.Worksheets("CodeList")
    
    nRowCurOutput = 2
   ' Loop through the VB components (e.g., modules)
    For Each vbcComponentCur In vbpProjectCur.VBComponents
        Set cmModuleCur = vbcComponentCur.CodeModule
        
        ' Expand to include class modules (currently just code modules) vbext_ct_ClassModule
        If vbcComponentCur.Type = vbext_ct_StdModule Then
            'declCurrent.Clear
            
            'szOutput = szOutput & vbNewLine
            ' Start line will advance to the start of each proc
            nLineStart = cmModuleCur.CountOfDeclarationLines + 1
            
            ' nLineStart is the first blank line or proc line after the prior declaration
            Do Until nLineStart >= cmModuleCur.CountOfLines
                ' Determine name of next procedure
                Set declCurrent = New ClsDeclaration
                szProcName = cmModuleCur.ProcOfLine(nLineStart, vbext_pk_Proc)
                declCurrent.Name = szProcName
                
                
                ' Look for the declaration of the next procedure
                ' #### (ProcStartLine) ProcBodyLine gives line number for decl. ####
                nLineDeclare = cmModuleCur.ProcBodyLine(szProcName, vbext_pk_Proc)
                
                ' Capture all leading comments (preceeding proc declaration)
                'szOutput = cmModuleCur.Lines(nLineStart, nLineDeclare - nLineStart)
                nLineNext = nLineStart
                nLineCnt = 1
                szOutput = ""
                szLineCur = ""
                While nLineNext <= cmModuleCur.CountOfLines _
                        And (cmModuleCur.Lines(nLineNext, 1) = "" _
                        Or Left(Trim(cmModuleCur.Lines(nLineNext, 1)), 1) = "'")
                    szLineCur = Trim(cmModuleCur.Lines(nLineNext + nLineCnt - 1, 1))
                    If Left(szLineCur, 1) = "'" Then
                        szLineCur = Trim(Mid(szLineCur, 2))
                    End If
                    If szLineCur <> "" Then
                        szOutput = szOutput & szLineCur & "<EOL>" & vbNewLine  ' "<EOL>" "<EOP>" ?
                    End If
                    nLineNext = nLineNext + 1
                Wend
                declCurrent.Description = szOutput
                
                ' Capture multi-line declarations, denoted by trailing underscore
                ' Ignores internal underscores, drops only the trailing underscore
                nLineNext = nLineDeclare
                nLineCnt = 1
                szOutput = ""
                szLineCur = ""
                While Right(Trim(cmModuleCur.Lines(nLineNext + nLineCnt - 1, 1)), 1) = "_"
                    szLineCur = Trim(cmModuleCur.Lines(nLineNext + nLineCnt - 1, 1))
                    szLineCur = Left(szLineCur, Len(szLineCur) - 1) ' Drop trailing underscore
                    'szOutput = szOutput & szLineCur & "<EOL>" & vbNewLine  ' "<EOL>" "<EOP>" ?
                    szOutput = szOutput & szLineCur
                    nLineCnt = nLineCnt + 1
                Wend
                ' Capture the last line of the proc declaration, w/o trailing underscore
                szOutput = szOutput & Trim(cmModuleCur.Lines(nLineNext + nLineCnt - 1, 1))
                
                ' Complete declaration
                vntRc = declCurrent.ProcStatementParse(szOutput)
                
                ' Capture comments following proc declaration; stop at first non-blank, non-comment line
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
                
                ' Skips remaining lines in function
                ' Remove this; look at every line
                ' Replace with flag for start of next proc: ProcStartLine + ProcCountLines
                ' Capture comments up to first Dim or other statement
                nLineStart = nLineStart + cmModuleCur.ProcCountLines _
                  (cmModuleCur.ProcOfLine(nLineStart, vbext_pk_Proc), vbext_pk_Proc)
                nRowCurOutput = nRowCurOutput + 1
                
                ' Add each proc declaration to the collection
                colDeclarations.Add Item:=declCurrent, Key:=declCurrent.Name '& "1"
            
                ' Write the function/macro name, etc., to the CodeList sheet
                wkshOutput.Cells(nRowCurOutput, 1).Value = declCurrent.Name
                wkshOutput.Cells(nRowCurOutput, 2).Value = declCurrent.Declaration
                wkshOutput.Cells(nRowCurOutput, 3).Value = declCurrent.Description
                declCurrent.Module = vbcComponentCur.Name
                wkshOutput.Cells(nRowCurOutput, 4).Value = declCurrent.Module
                wkshOutput.Cells(nRowCurOutput, 6).Value = declCurrent.Scope
                
                If declCurrent.Macro = True Then
                    wkshOutput.Cells(nRowCurOutput, 5).Value = "Macro"
                Else
                    wkshOutput.Cells(nRowCurOutput, 5).Value = "Function"
                End If
            Loop
            
        End If ' Standard code module
    Next vbcComponentCur
    
'    ' Write output files: colDeclarations
'    For Each declCurrent In colDeclarations
'        vntRc = HtmlProcFileCreate(declCurrent, "C:\Richard\Excel\CodeParser\Projects\PMO\")
'    Next declCurrent
    
End Sub



