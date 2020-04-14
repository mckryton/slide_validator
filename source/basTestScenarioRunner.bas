Attribute VB_Name = "basTestScenarioRunner"
'------------------------------------------------------------------------
' Description  : execute test steps for Gherkin scenarios / examples
'------------------------------------------------------------------------
'
'Declarations


'Declare variables
Dim mStopTestExecution As Boolean

'Options
Option Explicit
'-------------------------------------------------------------
' Description   : run all acceptance tests
' Parameter     : pvarScenario  - Gherkin scenario as variant array of strings
'                 pobjCaller    - reference to the calling test class
'-------------------------------------------------------------
Public Sub runScenario(pvarScenario As Variant, pobjCaller As Variant)

    Dim intLineIndex As Integer
    Dim colLine As Collection
    Dim strSyntaxErrMsg As String
    Dim strLastStepType As String

    On Error GoTo error_handler
    mStopTestExecution = False
    intLineIndex = 0
    Set colLine = getScenarioLine(pvarScenario, intLineIndex)
    'TODO: refactor add functioon for print scenario title
    If Left(colLine.Item("line"), Len("Scenario:")) <> "Scenario:" Then
        strSyntaxErrMsg = "can't find scenario start"
        GoTo syntax_error
    Else
        Debug.Print colLine.Item("line")
    End If
    'TODO: refactor add function execute step
    intLineIndex = intLineIndex + 1
    Set colLine = getScenarioLine(pvarScenario, intLineIndex)
    Do
        If colLine.Item("line_head") <> "And" Then
            strLastStepType = colLine.Item("line_head")
        End If
        colLine.Remove "step_type"
        colLine.Add strLastStepType, "step_type"
        Select Case colLine.Item("step_type")
        Case "Given", "When", "Then"
            pobjCaller.run_step colLine
            'TODO: print step + test result
        Case Else
            strSyntaxErrMsg = "unexpected step type " & colLine.Item("step_type")
            GoTo syntax_error
        End Select
        intLineIndex = intLineIndex + 1
        Set colLine = getScenarioLine(pvarScenario, intLineIndex)
    Loop Until mStopTestExecution = True Or intLineIndex = UBound(pvarScenario)

    Exit Sub
    
'TODO: refactor add function for raising syntax error
syntax_error:
    basSystemLogger.log_error "syntax error: " & strSyntaxErrMsg & vbCr & vbLf & "in line >" & colLine.Item("line") & "<"
error_handler:
    basSystemLogger.log_error "basTestScenarioRunner.runScenario", Join(pvarScenario, vbTab & vbCr & vbLf)
End Sub
'-------------------------------------------------------------
' Description   : tell about missing test for a step definition
' Parameter     : pstrStepDefinition  - a Gherkin step definition as string
'                 pobjCaller          - reference to the calling test class
'-------------------------------------------------------------
Public Sub missingTest(pstrStepDefinition As String, pobjCaller As Object)

    On Error GoTo error_handler
    mStopTestExecution = True
    Debug.Print vbCr & vbLf & "missing test step for >" & pstrStepDefinition & "<" & vbCr & vbLf & "  rule validator: " & TypeName(pobjCaller)
    Exit Sub

error_handler:
    basSystemLogger.log_error "basTestScenarioRunner.missingTest " & pstrStepDefinition
End Sub


'-------------------------------------------------------------
' Description   : pick a line from a given scenario
' Parameter     : pvarScenario  - Gherkin scenario as variant array of strings
'                 pintLineIndex - line number
' Returnvalue   : line properties as collection
'-------------------------------------------------------------
Public Function getScenarioLine(pvarScenario As Variant, pintLineIndex As Integer) As Collection

    Dim colLineProps As Collection
    Dim strLine As String     'the whole line
    Dim strStepType As String 'e.g. Given, When, Then, And
    Dim strStepDef As String  ' everything behind step type
    Dim varWords As Variant   'all words of the line as array
    
    On Error GoTo error_handler
    strLine = Trim(pvarScenario(pintLineIndex))
    varWords = Split(strLine, " ")
    strStepType = varWords(0)
    strStepDef = Right(strLine, Len(strLine) - Len(strStepType))
    Set colLineProps = New Collection
    With colLineProps
        .Add strLine, "line"
        .Add strStepType, "line_head"
        .Add strStepDef, "line_body"
        .Add vbNullString, "step_type"      'step type depends on context, e.g. previous steps
    End With
    Set getScenarioLine = colLineProps
    Exit Function

error_handler:
    basSystemLogger.log_error "basTestScenarioRunner.getScenarioLine"
End Function
