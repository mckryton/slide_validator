Attribute VB_Name = "basCuke"
'------------------------------------------------------------------------
' Description  : cuke alike test support functions
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

    Dim strScenarioLine As String
    Dim varWords As Variant
    Dim intLineIndex As Integer
    Dim strStepType As String
    Dim strStepDef As String
    Dim strSyntaxErrMsg As String

    On Error GoTo error_handler
    mStopTestExecution = False
    intLineIndex = 0
    'print scenario title
    strScenarioLine = pvarScenario(intLineIndex)
    If Left(strScenarioLine, Len("Scenario:")) <> "Scenario:" Then
        strSyntaxErrMsg = "can't find scenario start"
        GoTo syntax_error
    Else
        Debug.Print strScenarioLine
    End If
    'run given steps
    intLineIndex = intLineIndex + 1
    strScenarioLine = pvarScenario(intLineIndex)
    varWords = Split(strScenarioLine, " ")
    strStepDef = Right(strScenarioLine, Len(strScenarioLine) - Len(strStepType))
    Do
        If varWords(0) <> "And" Then strStepType = varWords(0)
        Select Case strStepType
        Case "Given"
            pobjCaller.runGivenSteps strStepDef
        Case "When"
            pobjCaller.runWhenSteps strStepDef
        Case "Then"
            pobjCaller.runThenSteps strStepDef
        Case Else
            strSyntaxErrMsg = "unexpected step type " & strStepType
            GoTo syntax_error
        End Select
        intLineIndex = intLineIndex + 1
        strScenarioLine = pvarScenario(intLineIndex)
        varWords = Split(strScenarioLine, " ")
        strStepType = varWords(0)
    Loop Until mStopTestExecution = True Or intLineIndex = UBound(pvarScenario)

    Exit Sub
    
syntax_error:
    basSystemLogger.log_error "syntax error: " & strSyntaxErrMsg & vbCr & vbLf & "in line >" & strScenarioLine & "<"
error_handler:
    basSystemLogger.log_error "basCuke.runScenario " & vbCr & vbLf & Join(pvarScenario, vbCr & vbLf)
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
    basSystemLogger.log_error "basCuke.missingTest " & pstrStepDefinition
End Sub

