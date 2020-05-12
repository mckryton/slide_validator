Attribute VB_Name = "TScenarioRunner"
'------------------------------------------------------------------------
' Description  : execute test steps for Gherkin scenarios / examples
'------------------------------------------------------------------------

Option Explicit

Public Const ERR_ID_SCENARIO_SYNTAX_ERROR = vbError + 6010

Dim mTestStopped As Boolean

Public Sub run_scenario(pScenarioLinesArray As Variant, pTestDefinitionObject As Variant)

    Dim intLineIndex As Integer
    Dim colLine As Collection
    Dim strLastStepType As String

    On Error GoTo error_handler
    TScenarioRunner.TestStopped = False
    intLineIndex = 0
    Set colLine = getScenarioLine(pScenarioLinesArray, intLineIndex)
    print_scenario_title colLine.Item("line")
    intLineIndex = intLineIndex + 1
    Do
        Set colLine = getScenarioLine(pScenarioLinesArray, intLineIndex)
        If colLine.Item("line_head") <> "And" Then
            strLastStepType = colLine.Item("line_head")
        End If
        colLine.Remove "step_type"
        colLine.Add strLastStepType, "step_type"
        run_step_line colLine, pTestDefinitionObject
        intLineIndex = intLineIndex + 1
    Loop Until TScenarioRunner.TestStopped = True Or intLineIndex > UBound(pScenarioLinesArray)
    If Not TScenarioRunner.TestStopped Then
        pTestDefinitionObject.after
    End If
    
    Debug.Print
    Exit Sub
    
error_handler:
    If Err.Number = ERR_ID_SCENARIO_SYNTAX_ERROR Then
        SystemLogger.log_error "syntax error: " & Err.description & vbCr & vbLf & "in line >" & colLine.Item("line") & "<"
    Else
        SystemLogger.log_error "TScenarioRunner.runScenario", Join(pScenarioLinesArray, vbTab & vbCr & vbLf)
    End If
End Sub

Private Sub print_scenario_title(pScenarioTitle As String)
    
    If Left(pScenarioTitle, Len("Scenario:")) <> "Scenario:" Then
        Err.Raise ERR_ID_SCENARIO_SYNTAX_ERROR, description:="can't find scenario start"
    Else
        Debug.Print vbTab & pScenarioTitle
    End If
End Sub

Private Sub run_step_line(pStepLine As Collection, pobjTestDefinition As Variant)

    Dim step_result As String

    Select Case pStepLine.Item("step_type")
    Case "Given", "When", "Then"
        step_result = pobjTestDefinition.run_step(pStepLine)
        If step_result = "OK" Then
            Debug.Print vbTab & step_result, vbTab & pStepLine.Item("line")
        ElseIf step_result = "PENDING" Or step_result = "MISSING" Then
            Debug.Print vbTab & step_result, vbTab & pStepLine.Item("line")
            End
        Else
            Debug.Print vbTab & "FAILED", vbTab & pStepLine.Item("line")
            Debug.Print step_result
            End
        End If
    Case Else
        Err.Raise ERR_ID_SCENARIO_SYNTAX_ERROR, description:="unexpected step type " & pStepLine.Item("step_type")
    End Select
End Sub

Public Sub missingTest(pstrStepDefinition As String, pobjCaller As Object)

    On Error GoTo error_handler
    TScenarioRunner.TestStopped = True
    'Debug.Print vbCr & vbLf & "missing test step for >" & pstrStepDefinition & "<" & vbCr & vbLf & "  rule validator: " & TypeName(pobjCaller)
    Exit Sub

error_handler:
    SystemLogger.log_error "TScenarioRunner.missingTest " & pstrStepDefinition
End Sub

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
    SystemLogger.log_error "TScenarioRunner.getScenarioLine"
End Function

Public Property Get TestStopped() As Boolean
    TestStopped = mTestStopped
End Property

Private Property Let TestStopped(ByVal pTestStopped As Boolean)
    mTestStopped = pTestStopped
End Property

Public Sub stop_test()
    TScenarioRunner.TestStopped = True
End Sub

Public Sub pending(pPendingMsg)
    
    Debug.Print "PENDING: " & pPendingMsg
End Sub

