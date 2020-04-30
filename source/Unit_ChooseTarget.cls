VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Unit_ChooseTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' Description  : Ability to presentation to be validated from a list
'------------------------------------------------------------------------
Option Explicit

'separate tags by comma, e.g.: "@debug,@production"
Const cTags = "wip"

Dim mOpenTestPresentations As Collection

Public Sub test_scenarios()

    Dim scenario As Variant

    On Error GoTo error_handler
    scenario = Array("Scenario: more than one possible target presentation", _
                       "Given two presentations opened in addition to Slide Validator", _
                       "When Slide Validator selects a presentation for validation", _
                       "Then a list will show the names of the two presentations as possible validation targets" _
                    )
    TScenarioRunner.run_scenario scenario, Me
    Exit Sub

error_handler:
    SystemLogger.log_error "Unit_ChooseTarget.runTest"
End Sub

Public Function run_step(pcolStepProps As Collection) As String
    
    Dim step_result As String
    Dim config_table As Table
    Dim test_presentation As Presentation
    Dim index As Integer
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
    
        Case "Given two presentations Pres_1 and Pres_2 opened in addition to Slide Validator"
            Set mOpenTestPresentations = New Collection
            For index = 1 To 2
                Set test_presentation = Presentations.Add(msoTrue)
                mOpenTestPresentations.Add test_presentation, test_presentation.Name
            Next
        
        Case "When Slide Validator selects the first presentation for validation"
            Validator.select_validation_target
        
        Case "Then a list will show the names of the two presentations as possible validation targets"
            TScenarioRunner.pending "verify showing a form with open presentations"
        
        Case Else
            TScenarioRunner.missingTest pcolStepProps.Item("step_type") & " " & Trim(pcolStepProps.Item("line_body")), Me
    End Select
    run_step = "OK"
    Exit Function

error_handler:
    If Err.Number = TSpec.ERR_ID_EXPECTATION_FAILED Then
        TScenarioRunner.stop_test
        run_step = Err.description
        Err.Clear
        Exit Function
    Else
        SystemLogger.log_error "Unit_ChooseTarget.runStep " & pcolStepProps.Item("step_type") & " " & pcolStepProps.Item("line_body")
    End If
End Function

Public Property Get description() As String

    description = "If only one presentation is available for validation, Slide Validator will " & vbLf _
                    & "validate this automatically. Otherwise a dialog will present as list with names " & vbLf _
                    & "of possible validation candidates."
End Property


Public Sub after_all_steps()

End Sub

Public Property Get tags() As Variant
    tags = cTags
End Property


