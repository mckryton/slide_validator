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
Const cTags = ""

Dim mOpenTestPresentations As Collection
Dim mTargetSelectionForm As SelectValidationTarget

Public Sub test_scenarios()

    Dim scenarios As Variant
    Dim scenario As Variant

    On Error GoTo error_handler
    scenarios = Array( _
                    Array("Scenario: more than one possible target presentation opened", _
                           "Given two presentations opened in addition to Slide Validator", _
                           "When the validator is looking for available validation targets", _
                           "Then a form will list the two presentations as possible validation targets" _
                        ), _
                    Array("Scenario: only one possible target presentation opened", _
                           "Given only one presentation opened in addition to Slide Validator", _
                           "When the validator is looking for available validation targets", _
                           "Then the presentation is set as validation target" _
                        ), _
                    Array("Scenario: no possible target presentation opened", _
                           "Given no presentation is opened in addition to Slide Validator", _
                           "When the validator is looking for available validation targets", _
                           "Then Slide Validator will return nothing as validation target" _
                        ) _
                    )
    For Each scenario In scenarios
        TScenarioRunner.run_scenario scenario, Me
    Next
    Exit Sub

error_handler:
    SystemLogger.log_error "Unit_ChooseTarget.test_scenarios"
End Sub

Public Function run_step(pcolStepProps As Collection) As String
    
    Dim step_result As String
    Dim config_table As Table
    Dim test_presentation As presentation
    Dim index As Integer
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
    
        Case "Given two presentations opened in addition to Slide Validator"
            Set mOpenTestPresentations = New Collection
            For index = 1 To 2
                Set test_presentation = Presentations.Add(msoTrue)
                mOpenTestPresentations.Add test_presentation, test_presentation.Name
            Next
        Case "When the validator is looking for available validation targets"
            Set mTargetSelectionForm = Validator.get_validation_target_form
        Case "Then a form will list the two presentations as possible validation targets"
            TSpec.expect(UBound(mTargetSelectionForm.lstPresentations.List)).to_be 1
            
        Case "Given only one presentation opened in addition to Slide Validator"
            Set mOpenTestPresentations = New Collection
            Set test_presentation = Presentations.Add(msoTrue)
            mOpenTestPresentations.Add test_presentation
        Case "Then the presentation is set as validation target"
            TSpec.expect(Validator.ValidationTarget).to_be mOpenTestPresentations(1)
            
        Case "Given no presentation is opened in addition to Slide Validator"
        Case "Then Slide Validator will return nothing as validation target"
            TSpec.expect(TypeName(Validator.ValidationTarget)).to_be "Nothing"
            
        Case Else
            TScenarioRunner.missingTest pcolStepProps.Item("step_type") & " " & Trim(pcolStepProps.Item("line_body")), Me
            run_step = "MISSING"
            Exit Function
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


Public Sub after()

    Dim test_presentation As Variant
    
    If TypeName(mOpenTestPresentations) <> "Nothing" Then
        For Each test_presentation In mOpenTestPresentations
            test_presentation.Saved = msoTrue
            test_presentation.Close
        Next
        Set mOpenTestPresentations = Nothing
    End If
End Sub

Public Property Get tags() As Variant
    tags = cTags
End Property
