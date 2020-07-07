VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Unit_ChooseTargetPresentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'separate tags by comma, e.g.: "@debug,@production"
Const cTags = ""

Dim mOpenTestPresentations As Collection
Dim mTargetSelectionForm As SelectValidationTarget

Public Property Get description() As String

    description = "If only one presentation is available for validation, Slide Validator will " & vbLf _
                    & vbTab & "validate this automatically. Otherwise a dialog will present as list with names " & vbLf _
                    & vbTab & "of possible validation candidates."
End Property

Public Sub test_examples()

    Dim examples As Collection
    Dim example As Variant

    Set examples = New Collection
    examples.Add Array("Rule: if only one presentation for validation is available, don't ask the user to select a target")
    
    examples.Add Array("Example: more than one possible target presentation opened", _
                           "Given two presentations opened in addition to Slide Validator", _
                           "When the validator is looking for available validation targets", _
                           "Then a form will list the two presentations as possible validation targets")
    
    examples.Add Array("Example: only one possible target presentation opened", _
                           "Given only one presentation opened in addition to Slide Validator", _
                           "When the validator is looking for available validation targets", _
                           "Then the presentation is set as validation target")
    
    examples.Add Array("Example: no possible target presentation opened", _
                           "Given no presentation is opened in addition to Slide Validator", _
                           "When the validator is looking for available validation targets", _
                           "Then Slide Validator will return nothing as validation target")
    
    For Each example In examples
        TExampleRunner.run_example example, Me
    Next
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
    
    Dim step_result As String
    Dim config_table As Table
    Dim test_presentation As Presentation
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
            Err.raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = TCase.raise_step_error(Err.Number, Err.description)
End Function

Public Sub before()
    
    Validator.Log.LogLevel = Validator.Log.LVL_WARNING
End Sub

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

Public Property Get Tags() As Variant
    Tags = cTags
End Property

Private Sub pending(pending_msg)
    TCase.pending pending_msg
End Sub

Public Property Get Log() As logger
    Set Log = TCase.Log
End Property

