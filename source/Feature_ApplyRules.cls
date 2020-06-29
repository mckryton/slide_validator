VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feature_ApplyRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = ""

Dim mTestPresentation As Presentation

Public Property Get description() As String

    description = "Slide Validator will apply chosen rules for every slide" & vbLf & _
                    vbTab & "in the target presentation and show results and erros in a summary slide"
End Property

Public Sub test_examples()

    Dim examples As Variant
    Dim example As Variant

    examples = Array( _
                    Array("Rule: add comments to the target slide for every rule violation"), _
                    Array("Example: slide with violations", _
                            "Given a presentation with one slide", _
                            "And a rule which fails to validate the slide", _
                            "When the presentation is validazed", _
                            "Then a comment about the violation will be added to the slide" _
                          ), _
                    Array("Rule: remove comments from earlier violation"), _
                    Array("Rule: create a validation summary after validation was finished"), _
                    Array("Example: summary report with violations"), _
                    Array("Example: summary report without violations"), _
                    Array("Example: empty target presentation"), _
                    Array("Rule: setup errors will shown in the validation summary"), _
                    Array("Example: no setup errors"), _
                    Array("Example: setup can't find a rule") _
                )
    For Each example In examples
        TExampleRunner.run_example example, Me
    Next
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
    
        Case "Given a presentation with one slide"
            Set mTestPresentation = Application.Presentations.Add
            mTestPresentation.Slides.AddSlide 1, mTestPresentation.SlideMaster.CustomLayouts(6)
        
        Case "Given a rule which fails to validate the slide"
            pending "have to create a rule for testing purposes"
        
        Case Else
            Err.raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = raise_step_error(Err.Number)
End Function

Private Function raise_step_error(pErrId As Long) As Variant
 
    TExampleRunner.stop_test
    Select Case Err.Number
    Case ERR_ID_STEP_IS_PENDING
        raise_step_error = Array("PENDING", Err.description)
    Case ERR_ID_STEP_IS_MISSING
        raise_step_error = Array("MISSING")
    Case Else
        raise_step_error = Array("FAILED", Err.description)
    End Select
End Function

Public Sub before()
    
    Validator.Log.LogLevel = Validator.Log.LVL_WARNING
End Sub

Public Sub after()

End Sub

Public Property Get Tags() As Variant
    Tags = cTags
End Property

Private Sub pending(pPendingMsg)
    Err.raise ERR_ID_STEP_IS_PENDING, description:=pPendingMsg
End Sub


