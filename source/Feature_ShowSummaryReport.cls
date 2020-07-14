VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feature_ShowSummaryReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = "backlog"

Public Property Get description() As String

    description = "Slide Validator will create a validation summary slide" & vbLf & _
                    vbTab & "in a new presentation after every validation"
End Property

Public Sub test_examples()

    Dim examples As Collection
    Dim example As Variant

    Set examples = New Collection
    examples.Add Array("Rule: create a summary report slide after validation was finished")
    examples.Add Array("Example: validation with violations")
    examples.Add Array("Example: validation without violations")
    examples.Add Array("Example: empty target presentation")
    
    examples.Add Array("Rule: setup errors will shown in the summary report slide")
    examples.Add Array("Example: no setup errors")
    examples.Add Array("Example: setup can't find a rule")
    
    'to be discussed - might be better to disable summary via separate config for SlideValidator
    examples.Add Array("Rule: disable summary by hiding summary config slide")
    
    For Each example In examples
        TExampleRunner.run_example example, Me
    Next
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
    
        Case "Given a precondition"
           
        Case "Given another precondition"
            
        Case "When any action happens"

        Case "Then some result is expected"
        
        Case Else
            Err.raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = TFeature.raise_step_error(Err.Number, Err.description)
End Function

Public Sub before()
    
    Validator.Log.LogLevel = Validator.Log.LVL_WARNING
End Sub

Public Sub after()

End Sub

Public Property Get Tags() As Variant

    Tags = cTags
End Property

Private Sub pending(pending_msg)
    
    TFeature.pending pending_msg
End Sub

Public Property Get Log() As logger
    
    Set Log = TFeature.Log
End Property





