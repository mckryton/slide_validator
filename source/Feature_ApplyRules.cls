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
Const cTags = "wip"

Dim mTestPresentation As Presentation

Public Property Get description() As String

    description = "Slide Validator will apply chosen rules for every slide" & vbLf & _
                    vbTab & "in the target presentation and show results and erros in a summary slide"
End Property

Public Sub test_examples()

    Dim examples As Collection
    Dim example As Variant

    Set examples = New Collection
    examples.Add Array("Rule: add comments to the target slide for every rule violation")
    examples.Add Array("Example: slide with violations", _
                            "Given a presentation with one slide", _
                            "And a rule which fails to validate the slide", _
                            "When the slide is validated", _
                            "Then a comment about the violation will be added to the slide")
                            
    examples.Add Array("Rule: remove comments from earlier violation")
        
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
    run_step = TCase.raise_step_error(Err.Number, Err.description)
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
    TCase.pending pending_msg
End Sub

Public Property Get Log() As logger
    
    Set Log = TCase.Log
End Property

