VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feature_Rule_Permitted_Fonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' Description  : Ability apply a white list for fonts
'------------------------------------------------------------------------
Option Explicit

'separate tags by comma, e.g.: "debug,production"
Const cTags = "wip"

Dim mLogger As Logger
Dim mobjRulePermittedFonts As Rule_Permitted_Fonts
Dim mpreTestPresentation As Presentation

Public Property Get description() As String

    description = "The presentation should only use permitted fonts listed in the configuration"
End Property

Public Sub test_examples()

    Dim examples As Variant
    Dim example As Variant

    examples = Array( _
                    Array("Example: one slide with valid and one slide with invalid font", _
                           "Given Arial is a permitted font", _
                           "And a new presentation", _
                           "And it contains a slide with a textbox using Arial", _
                           "And it contains a slide with a textbox using Times New Roman", _
                           "When the rule is applied to the presentation", _
                           "Then a comment about the rule validation is added to the second slide" _
                        ) _
                )
    For Each example In examples
        TExampleRunner.run_example example, Me
    Next
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
    
    Dim step_result As String
    Dim rule_config As Collection
    Dim validation_setup As ValidationSetup
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
        Case "Given Arial is a permitted font"
            Set mobjRulePermittedFonts = New Rule_Permitted_Fonts
            Set rule_config = New Collection
            rule_config.Add "Arial", "PermittedFonts"
            mobjRulePermittedFonts.Config = rule_config
            
        Case "Given a new presentation"
            Set mpreTestPresentation = Application.Presentations.Add
            
        Case "Given it contains a slide with a textbox using Arial"
            TSupport.add_slide_with_textbox mpreTestPresentation, "Arial"
            
        Case "Given it contains a slide with a textbox using Times New Roman"
            TSupport.add_slide_with_textbox mpreTestPresentation, "Times New Roman"
            
        Case "When the rule is applied to the presentation"
            Set validation_setup = New ValidationSetup
            validation_setup.ActiveRules.Add mobjRulePermittedFonts, "Permitted_Fonts"
            Validator.validate_presentation mpreTestPresentation, validation_setup, p_silent:=True
            
        Case "Then a comment about the rule validation is added to the second slide"
            TSpec.expect(mpreTestPresentation.Slides(1).Comments.Count).to_be 1
            TSpec.expect(mpreTestPresentation.Slides(2).Comments.Count).to_be 0
            
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
        raise_step_error = Array("PENDING")
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
    mpreTestPresentation.Saved = msoTrue
    mpreTestPresentation.Close
End Sub

Public Property Get Tags() As Variant
    Tags = cTags
End Property

Private Sub pending(pPendingMsg)
    Err.raise ERR_ID_STEP_IS_PENDING, description = pPendingMsg
End Sub

Private Property Get Log() As Logger
    
    If TypeName(mLogger) = "Nothing" Then
        Set mLogger = New Logger
    End If
    Set Log = mLogger
End Property

Public Property Let Log(ByVal pLogger As Logger)
    Set mLogger = pLogger
End Property

Private Sub Class_Terminate()
    Set mLogger = Nothing
End Sub

