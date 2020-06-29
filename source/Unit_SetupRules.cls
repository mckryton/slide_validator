VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Unit_SetupRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' Description  : Ability to find configs and apply them to rule functions
'------------------------------------------------------------------------
Option Explicit

'separate tags by comma, e.g.: "@debug,@production"
Const cTags = ""

Dim mTestPresentation As Presentation
Dim mTestSlideValidator As Presentation
Dim mTestConfigSlide As Slide
Dim mIsConfigSlide As Boolean
Dim mValidationSetup As ValidationSetup

Public Property Get description() As String

    description = "Slide Validator will find configuration for rule from it's own slides" & vbLf & _
                    vbTab & "and find matching rule classes using the configuration title"
End Property

Public Sub test_examples()

    Dim examples As Variant
    Dim example As Variant

    examples = Array( _
                    Array("Rule: a slide is a rule config slide if it has a title starting with Rule* and contains a table with 3 columns Parameter, Value and Description"), _
                    Array("Example: complete rule config slide", _
                           "Given a slide with the title ""Rule: just for test"" and a table with 3 columns ""Parameter,Value,Description""", _
                           "When the slide is tested for being a rule config slide", _
                           "Then the slide is considered to be a rule config slide" _
                    ), _
                    Array("Rule: rules are setup from config slides in SlideValidator, hidden slides are ignored"), _
                    Array("Rule: the slide title defines which rule object is used for validation"), _
                    Array("Example: config slide has a matching rule", _
                            "Given a SlideValidator having a config slide titled ""Permitted Fonts""", _
                            "And a matching Rule named ""Rule_Permitted_Fonts""", _
                            "And a target presentation using Times New Roman for a text", _
                            "When the rules for the next validations are set up", _
                            "Then the setup contains the rule ""Permitted_Fonts""" _
                    ) _
                )


    For Each example In examples
        TExampleRunner.run_example example, Me
    Next
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
        
    Dim step_result As String
    Dim config_table As Table
    Dim sample_rule As Variant
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
        
        Case "Given a slide with the title ""Rule: just for test"" and a table with 3 columns ""Parameter,Value,Description"""
            Set mTestSlideValidator = TSupport.create_slide_validator_pres
            Set mTestConfigSlide = TSupport.create_config_slide(mTestSlideValidator, "just for test")

        Case "When the slide is tested for being a rule config slide"
            mIsConfigSlide = Validator.is_config_slide(mTestConfigSlide)
            
        Case "Then the slide is considered to be a rule config slide"
            TSpec.expect(mIsConfigSlide).to_be True
        
        Case "Given a SlideValidator having a config slide titled ""Permitted Fonts"""
            Set mTestSlideValidator = TSupport.create_slide_validator_pres
            Set mTestConfigSlide = TSupport.create_config_slide(mTestSlideValidator, "Permitted Fonts")
        
        Case "Given a matching Rule named ""Rule_Permitted_Fonts"""
            'the rule object is created just to check it is available
            Set sample_rule = Validator.get_rule("Permitted_Fonts")
            TSpec.expect(TypeName(sample_rule)).to_be "Rule_Permitted_Fonts"
        
        Case "Given a target presentation using Times New Roman for a text"
            Set mTestPresentation = Application.Presentations.Add(msoFalse)
            TSupport.add_slide_with_textbox mTestPresentation, "Times New Roman"
            
        Case "When the rules for the next validations are set up"
            Set mValidationSetup = Validator.setup_rules(mTestSlideValidator)
            
        Case "Then the setup contains the rule ""Permitted_Fonts"""
            TSpec.expect(ExtraVBA.existsItem("Permitted_Fonts", mValidationSetup.ActiveRules)).to_be True
            
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

    If TypeName(mTestSlideValidator) <> "Nothing" Then
        mTestSlideValidator.Saved = msoTrue
        mTestSlideValidator.Close
    End If
    If TypeName(mTestPresentation) <> "Nothing" Then
        mTestPresentation.Saved = msoTrue
        mTestPresentation.Close
    End If
End Sub

Public Property Get Tags() As Variant
    Tags = cTags
End Property

Private Sub pending(pPendingMsg)
    Err.raise ERR_ID_STEP_IS_PENDING, description = pPendingMsg
End Sub
