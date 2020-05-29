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
                    Array("Rule: A slide is a configuration slide if it's layout is named rule_config"), _
                    Array("Example: valid configuration slide", _
                           "Given a config slide using the master slide template for rule config", _
                           "When the slide is validated to be a configuration", _
                           "Then the slide is recognized as a configuration" _
                    ), _
                    Array("Rule: Ignore hidden config slides"), _
                    Array("Rule: Choose Rule from config slide title"), _
                    Array("Example: config slide has a matching rule", _
                            "Given a SlideValidator having a config slide titled FontWhiteList", _
                            "And a matching Rule named Rule_FontWhiteList", _
                            "And a target presentation using Times New Roman for a text", _
                            "When the rules for the next validations are set up", _
                            "Then the setup contains the rule FontWhiteList" _
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
        
        Case "Given a config slide using the master slide template for rule config"
            Set mTestSlideValidator = TSupport.create_slide_validator_pres
            Set mTestConfigSlide = TSupport.create_config_slide(mTestSlideValidator)
            
        Case "When the slide is validated to be a configuration"
            mIsConfigSlide = Validator.is_config_slide(mTestConfigSlide)
            
        Case "Then the slide is recognized as a configuration"
            TSpec.expect(mIsConfigSlide).to_be True
        
        Case "Given a SlideValidator having a config slide titled FontWhiteList"
            Set mTestSlideValidator = TSupport.create_slide_validator_pres
            Set mTestConfigSlide = TSupport.create_config_slide(mTestSlideValidator, "FontWhiteList")
        
        Case "Given a matching Rule named Rule_FontWhiteList"
            Set sample_rule = Validator.get_rule("FontWhiteList")
            TSpec.expect(TypeName(sample_rule)).to_be "Rule_FontWhiteList"
        
        Case "Given a target presentation using Times New Roman for a text"
            Set mTestPresentation = Application.Presentations.Add(msoFalse)
            TSupport.add_slide_with_textbox mTestPresentation, "Times New Roman"
            
        Case "When the rules for the next validations are set up"
            Set mValidationSetup = Validator.setup_rules(mTestSlideValidator)
            
        Case "Then the setup contains the rule FontWhiteList"
            TSpec.expect(ExtraVBA.existsItem("FontWhiteList", mValidationSetup.ActiveRules)).to_be True
            
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
