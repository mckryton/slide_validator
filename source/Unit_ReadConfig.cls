VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Unit_ReadConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = ""

Const RULE_CFG_SLIDE_IDX = 2

Dim mSlideValidatorConfig As Presentation
Dim mConfigSlide As Slide
Dim mRuleConfig As Collection

Public Property Get description() As String

    description = "Slide Validator will read all config parameters from it's own slides" & vbLf & _
                    vbTab & "Configuration slides will use the name of the feature class as headline." & vbLf & _
                    vbTab & "A table with 3 columns for parameter name, value and description will held the config."
End Property

Public Sub test_examples()

    Dim examples As Collection
    Dim example As Variant

    Set examples = New Collection
    examples.Add Array("Rule: convert the content of the first two columns of the the config table into key value pairs")
    
    examples.Add Array("Example: config with one parameter", _
                           "Given a SlideValidator Setup with a config slide for the rule ""Permitted Fonts""", _
                           "And the config slide contains the parameter PermittedFonts with the value Arial", _
                           "When the configuration for the Feature_FontWhiteList is read", _
                           "Then the value for the parameter PermittedFonts is set to Arial")
    
    For Each example In examples
        TExampleRunner.run_example example, Me
    Next
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
    
    Dim config_table As Table
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
    
        Case "Given a SlideValidator Setup with a config slide for the rule ""Permitted Fonts"""
            Set mSlideValidatorConfig = TSupport.create_slide_validator_pres
            Set mConfigSlide = TSupport.create_config_slide(mSlideValidatorConfig, "Permitted Fonts")
           
        Case "Given the config slide contains the parameter PermittedFonts with the value Arial"
            Set config_table = TSupport.get_config_table(mConfigSlide)
            TSupport.add_config_parameter config_table, Array("PermittedFonts", "Arial", "parameter for test case Unit_ReadConfig")
            
        Case "When the configuration for the Feature_FontWhiteList is read"
            Set mRuleConfig = Validator.get_rule_config(mConfigSlide)
            TSpec.expect(TypeName(mRuleConfig)).to_be "Collection"
            
        Case "Then the value for the parameter PermittedFonts is set to Arial"
            TSpec.expect(mRuleConfig("PermittedFonts")).to_be "Arial"
        
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
    If TypeName(mSlideValidatorConfig) <> "Nothing" Then
        mSlideValidatorConfig.Saved = msoTrue
        mSlideValidatorConfig.Close
    End If
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

