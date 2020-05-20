VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Unit_ReadConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' Description  : Ability to read config from well structured slides
'------------------------------------------------------------------------
Option Explicit

'separate tags by comma, e.g.: "@debug,@production"
Const cTags = "wip"

Const RULE_CFG_SLIDE_IDX = 2
Const ERR_ID_MISSING_CFG_MASTER_SLIDE = vbError + 7000

Dim mTestPresentation As Presentation
Dim mConfigSlide As Slide
Dim mRuleConfig As Collection

Public Sub test_examples()

    Dim scenarios As Variant
    Dim scenario As Variant

    On Error GoTo error_handler
    scenarios = Array( _
                    Array("Scenario: config with one parameter", _
                           "Given a config slide for the rule FontWhiteList", _
                           "And the config slide contains the parameter AllowedFonts with the value Arial", _
                           "When the configuration for the Feature_FontWhiteList is read", _
                           "Then the value for the parameter AllowedFonts is set to Arial" _
                    ) _
                )
    For Each scenario In scenarios
        TExampleRunner.run_example scenario, Me
    Next
    Exit Sub

error_handler:
    SystemLogger.log_error "Unit_ReadConfig.test_scenarios"
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
    
    Dim step_result As String
    Dim config_table As Table
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
    
        Case "Given a config slide for the rule FontWhiteList"
           Set mTestPresentation = create_config_presentation
           Set mConfigSlide = create_config_slide(mTestPresentation, "FontWhiteList")
           
        Case "Given the config slide contains the parameter AllowedFonts with the value Arial"
            Set config_table = create_config_table(mConfigSlide)
            add_config_parameter config_table, Array("AllowedFonts", "Arial", "parameter for test case Unit_ReadConfig")
            
        Case "When the configuration for the Feature_FontWhiteList is read"
            Set mRuleConfig = Validator.read_config("FontWhiteList", mTestPresentation)
            TSpec.expect(TypeName(mRuleConfig)).to_be "Collection"
            
        Case "Then the value for the parameter AllowedFonts is set to Arial"
            TSpec.expect(IsArray(mRuleConfig(1))).to_be True
            TSpec.expect(mRuleConfig(1)).to_be Array("AllowedFonts", "Arial")
        
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
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

Public Property Get description() As String

    description = "Slide Validator will read all config parameters from it's own slides" & vbLf & _
                    vbTab & "Configuration slides will use the name of the feature class as headline." & vbLf & _
                    vbTab & "A table with 3 columns for parameter name, value and description will held the config."
End Property

Public Sub after()
    If TypeName(mTestPresentation) <> "Nothing" Then
        mTestPresentation.Saved = msoTrue
        mTestPresentation.Close
    End If
End Sub

Public Property Get Tags() As Variant
    Tags = cTags
End Property

Private Sub pending(pPendingMsg)
    Err.Raise ERR_ID_STEP_IS_PENDING, description = pPendingMsg
End Sub

Private Function create_config_presentation() As Presentation

    Dim config_presentation As Presentation
    Dim slide_validator As Presentation
    Dim master_slide As CustomLayout
        
    On Error GoTo error_handler
    Set slide_validator = Application.Presentations("SlideValidator.pptm")
    Set config_presentation = Application.Presentations.Add
    
    Set master_slide = slide_validator.SlideMaster.CustomLayouts(RULE_CFG_SLIDE_IDX)
    If master_slide.Name <> "rule_config" Then
        Err.Raise ERR_ID_MISSING_CFG_MASTER_SLIDE, description = "could not find rule_config at idx " & RULE_CFG_SLIDE_IDX & " found >" & master_slide.Name & "< instead"
    End If
    master_slide.Copy
    config_presentation.SlideMaster.CustomLayouts.Paste RULE_CFG_SLIDE_IDX
    
    Set create_config_presentation = config_presentation
    Exit Function
    
error_handler:
    SystemLogger.log_error "clsFeatureFontWhiteList.create_config_presentation"
End Function

Private Function create_config_slide(pConfigPresentation As Presentation, pRuleName As String) As Slide
    
    Dim config_slide As Slide
    
    Set config_slide = pConfigPresentation.Slides.AddSlide(1, pConfigPresentation.SlideMaster.CustomLayouts(RULE_CFG_SLIDE_IDX))
    config_slide.Shapes.Title.TextFrame.TextRange.Text = pRuleName
    config_slide.Name = pRuleName
    Set create_config_slide = config_slide
End Function

Private Function create_config_table(pConfigSlide As Slide) As Table
    
    Dim config_table As Table
    
    Set config_table = pConfigSlide.Shapes.AddTable(1, 3, ExtraPpt.cm2points(2.33), ExtraPpt.cm2points(4.7), ExtraPpt.cm2points(29.21), ExtraPpt.cm2points(1.42)).Table
    config_table.Cell(1, 1).shape.TextFrame.TextRange.Text = "Parameter"
    config_table.Cell(1, 2).shape.TextFrame.TextRange.Text = "Value"
    config_table.Cell(1, 3).shape.TextFrame.TextRange.Text = "Description"
    Set create_config_table = config_table
End Function

Private Sub add_config_parameter(pConfigTable As Table, pConfigParameter As Variant)

    Dim parameter_row As Row
    
    Set parameter_row = pConfigTable.Rows.Add
    parameter_row.Cells(1).shape.TextFrame.TextRange.Text = pConfigParameter(0)
    parameter_row.Cells(2).shape.TextFrame.TextRange.Text = pConfigParameter(1)
    parameter_row.Cells(3).shape.TextFrame.TextRange.Text = pConfigParameter(2)
End Sub

