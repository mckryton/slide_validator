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
Const tags = ""

Dim mTestPresentation As Presentation
Dim mConfigSlide As Slide
Dim mRuleConfig As Collection

Public Sub test_scenarios()

    Dim scenario As Variant

    On Error GoTo error_handler
    scenario = Array("Scenario: config with one parameter", _
                       "Given a config slide for the Feature_FontWhiteList", _
                       "And the config slide contains the parameter AllowedFonts with the value Arial", _
                       "When the configuration for the Feature_FontWhiteList is read", _
                       "Then the value for the parameter AllowedFonts is set to Arial" _
                    )
    TScenarioRunner.run_scenario scenario, Me
    Exit Sub

error_handler:
    SystemLogger.log_error "Unit_ReadConfig.runTest"
End Sub

Public Function run_step(pcolStepProps As Collection) As String
    
    Dim step_result As String
    Dim config_table As Table
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
    
        Case "Given a config slide for the Feature_FontWhiteList"
           Set mTestPresentation = Application.Presentations.Add
           Set mConfigSlide = create_config_slide("Feature_FontWhiteList")
           
        Case "Given the config slide contains the parameter AllowedFonts with the value Arial"
            Set config_table = create_config_table(mConfigSlide)
            add_config_parameter config_table, Array("AllowedFonts", "Arial", "parameter for test case Unit_ReadConfig")
            
        Case "When the configuration for the Feature_FontWhiteList is read"
            Set mRuleConfig = Validator.read_config("Feature_FontWhiteList", mTestPresentation)
            
        Case "Then the value for the parameter AllowedFonts is set to Arial"
            TSpec.expect(mRuleConfig).contains Array("AllowedFonts", "Arial")
        
        Case Else
            TScenarioRunner.missingTest pcolStepProps.Item("step_type") & " " & Trim(pcolStepProps.Item("line_body")), Me
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
        SystemLogger.log_error "Unit_ReadConfig.runStep " & pcolStepProps.Item("step_type") & " " & pcolStepProps.Item("line_body")
    End If
End Function

Private Sub add_slide_with_config(pprePresentation As Presentation, pvarConfig As Variant)

    Dim slide_with_config As Slide
    
    On Error GoTo error_handler
    Set slide_with_config = pprePresentation.Slides.AddSlide(1, pprePresentation.SlideMaster.CustomLayouts(7))
    Debug.Print "PENDING add config table to slide"
    Exit Sub
    
error_handler:
    SystemLogger.log_error "clsFeatureFontWhiteList.add_slide_with_textbox"
End Sub

Public Property Get description() As String

    description = "Slide Validator will read all config parameters from it's own slides" & vbLf & _
                    vbTab & "Configuration slides will use the name of the feature class as headline." & vbLf & _
                    vbTab & "A table with 3 columns for parameter name, value and description will held the config."
End Property
Public Sub after_all_steps()

End Sub

Private Sub Class_Terminate()
    If TypeName(mTestPresentation) <> "Nothing" Then
        mTestPresentation.Saved = msoTrue
        mTestPresentation.Close
    End If
End Sub

Public Function create_config_slide(pRuleName As String) As Slide
    
    Dim config_slide As Slide
    
    Set config_slide = mTestPresentation.Slides.AddSlide(1, mTestPresentation.SlideMaster.CustomLayouts(6))
    'todo: remove next line
    mTestPresentation.Slides.AddSlide 1, mTestPresentation.SlideMaster.CustomLayouts(7)
    config_slide.Shapes.Title.TextFrame.TextRange.Text = pRuleName
    config_slide.Name = pRuleName
    Set create_config_slide = config_slide
End Function

Public Function create_config_table(pConfigSlide As Slide) As Table
    
    Dim config_table As Table
    
    Set config_table = pConfigSlide.Shapes.AddTable(1, 3, cm2points(2.33), cm2points(4.7), cm2points(29.21), cm2points(1.42)).Table
    config_table.Cell(1, 1).shape.TextFrame.TextRange.Text = "Parameter"
    config_table.Cell(1, 2).shape.TextFrame.TextRange.Text = "Value"
    config_table.Cell(1, 3).shape.TextFrame.TextRange.Text = "Description"
    Set create_config_table = config_table
End Function

Public Sub add_config_parameter(pConfigTable As Table, pConfigParameter As Variant)

    Dim parameter_row As Row
    
    Set parameter_row = pConfigTable.Rows.Add
    parameter_row.Cells(1).shape.TextFrame.TextRange.Text = pConfigParameter(0)
    parameter_row.Cells(2).shape.TextFrame.TextRange.Text = pConfigParameter(1)
    parameter_row.Cells(3).shape.TextFrame.TextRange.Text = pConfigParameter(2)
End Sub

Public Function cm2points(pValueCm As Double) As Long

    #If Mac Then
        cm2points = CLng(pValueCm * (72 / 2.54))
    #Else
        cm2points = CLng(pValueCm * (96 / 2.54))
    #End If
End Function
