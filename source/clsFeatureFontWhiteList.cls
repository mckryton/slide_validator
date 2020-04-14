VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFeatureFontWhiteList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' Description  : Ability apply a white list for fonts
'------------------------------------------------------------------------
'
'Declarations
'separate tags by comma, e.g.: "@debug,@production"
Const tags = "@debug"

'Declare variables
Dim mobjRuleFontWhiteList As clsRuleFontWhiteList
Dim mpreTestPresentation As Presentation

'Options
Option Explicit

Public Sub test_scenarios()

    Dim scenario As Variant

    On Error GoTo error_handler
    scenario = Array("Scenario: first slide ok and second slide with invalid font", _
                       "Given a font white list rule listing Arial", _
                       "And a new presentation", _
                       "And it contains a first slide with a textbox using Arial", _
                       "And it contains a second slide with a textbox using Times New Roman", _
                       "When the rule is applied to the presentation", _
                       "Then a comment about the rule validation is added to the second slide" _
                    )
    basTestScenarioRunner.runScenario scenario, Me
    
    Exit Sub

error_handler:
    basSystemLogger.log_error "clsFeatureFontWhiteList.runTest"
End Sub

Public Sub run_step(pcolStepProps As Collection)
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
        Case "Given a font white list rule listing Arial"
            Set mobjRuleFontWhiteList = New clsRuleFontWhiteList
            mobjRuleFontWhiteList.font_white_list.Add "Arial"
        Case "Given a new presentation"
            Set mpreTestPresentation = Application.Presentations.Add
        Case "Given it contains a first slide with a textbox using Arial"
            add_slide_with_textbox mpreTestPresentation, "Arial"
        Case "Given it contains a second slide with a textbox using Times New Roman"
            add_slide_with_textbox mpreTestPresentation, "Times New Roman"
        Case "When the rule is applied to the presentation"
            basSlideValidator.run_slide_validator mpreTestPresentation, Array(mobjRuleFontWhiteList)
        Case Else
            basTestScenarioRunner.missingTest pcolStepProps.Item("step_type") & " " & Trim(pcolStepProps.Item("line_body")), Me
    End Select
    
    Exit Sub

error_handler:
    basSystemLogger.log_error "clsFeatureFontWhiteList.runStep " & pcolStepProps.Item("step_type") & " " & pcolStepProps.Item("line_body")
End Sub

Private Sub add_slide_with_textbox(pprePresentation As Presentation, pstrFontName As String)

    Dim slide_with_textbox As Slide
    Dim textbox As Shape
    
    On Error GoTo error_handler
    Set slide_with_textbox = pprePresentation.Slides.AddSlide(1, pprePresentation.SlideMaster.CustomLayouts(7))
    Set textbox = slide_with_textbox.Shapes.AddTextbox(msoTextOrientationHorizontal, 200, 200, 400, 200)
    textbox.TextFrame.TextRange.Font.Name = pstrFontName
    textbox.TextFrame.TextRange.Text = "This text is using " & pstrFontName & " as font."
    Exit Sub
    
error_handler:
    basSystemLogger.log_error "clsFeatureFontWhiteList.add_slide_with_textbox"
End Sub
