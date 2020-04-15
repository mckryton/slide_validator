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
Const tags = "@debug"

Dim mpreTestPresentation As Presentation


Public Sub test_scenarios()

    Dim scenario As Variant

    On Error GoTo error_handler
    scenario = Array("Scenario: config with one parameter", _
                       "Given a config slide for the Feature_FontWhiteList", _
                       "And the slide contains a configuration table with the parameter AllowedFonts with the value Arial", _
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
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
        Case "Given a config slide for the Feature_FontWhiteList"
           Debug.Print "PENDING test step"
           'Set mpreTestPresentation = Application.Presentations.Add
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

Private Sub Class_Terminate()
    mpreTestPresentation.Saved = msoTrue
    mpreTestPresentation.Close
End Sub


