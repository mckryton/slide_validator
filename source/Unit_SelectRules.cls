VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Unit_SelectRules"
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
Dim mConfigSlide As Slide
Dim mRuleConfig As Collection

Public Sub test_examples()

    Dim scenarios As Variant
    Dim scenario As Variant

    On Error GoTo error_handler
    scenarios = Array( _
                    Array("Rule: A slide is a configuration slide if it's layout is named Rule_Config"), _
                    Array("Scenario: valid configuration slide", _
                           "Given a config slide", _
                           "When the slide is validated to be a configuration", _
                           "Then the validation result confirms the configuration" _
                    ) _
                )
    For Each scenario In scenarios
        TExampleRunner.run_example scenario, Me
    Next
    Exit Sub

error_handler:
    SystemLogger.log_error "Unit_ReadConfig.test_scenarios"
End Sub

Public Function run_step(pcolStepProps As Collection) As String
    
    Dim step_result As String
    Dim config_table As Table
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
        Case "Given a config slide"
            Set mTestPresentation = ActivePresentation
        Case Else
            TExampleRunner.missingTest pcolStepProps.Item("step_type") & " " & Trim(pcolStepProps.Item("line_body")), Me
    End Select
    run_step = "OK"
    Exit Function

error_handler:
    Select Case Err.Number
    Case TSpec.ERR_ID_EXPECTATION_FAILED
        TExampleRunner.stop_test
        run_step = Err.description
        Err.Clear
    Case ERR_ID_STEP_IS_PENDING
        TExampleRunner.stop_test
        run_step = "PENDING"
        Err.Clear
    Case Else
        SystemLogger.log_error "Unit_ReadConfig.runStep " & pcolStepProps.Item("step_type") & " " & pcolStepProps.Item("line_body")
    End Select
End Function

Public Property Get description() As String

    description = "Slide Validator will find configuration for rule from it's own slides" & vbLf & _
                    vbTab & "and apply configuration found to matching rules"
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
