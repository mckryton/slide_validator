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

'Options
Option Explicit
'-------------------------------------------------------------
' Description   : run all acceptance tests
' Parameter     :
'-------------------------------------------------------------
Public Sub testScenarios()

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
    basTestRunner.runScenario scenario, Me
    
    Exit Sub

error_handler:
    basSystemLogger.log_error "clsFeatureFontWhiteList.runTest"
End Sub
'-------------------------------------------------------------
' Description   : execute a single test step
' Parameter     : pcolStepProps     - all properties of the current step, e.g. type, head, body, line
'-------------------------------------------------------------
Public Sub runStep(pcolStepProps As Collection)

    Dim strStepType As String
    
    On Error GoTo error_handler
    
    Select Case pcolStepProps.Item("step_type") & " " & pcolStepProps.Item("line_body")
        Case "Given a font white list rule listing Arial"
            Debug.Print "pending"
            
        Case Else
            basTestRunner.missingTest pcolStepProps.Item("step_type") & " " & Trim(pcolStepProps.Item("line_body")), Me
    End Select
    
    Exit Sub

error_handler:
    basSystemLogger.log_error "clsFeatureFontWhiteList.runStep " & pcolStepProps.Item("step_type") & " " & pcolStepProps.Item("line_body")
End Sub
