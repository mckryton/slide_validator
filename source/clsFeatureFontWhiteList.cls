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
    basCuke.runScenario scenario, Me
    
    Exit Sub

error_handler:
    basSystemLogger.log_error "clsFeatureFontWhiteList.runTest"
End Sub
'-------------------------------------------------------------
' Description   : provide given pre-conditions
' Parameter     :
'-------------------------------------------------------------
Public Sub runGivenSteps(pstrGiven As String)

    Dim strGiven As String
    
    On Error GoTo error_handler
    strGiven = Replace(pstrGiven, "And", "Given", 1, 1)
    Select Case pstrGiven
    
    Case Else
        basCuke.missingTest pstrGiven, Me
    End Select
    
    Exit Sub

error_handler:
    basSystemLogger.log_error "clsFeatureFontWhiteList.runGivenSteps " & pstrGiven
End Sub
'-------------------------------------------------------------
' Description   : execute when statements
' Parameter     :
'-------------------------------------------------------------
Public Sub runWhenSteps(pstrWhen As String)

    On Error GoTo error_handler
    Select Case pstrWhen
    
    Case Else
        basCuke.missingTest pstrWhen, Me
    End Select
    
    Exit Sub

error_handler:
    basSystemLogger.log_error "clsFeatureFontWhiteList.runWhenSteps " & pstrWhen
End Sub
'-------------------------------------------------------------
' Description   : execute then statements
' Parameter     :
'-------------------------------------------------------------
Public Sub runThenSteps(pstrThen As String)

    On Error GoTo error_handler
    Select Case pstrThen
    
    Case Else
        basCuke.missingTest pstrThen, Me
    End Select
    
    Exit Sub

error_handler:
    basSystemLogger.log_error "clsFeatureFontWhiteList.runThenSteps " & pstrThen
End Sub

