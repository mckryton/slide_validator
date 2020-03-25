Attribute VB_Name = "basTestAcceptance"
'------------------------------------------------------------------------
' Description  : test all features
'------------------------------------------------------------------------
'
'Declarations


'Declare variables

'Options
Option Explicit
'-------------------------------------------------------------
' Description   : run all acceptance tests
' Parameter     :
'-------------------------------------------------------------
Public Sub runAllTests()

    Dim objRuleValidator As Object

    On Error GoTo error_handler
    Set objRuleValidator = New clsFeatureFontWhiteList
    objRuleValidator.testScenarios

    Exit Sub

error_handler:
    basSystemLogger.log_error "basTestAcceptance.runAllTests"
End Sub


