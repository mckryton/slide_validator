Attribute VB_Name = "TAcceptance"
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
Public Sub run_all_tests()

    Dim all_features As Variant
    Dim feature As Variant

    On Error GoTo error_handler
    all_features = Array(New Feature_FontWhiteList)
    For Each feature In all_features
        feature.test_scenarios
    Next
    Exit Sub

error_handler:
    SystemLogger.log_error "TAcceptance.run_all_tests"
End Sub


