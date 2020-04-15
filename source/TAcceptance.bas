Attribute VB_Name = "TAcceptance"
'------------------------------------------------------------------------
' Description  : starting point for running test sets for acceptance tests
'------------------------------------------------------------------------
Option Explicit

Public Sub run_all_tests()

    Dim all_features As Variant
    Dim feature As Variant

    On Error GoTo error_handler
    all_features = Array(New Feature_FontWhiteList)
    For Each feature In all_features
        Debug.Print "Feature: " & TypeName(feature)
        Debug.Print vbTab & feature.description & vbLf
        feature.test_scenarios
        Set feature = Nothing
    Next
    Exit Sub

error_handler:
    SystemLogger.log_error "TAcceptance.run_all_tests"
End Sub


