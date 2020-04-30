Attribute VB_Name = "TAcceptance"
'------------------------------------------------------------------------
' Description  : starting point for running test sets for acceptance tests
'------------------------------------------------------------------------
Option Explicit


Public Sub run_all_tests(Optional pTags)
    
    Dim case_runner As TCaseRunner
    Dim acceptance_testcases As Variant

    On Error GoTo error_handler
    acceptance_testcases = Array(New Feature_FontWhiteList)
    Set case_runner = New TCaseRunner
    case_runner.run_testcases acceptance_testcases, pTags
    Exit Sub

error_handler:
    SystemLogger.log_error "TAcceptance.run_all_tests"
End Sub
