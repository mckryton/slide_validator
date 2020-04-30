Attribute VB_Name = "TUnit"
'------------------------------------------------------------------------
' Description  : starting point for running test sets for unit tests
'------------------------------------------------------------------------
Option Explicit

Public Sub run_all_tests(Optional pTags)
    
    Dim case_runner As TCaseRunner
    Dim unit_testcases As Variant

    'On Error GoTo error_handler
    unit_testcases = Array(New Unit_ReadConfig, New Unit_ChooseTarget)
    Set case_runner = New TCaseRunner
    case_runner.run_testcases unit_testcases, pTags
    Exit Sub

error_handler:
    SystemLogger.log_error "TUnit.run_all_tests"
End Sub

