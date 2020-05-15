Attribute VB_Name = "TestStart"
Option Explicit

Public Sub run_acceptance_tests(Optional pTags)
    
    Dim case_runner As TCaseRunner
    Dim acceptance_testcases As Variant
    Dim log As Logger


    On Error GoTo error_handler
    Set log = New Logger
    acceptance_testcases = Array(New Feature_FontWhiteList)
    Set case_runner = New TCaseRunner
    case_runner.run_testcases acceptance_testcases, pTags
    Exit Sub

error_handler:
    log.log_function_error "TestStart.run_acceptance_tests"
End Sub

Public Sub run_acceptance_wip_tests()
    'wip = work in progress
    TestStart.run_acceptance_tests "wip"
End Sub


Public Sub run_unit_tests(Optional pTags)
    
    Dim case_runner As TCaseRunner
    Dim unit_testcases As Variant
    Dim log As Logger

    On Error GoTo error_handler
    Set log = New Logger
    unit_testcases = Array(New Unit_ReadConfig, New Unit_ChooseTarget, New Unit_SelectRules)
    Set case_runner = New TCaseRunner
    case_runner.run_testcases unit_testcases, pTags
    Exit Sub

error_handler:
    log.log_function_error "TestStart.run_unit_tests"
End Sub


Public Sub run_unit_wip_tests()
    'wip = work in progress
    TestStart.run_unit_tests "wip"
End Sub

