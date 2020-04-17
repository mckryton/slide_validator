Attribute VB_Name = "TUnit"
'------------------------------------------------------------------------
' Description  : starting point for running test sets for unit tests
'------------------------------------------------------------------------
Option Explicit

Public Sub run_all_tests()
    
    Dim all_testcases As Variant
    Dim testcase As Variant

    On Error GoTo error_handler
    all_testcases = Array(New Unit_ReadConfig)
    For Each testcase In all_testcases
        Debug.Print "Test case: " & TypeName(testcase)
        Debug.Print vbTab & testcase.description & vbLf
        testcase.test_scenarios
        Set testcase = Nothing
    Next
    Exit Sub

error_handler:
    SystemLogger.log_error "TUnit.run_all_tests"
End Sub

Public Sub pending(pPendingMsg)
    
    Debug.Print "PENDING: " & pPendingMsg
End Sub
