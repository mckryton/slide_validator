VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TCaseRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub run_testcases(pTestCases As Variant, Optional pTags)
    
    Dim all_testcases As Variant
    Dim testcase As Variant
    Dim run_test As Boolean

    On Error GoTo error_handler
    For Each testcase In pTestCases
        run_test = True
        If Not IsMissing(pTags) Then
            run_test = test_has_tag(pTags, testcase)
        End If
        If run_test Then
            Debug.Print vbLf & vbLf
            Debug.Print "Test case: " & TypeName(testcase)
            Debug.Print vbTab & testcase.description & vbLf
            testcase.test_scenarios
            Set testcase = Nothing
        End If
    Next
    Exit Sub

error_handler:
    SystemLogger.log_error "TCaseRunner.run_testcases"
End Sub

Private Function test_has_tag(pTags As Variant, pTestCase As Variant)

    Dim input_tags As Variant
    Dim feature_tags As Variant
    Dim match As Variant
    Dim tag As Variant
    
    On Error GoTo error_handler
    test_has_tag = False
    input_tags = Split(Replace(pTags, " ", ""), ",")
    feature_tags = Split(Replace(pTestCase.tags, " ", ""), ",")
    For Each tag In input_tags
        match = Filter(feature_tags, tag, True, vbBinaryCompare)
        If UBound(match) > -1 Then
            test_has_tag = True
        End If
    Next
    Exit Function

error_handler:
    SystemLogger.log_error "TCaseRunner.run_testcases"
End Function