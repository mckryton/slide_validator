VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TSpecExpectation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mvarGivenValue As Variant

Public Sub to_be(pvarExpectedValue)

    Dim matcher_fail_msg As String
    
    If (IsObject(Me.given_value) And Not IsObject(pvarExpectedValue)) _
     Or (IsObject(pvarExpectedValue) And Not IsObject(Me.given_value)) Then
        matcher_fail_msg = vbTab & "can't compare given with expected value - check if only one of them is an object"
        raise_expectation_err matcher_fail_msg
    End If
    
    If IsObject(Me.given_value) Then
        If Not (Me.given_value Is pvarExpectedValue) Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                                vbTab & "can't match objects"
            raise_expectation_err matcher_fail_msg
        End If
    Else
        If Me.given_value <> pvarExpectedValue Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                                vbTab & "found:" & vbTab & vbTab & ">" & Me.given_value & "<" & vbLf & _
                                vbTab & "expected:" & vbTab & ">" & pvarExpectedValue & "<"
            raise_expectation_err matcher_fail_msg
        End If
    End If
End Sub

Public Property Get given_value() As Variant
    On Error Resume Next
    Set given_value = mvarGivenValue
    given_value = mvarGivenValue
    On Error GoTo 0
End Property

Public Property Let given_value(ByVal pvarGivenValue As Variant)
    On Error Resume Next
    Set mvarGivenValue = pvarGivenValue
    mvarGivenValue = pvarGivenValue
    On Error GoTo 0
End Property

Public Sub contains(pvarExpectedItem As Variant)

    Dim matcher_fail_msg As String
    
    Select Case TypeName(Me.given_value)
    Case "Collection"
        If Not ExtraVBA.existsItem(pvarExpectedItem(0), Me.given_value) Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                        vbTab & "collection does not contain item " & ">" & pvarExpectedItem(0) & "<"
            raise_expectation_err matcher_fail_msg
        ElseIf Me.given_value(pvarExpectedItem(0)) <> pvarExpectedItem(1) Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                        vbTab & "found:" & vbTab & vbTab & ">" & Me.given_value(pvarExpectedItem(0)) & "<" & vbLf & _
                        vbTab & "expected:" & vbTab & ">" & pvarExpectedItem(1) & "<"
            raise_expectation_err matcher_fail_msg
        End If
    Case "Array"
        TExampleRunner.pending "contains matcher for array"
    End Select
End Sub

Private Sub raise_expectation_err(pFailMsg)
    
    Err.Raise TSpec.ERR_ID_EXPECTATION_FAILED, pFailMsg, pFailMsg
    Debug.Print ""
    'Err.Raise Number:=TSpec.ERR_ID_EXPECTATION_FAILED, Source:=pFailMsg, description:=pFailMsg
End Sub
