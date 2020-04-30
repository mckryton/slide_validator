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
    
    If Me.given_value <> pvarExpectedValue Then
        matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                            vbTab & "found:" & vbTab & vbTab & ">" & Me.given_value & "<" & vbLf & _
                            vbTab & "expected:" & vbTab & ">" & pvarExpectedValue & "<"
        raise_expectation_err matcher_fail_msg
    End If
End Sub

Public Property Get given_value() As Variant
    If TypeName(mvarGivenValue) = "Collection" Or TypeName(mvarGivenValue) = "Empty" Then
        Set given_value = mvarGivenValue
    Else
        given_value = mvarGivenValue
    End If
End Property

Public Property Let given_value(ByVal pvarGivenValue As Variant)
    If TypeName(mvarGivenValue) = "Collection" Or TypeName(mvarGivenValue) = "Empty" Then
        Set mvarGivenValue = pvarGivenValue
    Else
        mvarGivenValue = pvarGivenValue
    End If
End Property

Public Sub contains(pvarExpectedItem As Variant)

    Dim matcher_fail_msg As String
    
    Select Case TypeName(Me.given_value)
    Case "Collection"
        If Not System.existsItem(pvarExpectedItem(0), Me.given_value) Then
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
        TScenarioRunner.pending "contains matcher for array"
    End Select
End Sub

Private Sub raise_expectation_err(pFailMsg)
    
    Err.Raise Number:=TSpec.ERR_ID_EXPECTATION_FAILED, description:=pFailMsg
End Sub
