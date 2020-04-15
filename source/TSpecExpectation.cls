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
        Err.Raise Number:=TSpec.ERR_ID_EXPECTATION_FAILED, description:=matcher_fail_msg
    End If
End Sub

Public Property Get given_value() As Variant
    given_value = mvarGivenValue
End Property

Public Property Let given_value(ByVal pvarGivenValue As Variant)
    mvarGivenValue = pvarGivenValue
End Property
