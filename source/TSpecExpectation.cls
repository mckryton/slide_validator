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

Public Function to_be(pvarExpectedValue) As Boolean

    If Me.given_value = pvarExpectedValue Then
        to_be = True
    Else
        Debug.Print vbTab & "missed expectation"
        Debug.Print "", "found:", ">" & Me.given_value & "<"
        Debug.Print "", "expected:", ">" & pvarExpectedValue & "<"
        to_be = True = False
    End If
End Function

Public Property Get given_value() As Variant
    given_value = mvarGivenValue
End Property

Public Property Let given_value(ByVal pvarGivenValue As Variant)
    mvarGivenValue = pvarGivenValue
End Property
