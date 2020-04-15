Attribute VB_Name = "TSpec"
Option Explicit

Public Const ERR_ID_EXPECTATION_FAILED = vbError + 6000

Public Function expect(pvarGivenValue As Variant) As Variant

    Dim expectation As Variant
    
    Set expectation = New TSpecExpectation
    expectation.given_value = pvarGivenValue
    Set expect = expectation
End Function
