Attribute VB_Name = "TSpec"
Option Explicit

Public Const ERR_ID_EXPECTATION_FAILED = vbError + 6000

Public Function expect(pvarGivenValue As Variant) As TSpecExpectation

    Dim expectation As Variant      'expectation uses type variant instead of TSpecExpectation
                                    '   because VBA will modify error description for errors
                                    '   raised from explicitly typed objects, for more details see
                                    ' https://stackoverflow.com/questions/31234805/err-raise-is-ignoring-custom-description-and-source
    
    Set expectation = New TSpecExpectation
    expectation.given_value = pvarGivenValue
    Set expect = expectation
End Function
