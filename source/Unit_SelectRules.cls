VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Unit_SelectRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' Description  : Ability to find configs and apply them to rule functions
'------------------------------------------------------------------------
Option Explicit

'separate tags by comma, e.g.: "@debug,@production"
Const cTags = ""

Dim mLogger As Logger
Dim mTestPresentation As Presentation
Dim mTestConfigSlide As Slide
Dim mRuleConfig As Collection
Dim mIsConfigSlide As Boolean

Public Property Get description() As String

    description = "Slide Validator will find configuration for rule from it's own slides" & vbLf & _
                    vbTab & "and apply configuration found to matching rules"
End Property

Public Sub test_examples()

    Dim examples As Variant
    Dim example As Variant

    On Error GoTo error_handler
    examples = Array( _
                    Array("Rule: A slide is a configuration slide if it's layout is named rule_config"), _
                    Array("Example: valid configuration slide", _
                           "Given a config slide using the master slide template for rule config", _
                           "When the slide is validated to be a configuration", _
                           "Then the slide is recognized as a configuration" _
                    ), _
                    Array("Rule: Ignore hidden config slides"), _
                    Array("Rule: Choose Rule from config slide title"), _
                    Array("Example: config slide has a matching rule", _
                            "Given a SlideValidator having a config slide titled FontWhiteList", _
                            "And a matching Rule named Rule_FontWhiteList", _
                            "And a target presentation using Times New Roman for a text", _
                            "When the target presentation is validated", _
                            "Then the rule is applied to the target presentation" _
                    ) _
                )
    For Each example In examples
        TExampleRunner.run_example example, Me
    Next
    Exit Sub

error_handler:
    SystemLogger.log_error "Unit_ReadConfig.test_scenarios"
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
        
    Dim step_result As String
    Dim config_table As Table
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
        
        Case "Given a config slide using the master slide template for rule config"
            Set mTestPresentation = TSupport.create_slide_validator_pres
            Set mTestConfigSlide = TSupport.create_config_slide(mTestPresentation)
            
        Case "When the slide is validated to be a configuration"
            mIsConfigSlide = Validator.is_config_slide(mTestConfigSlide)
        
        Case Else
            Err.raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = raise_step_error(Err.Number)
End Function

Private Function raise_step_error(pErrId As Long) As Variant
 
    TExampleRunner.stop_test
    Select Case Err.Number
    Case ERR_ID_STEP_IS_PENDING
        raise_step_error = Array("PENDING")
    Case ERR_ID_STEP_IS_MISSING
        raise_step_error = Array("MISSING")
    Case Else
        raise_step_error = Array("FAILED", Err.description)
    End Select
End Function
Public Sub after()
    If TypeName(mTestPresentation) <> "Nothing" Then
        mTestPresentation.Saved = msoTrue
        mTestPresentation.Close
    End If
End Sub

Public Property Get Tags() As Variant
    Tags = cTags
End Property

Private Sub pending(pPendingMsg)
    Err.raise ERR_ID_STEP_IS_PENDING, description = pPendingMsg
End Sub

Private Property Get Log() As Logger
    
    If TypeName(mLogger) = "Nothing" Then
        Set mLogger = New Logger
    End If
    Set Log = mLogger
End Property

Public Property Let Log(ByVal pLogger As Logger)
    Set mLogger = pLogger
End Property

Private Sub Class_Terminate()
    Set mLogger = Nothing
End Sub
