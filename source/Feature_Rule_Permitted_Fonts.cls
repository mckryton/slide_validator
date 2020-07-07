VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feature_Rule_Permitted_Fonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'separate tags by comma, e.g.: "debug,production"
Const cTags = ""

Dim m_logger As Logger
Dim m_rule_permitted_fonts As Rule_Permitted_Fonts
Dim m_test_presentation As Presentation
Dim m_test_slide As Slide


Public Property Get description() As String

    description = "The presentation should only use permitted fonts listed in the configuration"
End Property

Public Sub test_examples()

    Dim examples As Collection
    Dim example As Variant

    Set examples = New Collection
    examples.Add Array("Rule: add a validation comment for every shape using a font not permitted")

    examples.Add Array("Example: slide with a valid font", _
                           "Given a new presentation", _
                           "And it contains a slide with a textbox using ""Arial""", _
                           "And Arial is the only permitted font", _
                           "When the ""Permitted_Fonts"" rule is applied to the slide", _
                           "Then no comment is added to the slide")
    
    examples.Add Array("Example: slide with an invalid font", _
                           "Given a new presentation", _
                           "And it contains a slide with a textbox named ""fonttest_1"" using ""Times New Roman"" as font", _
                           "And Arial is the only permitted font", _
                           "When the ""Permitted_Fonts"" rule is applied to the slide", _
                           "Then a comment ""illegal font >Times New Roman< in shape fonttest_1"" is added to the slide")
    
    examples.Add Array("Example: a shape is using multiple fonts", _
                           "Given Arial is the only permitted font", _
                           "Given a new presentation", _
                           "And a slide with a textbox named ""fonttest_1"" using ""Arial"" for the first word and ""Courier New"" for the second word", _
                           "And Arial is the only permitted font", _
                           "When the ""Permitted_Fonts"" rule is applied to the slide", _
                           "Then a comment ""multiple fonts in shape fonttest_1 detected"" is added to the slide")
    For Each example In examples
        TExampleRunner.run_example example, Me
    Next
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
    
    Dim step_result As String
    Dim rule_config As Collection
    Dim textbox As shape
    Dim rules As Collection
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
        Case "Given Arial is the only permitted font"
            Set m_rule_permitted_fonts = New Rule_Permitted_Fonts
            Set rule_config = New Collection
            rule_config.Add "Arial", "PermittedFonts"
            m_rule_permitted_fonts.Config = rule_config
            
        Case "Given a new presentation"
            pending "test pending"
            Set m_test_presentation = Application.Presentations.Add
            
        Case "Given it contains a slide with a textbox using ""Arial"""
            Set m_test_slide = TSupport.add_slide_with_textbox(m_test_presentation, "Arial")
            
        Case "Given it contains a slide with a textbox named ""fonttest_1"" using ""Times New Roman"" as font"
            Set m_test_slide = TSupport.add_slide_with_textbox(m_test_presentation, "Times New Roman", "fonttest_1")
            
        Case "Given a slide with a textbox named ""fonttest_1"" using ""Arial"" for the first word and ""Courier New"" for the second word"
            Set m_test_slide = TSupport.add_empty_slide(m_test_presentation)
            Set textbox = m_test_slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 200, 200, 400, 200)
            textbox.TextFrame.TextRange.Text = "Arial Helvetica"
            textbox.TextFrame.TextRange.Words(1, 1).font.Name = "Arial"
            textbox.TextFrame.TextRange.Words(2, 1).font.Name = "Courier New"
            textbox.Name = "fonttest_1"
            
        Case "When the ""Permitted_Fonts"" rule is applied to the slide"
            Set rules = New Collection
            rules.Add m_rule_permitted_fonts
            Validator.apply_rules_on_slide rules, m_test_slide
            
        Case "Then no comment is added to the slide"
            TSpec.expect(m_test_presentation.Slides(1).Comments.Count).to_be 0
        
        Case "Then a comment about the rule validation is added to the slide"
            TSpec.expect(m_test_presentation.Slides(1).Comments.Count).to_be 1
        
        Case "Then a comment ""illegal font >Times New Roman< in shape fonttest_1"" is added to the slide"
            TSpec.expect(m_test_presentation.Slides(1).Comments.Count).to_be 1
            TSpec.expect(m_test_presentation.Slides(1).Comments(1).Text).to_be "illegal font >Times New Roman< in shape fonttest_1"
            
        Case "Then a comment ""multiple fonts in shape fonttest_1 detected"" is added to the slide"
            TSpec.expect(m_test_presentation.Slides(1).Comments.Count).to_be 1
            TSpec.expect(m_test_presentation.Slides(1).Comments(1).Text).to_be "multiple fonts in shape fonttest_1 detected"
            
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
        raise_step_error = Array("PENDING", Err.description)
    Case ERR_ID_STEP_IS_MISSING
        raise_step_error = Array("MISSING")
    Case Else
        raise_step_error = Array("FAILED", Err.description)
    End Select
End Function

Public Sub before()
    
    Validator.Log.LogLevel = Validator.Log.LVL_WARNING
End Sub

Public Sub after()
    m_test_presentation.Saved = msoTrue
    m_test_presentation.Close
    Set m_test_presentation = Nothing
End Sub

Public Property Get Tags() As Variant
    Tags = cTags
End Property

Private Sub pending(pPendingMsg)
    Err.raise ERR_ID_STEP_IS_PENDING, description:=pPendingMsg
End Sub

Private Property Get Log() As Logger
    
    If TypeName(m_logger) = "Nothing" Then
        Set m_logger = New Logger
    End If
    Set Log = m_logger
End Property

Public Property Let Log(ByVal pLogger As Logger)
    Set m_logger = pLogger
End Property

Private Sub Class_Terminate()
    Set m_logger = Nothing
End Sub

