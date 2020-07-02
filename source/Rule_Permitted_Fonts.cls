VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rule_Permitted_Fonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_permitted_fonts As Collection
Dim m_rule_config As Collection


Private Sub Class_Terminate()
    Set m_permitted_fonts = Nothing
    Set m_rule_config = Nothing
End Sub

Public Property Get permitted_fonts() As Collection

    Dim fonts As Variant
    Dim font As String
    
    If TypeName(m_permitted_fonts) = "Nothing" Then
        Set m_permitted_fonts = collect_permitted_fonts(Me.Config("PermittedFonts"))
    End If
    Set permitted_fonts = m_permitted_fonts
End Property

Public Function apply_rule(psldCurrentSlide As Slide) As String

    Dim shape As shape
    Dim violations As String
    
    violations = ""
    For Each shape In psldCurrentSlide.Shapes
        If shape.TextFrame.HasText Then
            If shape.TextFrame.TextRange.font.Name = "" Then
                violations = add_violation_msg(violations, "multiple fonts in shape " & shape.Name & " detected")
            ElseIf Not ExtraVBA.existsItem(shape.TextFrame2.TextRange.font.Name, Me.permitted_fonts) Then
                violations = add_violation_msg(violations, "illegal font >" & shape.TextFrame2.TextRange.font.Name & "< in shape " & shape.Name)
            End If
        End If
    Next
    apply_rule = violations
End Function

Private Function add_violation_msg(violations As String, new_msg As String) As String
    
    If violations <> "" Then
        add_violation_msg = violations & vbLf & new_msg
    Else
        add_violation_msg = new_msg
    End If
End Function

Private Function collect_permitted_fonts(permitted_fonts_config As String) As Collection
    
    Dim permitted_fonts_config_values As Variant
    Dim font_name As Variant
    Dim permitted_fonts As Collection
    
    Set permitted_fonts = New Collection
    permitted_fonts_config_values = Split(permitted_fonts_config, ",")
    For Each font_name In permitted_fonts_config_values
        permitted_fonts.Add Trim(font_name), Trim(font_name)
    Next
    Set collect_permitted_fonts = permitted_fonts
End Function


Public Property Get Config() As Collection

    If TypeName(m_rule_config) = "Nothing" Then
        Set m_rule_config = New Collection
        m_rule_config.Add "", "PermittedFonts"
    End If
    Set Config = m_rule_config
End Property

Public Property Let Config(ByVal pConfig As Collection)
    Set m_rule_config = pConfig
    'a new config will reset the white list
    Set m_permitted_fonts = Nothing
End Property
