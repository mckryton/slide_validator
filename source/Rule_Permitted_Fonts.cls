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

Private Property Get white_list() As Collection

    Dim fonts As Variant
    Dim font As String
    
    If TypeName(m_permitted_fonts) = "Nothing" Then
        Set m_permitted_fonts = create_white_list(Me.Config("PermittedFonts"))
    End If
    Set white_list = m_permitted_fonts
End Property

Public Function apply_rule(psldCurrentSlide As Slide) As String

    Dim shape As shape
    Dim violations As String
    
    violations = ""
    For Each shape In psldCurrentSlide.Shapes
        If shape.TextFrame.HasText Then
            If Not ExtraVBA.existsItem(shape.TextFrame2.TextRange.font.Name, white_list) Then
                violations = violations & "illegal font >" & shape.TextFrame2.TextRange.font.Name & "< in shape " & shape.Name & vbCrLf
            End If
        End If
    Next
    apply_rule = violations
End Function

Private Function create_white_list(pWhiteList_Config_String As String) As Collection
    
    Dim white_list_config As Variant
    Dim font_name As Variant
    Dim new_white_list As Collection
    
    Set new_white_list = New Collection
    white_list_config = Split(pWhiteList_Config_String, ",")
    For Each font_name In white_list_config
        new_white_list.Add Trim(font_name), Trim(font_name)
    Next
    Set create_white_list = new_white_list
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
