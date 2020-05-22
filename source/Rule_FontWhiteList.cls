VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rule_FontWhiteList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mWhiteList As Collection
Dim mConfig As Collection


Private Sub Class_Terminate()
    Set mWhiteList = Nothing
    Set mConfig = Nothing
End Sub

Private Property Get white_list() As Collection

    Dim fonts As Variant
    Dim font As String
    
    If TypeName(mWhiteList) = "Nothing" Then
        Set mWhiteList = create_white_list(Me.Config("AllowedFonts"))
    End If
    Set white_list = mWhiteList
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

    If TypeName(mConfig) = "Nothing" Then
        Set mConfig = New Collection
        mConfig.Add "", "AllowedFonts"
    End If
    Set Config = mConfig
End Property

Public Property Let Config(ByVal pConfig As Collection)
    Set mConfig = pConfig
    'a new config will reset the white list
    Set mWhiteList = Nothing
End Property
