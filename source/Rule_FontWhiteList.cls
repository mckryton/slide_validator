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

Dim mcolFontWhiteList As Collection

Private Sub Class_Initialize()
    Set mcolFontWhiteList = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolFontWhiteList = Nothing
End Sub

Public Property Get font_white_list() As Collection
    Set font_white_list = mcolFontWhiteList
End Property

Public Function apply_rule(psldCurrentSlide As Slide) As String

    Dim shape As shape
    Dim violations As String
    
    violations = ""
    For Each shape In psldCurrentSlide.Shapes
        If shape.TextFrame.HasText Then
            If Not System.existsItem(shape.TextFrame2.TextRange.Font.Name, Me.font_white_list) Then
                violations = violations & "illegal font >" & shape.TextFrame2.TextRange.Font.Name & "< in shape " & shape.Name & vbCrLf
            End If
        End If
    Next
    apply_rule = violations
End Function

Public Sub create_white_list(pFontWhiteList_Config_String As String)
    
    Dim white_list_config As Variant
    Dim font_name As Variant
    
    Me.clear_white_list
    white_list_config = Split(pFontWhiteList_Config_String, ",")
    For Each font_name In white_list_config
        Me.font_white_list.Add Trim(font_name), Trim(font_name)
    Next
End Sub

Public Sub clear_white_list()
    
    Do While Me.font_white_list.Count > 0
        Me.font_white_list.Remove 0
    Loop
End Sub
