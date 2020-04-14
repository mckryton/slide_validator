VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRuleFontWhiteList"
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
    apply_rule = "PENDING rule validation result for font white list"
End Function
