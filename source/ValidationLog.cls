VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mViolations As Collection

Public Property Get violations() As Collection

    If TypeName(mViolations) = "Nothing" Then
        Set mViolations = New Collection
    End If
    Set violations = mViolations
End Property

