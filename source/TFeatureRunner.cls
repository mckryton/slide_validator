VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TFeatureRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub run_tesTFeatures(pTesTFeatures As Variant, Optional pTags)
    
    Dim all_tesTFeatures As Variant
    Dim tesTFeature As Variant
    Dim run_test As Boolean

    For Each tesTFeature In pTesTFeatures
        run_test = True
        If Not IsMissing(pTags) Then
            run_test = test_has_tag(pTags, tesTFeature)
        End If
        If run_test Then
            Debug.Print ""
            Debug.Print "Test case: " & TypeName(tesTFeature)
            Debug.Print vbTab & tesTFeature.description & vbLf
            tesTFeature.test_examples
            Set tesTFeature = Nothing
        End If
    Next
End Sub

Private Function test_has_tag(pTags As Variant, pTesTFeature As Variant)

    Dim input_tags As Variant
    Dim feature_tags As Variant
    Dim match As Variant
    Dim tag As Variant
    
    test_has_tag = False
    input_tags = Split(Replace(pTags, " ", ""), ",")
    feature_tags = Split(Replace(pTesTFeature.Tags, " ", ""), ",")
    For Each tag In input_tags
        match = Filter(feature_tags, tag, True, vbBinaryCompare)
        If UBound(match) > -1 Then
            test_has_tag = True
        End If
    Next

End Function

Public Property Get Log() As logger
    
    Set Log = TFeature.Log
End Property
