Attribute VB_Name = "TestStart"
Option Explicit

Public Sub run_tests(Optional pTags)
    
    Dim feature_runner As TFeatureRunner
    Dim tesTFeatures As Variant
    Dim Log As logger

    Set Log = New logger
    tesTFeatures = Array(New Feature_ApplyRules, New Feature_Rule_Permitted_Fonts)
    Set feature_runner = New TFeatureRunner
    feature_runner.run_tesTFeatures tesTFeatures, pTags
End Sub

Public Sub run_acceptance_tests()
    TestStart.run_tests "feature"
End Sub

Public Sub run_acceptance_wip_tests()
    'wip = work in progress
    TestStart.run_tests "feature,wip"
End Sub

Public Sub run_unit_tests()
    TestStart.run_tests "unit"
End Sub

Public Sub run_unit_wip_tests()
    'wip = work in progress
    TestStart.run_tests "unit,wip"
End Sub

