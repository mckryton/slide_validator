VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mActiveRules As Collection
Dim mSetupErrors As Collection

Public Property Get ActiveRules() As Collection

    If TypeName(mActiveRules) = "Nothing" Then
        Set mActiveRules = New Collection
    End If
    Set ActiveRules = mActiveRules
End Property

Public Property Get SetupErrors() As Collection

    If TypeName(mSetupErrors) = "Nothing" Then
        Set mSetupErrors = New Collection
    End If
    Set SetupErrors = mSetupErrors
End Property
