VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectValidationTarget 
   Caption         =   "Select a presentation for validation"
   ClientHeight    =   3405
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   6020
   OleObjectBlob   =   "SelectValidationTarget.frx":0000
   StartUpPosition =   1  'CenterOwner
End
 VB_Name = "SelectValidationTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWasCanceled As Boolean
Dim mPresentationsInfo As Collection

Public Property Get WasCanceled() As Boolean
    WasCanceled = mWasCanceled
End Property

Public Property Let WasCanceled(ByVal pWasCanceled As Boolean)
    mWasCanceled = pWasCanceled
End Property

Private Sub cmdCancel_Click()
    Me.WasCanceled = True
    Me.Hide
End Sub

Private Sub cmdValidate_Click()
    Me.Hide
End Sub

Private Sub lstPresentations_Click()
    Me.cmdValidate.Enabled = True
    If Me.PresentationsInfo(Me.lstPresentations.Value)(1) = "" Then
        Me.lblSelectedFile.Caption = "unsaved presentation"
    Else
        Me.lblSelectedFile.Caption = Me.PresentationsInfo(Me.lstPresentations.Value)(1)
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.WasCanceled = False
End Sub

Public Property Get PresentationsInfo() As Collection
    If TypeName(mPresentationsInfo) = "Nothing" Then
        Set PresentationsInfo = Nothing
    Else
        Set PresentationsInfo = mPresentationsInfo
    End If
End Property

Public Property Let PresentationsInfo(ByVal pPresentationInfo As Collection)
    
    Dim presentation_info As Variant
    
    Set mPresentationsInfo = pPresentationInfo
    Me.lstPresentations.Clear
    For Each presentation_info In mPresentationsInfo
        Me.lstPresentations.AddItem presentation_info(0)
    Next
End Property
Attribute VB_Name = "SelectValidationTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWasCanceled As Boolean
Dim mPresentationsInfo As Collection

Public Property Get WasCanceled() As Boolean
    WasCanceled = mWasCanceled
End Property

Public Property Let WasCanceled(ByVal pWasCanceled As Boolean)
    mWasCanceled = pWasCanceled
End Property

Private Sub cmdCancel_Click()
    Me.WasCanceled = True
    Me.Hide
End Sub

Private Sub cmdValidate_Click()
    Me.Hide
End Sub

Private Sub lstPresentations_Click()
    Me.cmdValidate.Enabled = True
    If Me.PresentationsInfo(Me.lstPresentations.Value)(1) = "" Then
        Me.lblSelectedFile.Caption = "unsaved presentation"
    Else
        Me.lblSelectedFile.Caption = Me.PresentationsInfo(Me.lstPresentations.Value)(1)
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.WasCanceled = False
End Sub

Public Property Get PresentationsInfo() As Collection
    If TypeName(mPresentationsInfo) = "Nothing" Then
        Set PresentationsInfo = Nothing
    Else
        Set PresentationsInfo = mPresentationsInfo
    End If
End Property

Public Property Let PresentationsInfo(ByVal pPresentationInfo As Collection)
    
    Dim presentation_info As Variant
    
    Set mPresentationsInfo = pPresentationInfo
    Me.lstPresentations.Clear
    For Each presentation_info In mPresentationsInfo
        Me.lstPresentations.AddItem presentation_info(0)
    Next
End Property
