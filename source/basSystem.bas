Attribute VB_Name = "basSystem"
'------------------------------------------------------------------------
' Description  : extends system related functions
'------------------------------------------------------------------------
'
'Declarations

'Declare variables

'Options
Option Explicit
'-------------------------------------------------------------
' Description   : checks if item exists in a collection object
' Parameter     : pvarKey           - item name
'                 pcolACollection   - collection object
' Returnvalue   : true if item exits, false if not
'-------------------------------------------------------------
Public Function existsItem(pvarKey As Variant, pcolACollection As Collection) As Boolean
                    
    Dim varItemValue As Variant
                     
    On Error GoTo NOT_FOUND
    varItemValue = pcolACollection.Item(pvarKey)
    On Error GoTo 0
    existsItem = True
    Exit Function
                     
NOT_FOUND:
    existsItem = False
End Function

'-------------------------------------------------------------
' Description   : save source code as text files
'-------------------------------------------------------------
Private Sub exportCode()

    Dim vcomSource As VBComponent
    Dim strPath As String
    Dim strSeparator As String
    Dim strSuffix As String

    On Error GoTo error_handler
    #If MAC_OFFICE_VERSION >= 15 Then
        'in Office 2016 MAC M$ switched to / as path separator
        strSeparator = "/"
    #ElseIf Mac Then
        strSeparator = ":"
    #Else
        strSeparator = "\"
    #End If
    strPath = ActivePresentation.Path & strSeparator & "source"
    For Each vcomSource In Application.VBE.VBProjects("SlideValidator").VBComponents
        Select Case vcomSource.Type
            Case vbext_ct_StdModule
                strSuffix = "bas"
            Case vbext_ct_ClassModule
                strSuffix = "cls"
            Case vbext_ct_Document
                strSuffix = "cls"
            Case vbext_ct_MSForm
                strSuffix = "frm"
            Case Else
                strSuffix = "txt"
        End Select
        vcomSource.Export strPath & strSeparator & vcomSource.Name & "." & strSuffix
        basSystemLogger.log "export code to " & strPath & strSeparator & vcomSource.Name & "." & strSuffix
    Next
    Exit Sub

error_handler:
    basSystemLogger.log_error "basSystem.exportCode"
End Sub
