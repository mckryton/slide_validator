Attribute VB_Name = "ExtraVBA"
Option Explicit

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

Private Sub exportCode()

    Dim vcomSource As VBComponent
    Dim strPath As String
    Dim strSeparator As String
    Dim strSuffix As String
    Dim export_logger As Logger

    On Error GoTo error_handler
    Set export_logger = New Logger
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
        export_logger.log "export code to " & strPath & strSeparator & vcomSource.Name & "." & strSuffix
    Next
    Exit Sub

error_handler:
    export_logger.log_function_error "System.exportCode"
End Sub
