Attribute VB_Name = "basSystemLogger"
'------------------------------------------------------------------------
' Description  : contains all global constants
'------------------------------------------------------------------------

'Options
Option Explicit

'log level (range is 1 to 100)
Global Const cLogDebug = 100
Global Const cLogInfo = 90
Global Const cLogWarning = 50
Global Const cLogError = 30
Global Const cLogCritical = 1

'current log level - decreasing log level means decreasing amount of messages
Global Const cCurrentLogLevel = 100

'-------------------------------------------------------------
' Description   : prints log messages to direct window
' Parameter     :   pstrLogMsg      - log message
'                   pintLogLevel    - log level for this message
'-------------------------------------------------------------
Public Sub log(pstrLogMsg As String, Optional pintLogLevel)

    Dim intLogLevel As Integer      'aktueller Loglevel
    Dim strLog As String            'auszugebender Text
    
    'default log level is cLogInfo
    If IsMissing(pintLogLevel) Then
        intLogLevel = cLogInfo
    Else
        intLogLevel = pintLogLevel
    End If
   
    'print log message only if given log level is lower or equal then
    ' log level set in module basConstants
    If intLogLevel <= cCurrentLogLevel Then
        'start with current time
        strLog = Time
        'add log level
        Select Case intLogLevel
            Case cLogDebug
                strLog = strLog & " debug:"
            Case cLogInfo
                strLog = strLog & " info:"
            Case cLogWarning
                strLog = strLog & " warning:"
            Case cLogError
                strLog = strLog & " error:"
            Case cLogCritical
                strLog = strLog & " critical:"
            Case Else
                strLog = strLog & " custom(" & intLogLevel & "):"
        End Select
        'add log message
        strLog = strLog & " " & pstrLogMsg
        Debug.Print strLog
    End If
End Sub
'-------------------------------------------------------------
' Description   : function print error messages to the direct window
' Parameter     : pstrFunctionName  - name of the calling function
'                 pstrLogMsg        - optional: custom error message
'-------------------------------------------------------------
Public Sub log_error(pstrFunctionName As String, Optional pstrLogMsg As Variant)

    Dim intLogLevel As Integer      'current log level
    Dim strLog As String            'complete log messages
    Dim strError As String          'system error message from Err object
    
    strError = Err.Description
    'start log messages with time
    strLog = Time
    'log level = error
    strLog = strLog & " error:"
    'add caller name
    strLog = strLog & "error in " & pstrFunctionName & ": "
    'if given add custom log message
    If Not IsMissing(pstrLogMsg) Then
        strLog = strLog & " " & pstrLogMsg
    Else
        'use message from Err object
        On Error Resume Next
        strLog = strLog & " " & strError
    End If
    Debug.Print strLog
End Sub
'-------------------------------------------------------------
' Description   : alias to log function with cLogDebug level
' Parameter     :   pstrLogMsg      - log message
'-------------------------------------------------------------
Public Sub logd(pstrLogMsg As String)

    On Error GoTo error_handler
    basSystemLogger.log pstrLogMsg, cLogDebug
    Exit Sub

error_handler:
    basSystemLogger.log_error "basSystem.logd"
End Sub
