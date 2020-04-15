Attribute VB_Name = "SystemLogger"
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

Const cLineBreak = vbCr & vbLf

'current log level - decreasing log level means decreasing amount of messages
Global Const cCurrentLogLevel = 30

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

    Dim strLog As String            'complete log messages
    
    strLog = Time & " error:" & cLineBreak & _
                vbTab & "source:" & vbTab & Err.Source & cLineBreak & _
                vbTab & "caller:" & vbTab & pstrFunctionName & cLineBreak & _
                vbTab & "desc:" & vbTab & Err.Description & cLineBreak
    If Not IsMissing(pstrLogMsg) Then
        strLog = strLog & vbTab & "custom:" & vbTab & pstrLogMsg
    End If
    Debug.Print strLog
End Sub
'-------------------------------------------------------------
' Description   : alias to log function with cLogDebug level
' Parameter     :   pstrLogMsg      - log message
'-------------------------------------------------------------
Public Sub logd(pstrLogMsg As String)

    On Error GoTo error_handler
    SystemLogger.log pstrLogMsg, cLogDebug
    Exit Sub

error_handler:
    SystemLogger.log_error "System.logd"
End Sub
