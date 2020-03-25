Attribute VB_Name = "basSlideValidator"
'------------------------------------------------------------------------
' Description  : apply rules to active presentation, add violations as comment
'------------------------------------------------------------------------
'
'Declarations
Const mcRuleCheckAuthor = "Slide Validator"
Const mcRuleCheckInitials = "bot"

'Declare variables

'Options
Option Explicit
'-------------------------------------------------------------
' Description   : add rule violation message as ppt comment
' Parameter     : psldValidatedSlide    - the current validated slide
'                 pcolViolationMessages - feedback from the rule check
'-------------------------------------------------------------
Public Sub addViolation(psldValidatedSlide As Slide, pcolViolationMessages As Collection)
    
    Dim varViolationMsg As Variant
    Dim lngCommentPosX As Long
    
    On Error GoTo error_handler
    lngCommentPosX = 10
    For Each varViolationMsg In pcolViolationMessages
        psldValidatedSlide.Comments.Add 10, 10, mcRuleCheckAuthor, mcRuleCheckInitials, varViolationMsg
    Next
    lngCommentPosX = lngCommentPosX + 10
    Exit Sub

error_handler:
    basSystemLogger.log_error "basSlideValidator.addViolation"
End Sub
'-------------------------------------------------------------
' Description   : apply rules to slides of the active presentation
' Parameter     :
'-------------------------------------------------------------
Public Sub runSlideValidator()

    Dim sldCurrent As Slide
    Dim colViolationMessages As Collection

    On Error GoTo error_handler
    basSlideValidator.cleanupViolationMessages
    For Each sldCurrent In ActivePresentation.Slides
        'ignore hidden slides
        If sldCurrent.SlideShowTransition.Hidden = msoFalse Then
            basSystemLogger.log "apply rules to slide " & sldCurrent.SlideIndex
            'TODO: apply rules
            Set colViolationMessages = New Collection
            colViolationMessages.Add "message six"
            colViolationMessages.Add "message five"
            If colViolationMessages.Count > 0 Then
                addViolation sldCurrent, colViolationMessages
            End If
            Set colViolationMessages = Nothing
        Else
            basSystemLogger.log "skip hidden slide " & sldCurrent.SlideIndex
        End If
    Next
    Exit Sub

error_handler:
    basSystemLogger.log_error "basSlideValidator.runRuleCheck"
End Sub
'-------------------------------------------------------------
' Description   : delete old violation messages
' Parameter     :
'-------------------------------------------------------------
Public Sub cleanupViolationMessages()

    Dim sldCurrent As Slide
    Dim comCurrentMsg As Comment
    Dim colOldMessages As Collection      'comment objects for old violation messages
    
    On Error GoTo error_handler
    basSystemLogger.log "delete old violation messages"
    For Each sldCurrent In ActivePresentation.Slides
        'ignore hidden slides
        If sldCurrent.SlideShowTransition.Hidden = msoFalse Then
            Set colOldMessages = New Collection
            'Powerpoint has problems deleting comments inside a for each loop from the comments property
            For Each comCurrentMsg In sldCurrent.Comments
                If comCurrentMsg.Author = mcRuleCheckAuthor Then
                    colOldMessages.Add comCurrentMsg
                End If
            Next
            For Each comCurrentMsg In colOldMessages
                comCurrentMsg.Delete
            Next
            Set colOldMessages = Nothing
        End If
    Next
    Exit Sub

error_handler:
    basSystemLogger.log_error "basSlideValidator.cleanupViolationMessages"
End Sub
