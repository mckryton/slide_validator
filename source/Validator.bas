Attribute VB_Name = "Validator"
'------------------------------------------------------------------------
' Description  : apply rules to active presentation, add violations as comment
'------------------------------------------------------------------------
Option Explicit

Const mcRuleCheckAuthor = "Slide Validator"
Const mcRuleCheckInitials = "bot"

Public Sub run_slide_validator(Optional ppresPresentation, Optional pvarRules)

    Dim sldCurrent As Slide

    On Error GoTo error_handler
    If IsMissing(ppresPresentation) Then
        Set ppresPresentation = ActivePresentation
    End If
    If IsMissing(pvarRules) Then
        'TODO: add function to setup all rules
        pvarRules = Array()
    End If
    'comments from earlier validations may not reflect the current content
    Validator.cleanup_violation_messages
    For Each sldCurrent In ppresPresentation.Slides
        'hidden slides contain most often discarded content and can be ignored
        If sldCurrent.SlideShowTransition.Hidden = msoFalse Then
            SystemLogger.log "apply rules to slide " & sldCurrent.SlideIndex
            apply_rules pvarRules, sldCurrent
        Else
            SystemLogger.log "skip hidden slide " & sldCurrent.SlideIndex
        End If
    Next
    Exit Sub

error_handler:
    SystemLogger.log_error "Validator.runRuleCheck"
End Sub

Private Function apply_rules(pvarRules As Variant, psldCurrentSlide As Slide)

    Dim varRule As Variant
    Dim strValidationResult As String
    
    On Error GoTo error_handler
    For Each varRule In pvarRules
        strValidationResult = varRule.apply_rule(psldCurrentSlide)
        If Len(Trim(strValidationResult)) > 0 Then
           add_violation psldCurrentSlide, strValidationResult
        End If
    Next
    Exit Function
    
error_handler:
    SystemLogger.log_error "Validator.apply_rule"
End Function

Public Sub add_violation(psldValidatedSlide As Slide, pstrViolationMessage As String)
    
    Dim lngCommentPosX As Long
    
    On Error GoTo error_handler
    'improve visibilty by putting all comments for violation messages in a row
    lngCommentPosX = 10 * (psldValidatedSlide.Comments.Count + 1)
    psldValidatedSlide.Comments.Add lngCommentPosX, 10, mcRuleCheckAuthor, mcRuleCheckInitials, pstrViolationMessage
    Exit Sub

error_handler:
    SystemLogger.log_error "Validator.add_violation"
End Sub

Private Sub cleanup_violation_messages()

    Dim sldCurrent As Slide
    Dim comCurrentMsg As Comment
    Dim colOldMessages As Collection      'comment objects for old violation messages
    
    On Error GoTo error_handler
    SystemLogger.log "delete old violation messages"
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
    SystemLogger.log_error "Validator.cleanup_violation_messages"
End Sub

Public Function read_config(pRuleName As String, Optional pConfigPresentation) As Collection
    
    Dim rule_config As Collection

    Set rule_config = New Collection

    Set read_config = rule_config
End Function
