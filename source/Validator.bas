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
    Dim config_presentation As presentation
    Dim config_slide As Slide
    Dim config_table As Table
    
    Set rule_config = New Collection
    If IsMissing(pConfigPresentation) Then
        Set config_presentation = Presentations("SlideValidator.pptm")
    Else
        Set config_presentation = pConfigPresentation
    End If
    Set config_slide = get_config_slide(pRuleName, config_presentation)
    If TypeName(config_slide) <> "Nothing" Then
        Set config_table = get_config_table(config_slide)
        If TypeName(config_table) <> "Nothing" Then
            Set rule_config = read_config_from_table(config_table)
        End If
    End If
    Set read_config = rule_config
End Function

Private Function get_config_slide(pstrRuleName As String, pConfigPresentation As presentation) As Slide

    Dim config_slide As Slide
    Dim slide_title As String
    
    For Each config_slide In pConfigPresentation.Slides
        slide_title = ""
        On Error Resume Next
        slide_title = Trim(config_slide.Shapes.Title.TextFrame.TextRange.Text)
        If slide_title = pstrRuleName Then
            Set get_config_slide = config_slide
            Exit Function
        End If
    Next
    Set get_config_slide = Nothing
End Function

Private Function get_config_table(pConfigSlide As Slide) As Table

    Dim config_shape As shape
    
    For Each config_shape In pConfigSlide.Shapes
        If config_shape.HasTable Then
            Set get_config_table = config_shape.Table
            Exit Function
        End If
    Next
    Set get_config_table = Nothing
End Function

Private Function read_config_from_table(pConfigTable As Table) As Collection

    Dim config_row As Row
    Dim row_nr As Long
    Dim config_parameters As Collection
    
    Set config_parameters = New Collection
    For row_nr = 2 To pConfigTable.Rows.Count
        Set config_row = pConfigTable.Rows(row_nr)
        config_parameters.Add Trim(config_row.Cells(2).shape.TextFrame.TextRange), Trim(config_row.Cells(1).shape.TextFrame.TextRange)
    Next
    Set read_config_from_table = config_parameters
End Function


Public Function get_validation_target_form() As SelectValidationTarget

    Dim select_target_form As SelectValidationTarget
    
    Set select_target_form = New SelectValidationTarget
    select_target_form.PresentationsInfo = get_target_presentations_info_info()
    Set get_validation_target_form = select_target_form
End Function

Public Function get_target_presentations_info_info() As Collection
    
    Dim target_presentation_names As Collection
    Dim open_presentation As presentation
    
    Set target_presentation_names = New Collection
    For Each open_presentation In Application.Presentations
        If open_presentation.Name <> "SlideValidator.pptm" Then
            target_presentation_names.Add Array(open_presentation.Name, open_presentation.Path), open_presentation.Name
        End If
    Next
    Set get_target_presentations_info_info = target_presentation_names
End Function
