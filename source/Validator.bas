Attribute VB_Name = "Validator"
'------------------------------------------------------------------------
' Description  : apply rules to active presentation, add violations as comment
'------------------------------------------------------------------------
Option Explicit

Public Const CONFIG_TEMPLATE_NAME = "rule_config"

Public Const ERR_ID_MISSING_CFG_MASTER_SLIDE = vbError + 7000
Public Const ERR_ID_UNKNOWN_RULE_NAME = vbError + 7050

Const COMMENT_AUTHOR = "Slide Validator"
Const COMMENT_INITIALS = "bot"

Dim mLogger As Logger

Public Sub validate_presentation(Optional pTargetPresentation, Optional pValidationSetup, Optional p_silent)

    Dim sldCurrent As Slide
    Dim target_presentation As Presentation
    Dim validation_setup As ValidationSetup
    Dim validation_log As ValidationLog
    Dim violations As Collection

    On Error GoTo error_handler
    If Application.SlideShowWindows.Count > 0 Then
        Log.info_log "exit presentation mode"
        ActivePresentation.SlideShowWindow.View.Exit
    End If
    'if the macro is called from presentation mode, it will call the function with the clicked shape object as parameter
    If IsMissing(pTargetPresentation) Or TypeName(pTargetPresentation) = "Shape" Then
        Set target_presentation = Validator.ValidationTarget
        If TypeName(target_presentation) = "Nothing" Then
            MsgBox "Couldn't find any open presentation to apply validation rules.", vbExclamation + vbOKOnly, "No presentation for validation is available"
            End
        End If
    Else
        Set target_presentation = pTargetPresentation
    End If
    Log.info_log "start validation of >" & target_presentation.Name & "<"
    target_presentation.Windows(1).Activate
    If IsMissing(pValidationSetup) Then
        Set validation_setup = setup_rules(Application.Presentations("SlideValidator.pptm"))
    Else
        Set validation_setup = pValidationSetup
    End If
    Set validation_log = New ValidationLog
    'remove comments from earlier validations because they may not reflect the current content
    Validator.cleanup_violation_comments
    For Each sldCurrent In target_presentation.Slides
        'hidden slides contain most often discarded content and can be ignored
        If sldCurrent.SlideShowTransition.Hidden = msoFalse Then
            Log.info_log "apply rules to slide " & sldCurrent.SlideIndex
            Set violations = apply_rules_on_slide(validation_setup.ActiveRules, sldCurrent)
            If Not violations Is Nothing Then
                validation_log.violations.Add violations
            End If
        Else
            Log.info_log "skip hidden slide " & sldCurrent.SlideIndex
        End If
    Next
    If IsMissing(p_silent) Then
        MsgBox "Validation is complete. Found violations on " & validation_log.violations.Count & " slide(s).", vbOKOnly, "SlideValidator finished validation"
    End If
    Exit Sub

error_handler:
    Log.log_function_error "Validator.validate_presentation"
End Sub

Public Function apply_rules_on_slide(pvarRules As Collection, psldCurrentSlide As Slide) As Collection

    Dim rule As Variant
    Dim validation_result As String
    Dim violations As Collection
    
    Set violations = New Collection
    For Each rule In pvarRules
        validation_result = rule.apply_rule(psldCurrentSlide)
        If Len(Trim(validation_result)) > 0 Then
           add_violation_comment psldCurrentSlide, validation_result
           violations.Add validation_result
        End If
    Next
    If violations.Count = 0 Then
        Set apply_rules_on_slide = Nothing
    Else
        Set apply_rules_on_slide = violations
    End If
End Function

Public Sub add_violation_comment(psldValidatedSlide As Slide, pstrViolationMessage As String)
    
    Dim lngCommentPosX As Long
    
    On Error GoTo error_handler
    'improve visibility by putting all comments for violation messages in a row
    lngCommentPosX = 10 * (psldValidatedSlide.Comments.Count + 1)
    psldValidatedSlide.Comments.Add lngCommentPosX, 10, COMMENT_AUTHOR, COMMENT_INITIALS, pstrViolationMessage
    Exit Sub

error_handler:
    Log.log_function_error "Validator.add_violation_comment"
End Sub

Private Sub cleanup_violation_comments()

    Dim sldCurrent As Slide
    Dim comCurrentMsg As Comment
    Dim colOldMessages As Collection      'comment objects for old violation messages
    
    On Error GoTo error_handler
    Log.info_log "delete old violation messages"
    For Each sldCurrent In ActivePresentation.Slides
        'ignore hidden slides
        If sldCurrent.SlideShowTransition.Hidden = msoFalse Then
            Set colOldMessages = New Collection
            'Powerpoint has problems deleting comments inside a for each loop from the comments property
            For Each comCurrentMsg In sldCurrent.Comments
                If comCurrentMsg.Author = COMMENT_AUTHOR Then
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
    Log.log_function_error "Validator.cleanup_violation_messages"
End Sub

Public Function get_rule_config(pConfigSlide As Slide) As Collection
    
    Dim config_table As Table
    
    Set config_table = get_config_table(pConfigSlide)
    If TypeName(config_table) <> "Nothing" Then
        Set get_rule_config = read_config_from_table(config_table)
    Else
        'return an empty collection to be able to count available settings in any case
        Set get_rule_config = New Collection
    End If
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
    Dim param_name As String
    Dim param_value As String
    
    Set config_parameters = New Collection
    For row_nr = 2 To pConfigTable.Rows.Count
        Set config_row = pConfigTable.Rows(row_nr)
        param_name = Trim(config_row.Cells(1).shape.TextFrame.TextRange)
        param_value = Trim(config_row.Cells(2).shape.TextFrame.TextRange)
        config_parameters.Add param_value, param_name
    Next
    Set read_config_from_table = config_parameters
End Function

Public Function get_validation_target_form() As SelectValidationTarget

    Dim select_target_form As SelectValidationTarget
    
    Set select_target_form = New SelectValidationTarget
    select_target_form.PresentationsInfo = get_target_presentations_info()
    Set get_validation_target_form = select_target_form
End Function

Public Function get_target_presentations_info() As Collection
    
    Dim target_presentation_names As Collection
    Dim open_presentation As Presentation
    
    Set target_presentation_names = New Collection
    For Each open_presentation In Application.Presentations
        If open_presentation.Name <> "SlideValidator.pptm" Then
            target_presentation_names.Add Array(open_presentation.Name, open_presentation.Path), open_presentation.Name
        End If
    Next
    Set get_target_presentations_info = target_presentation_names
End Function

Public Property Get ValidationTarget() As Presentation
    
    Dim selection_form As SelectValidationTarget
    
    Set selection_form = Validator.get_validation_target_form
    If UBound(selection_form.lstPresentations.List) = -1 Then
        Set ValidationTarget = Nothing
    ElseIf UBound(selection_form.lstPresentations.List) = 0 Then
        Set ValidationTarget = Application.Presentations(selection_form.lstPresentations.List(0))
    Else
        Set ValidationTarget = Application.Presentations(selection_form.lstPresentations.Value)
    End If
End Property

Public Function is_config_slide(pConfigSlide As Slide) As Boolean

    Dim slide_shape As shape
    
    is_config_slide = False
    If LCase(Left(Trim(pConfigSlide.Shapes.Title.TextFrame.TextRange.Text), 4)) <> "rule" Then
        Exit Function
    End If
    For Each slide_shape In pConfigSlide.Shapes
        If slide_shape.HasTable Then
            If slide_shape.Table.Columns.Count >= 3 Then
                If LCase(Trim(slide_shape.Table.Cell(1, 1).shape.TextFrame.TextRange.Text)) = "parameter" _
                  And LCase(Trim(slide_shape.Table.Cell(1, 2).shape.TextFrame.TextRange.Text)) = "value" _
                  And LCase(Trim(slide_shape.Table.Cell(1, 3).shape.TextFrame.TextRange.Text)) = "description" Then
                    is_config_slide = True
                End If
            End If
        End If
    Next
End Function

Public Function get_rule(pRuleName As String) As Variant

    Dim rule_catalog As RuleCatalog
    
    Set rule_catalog = New RuleCatalog
    On Error GoTo error_handler
    Set get_rule = CallByName(rule_catalog, pRuleName, VbGet)
    Set rule_catalog = Nothing
    Exit Function
    
error_handler:
    Err.raise ERR_ID_UNKNOWN_RULE_NAME, description:="can't find a rule class named >" & pRuleName & "<"
End Function

Public Function setup_rules(Optional pConfigPres) As ValidationSetup
    
    Dim slide_validator As Presentation
    Dim validation_setup As ValidationSetup
    Dim config_slide As Slide
    Dim rule_name As String
    Dim validation_rule As Variant
    
    If IsMissing(pConfigPres) Then
        Set slide_validator = Application.Presentations("SlideValidator.pptm")
    Else
        Set slide_validator = pConfigPres
    End If
    Set validation_setup = New ValidationSetup
    For Each config_slide In slide_validator.Slides
        If is_config_slide(config_slide) Then
            rule_name = Replace(Trim(Split(config_slide.Shapes.Title.TextFrame.TextRange.Text, ":")(1)), " ", "_")
            On Error GoTo missing_rule
            Set validation_rule = get_rule(rule_name)
            validation_rule.Config = get_rule_config(config_slide)
            On Error GoTo 0
            If TypeName(validation_rule) <> "Nothing" Then
                validation_setup.ActiveRules.Add validation_rule, rule_name
            End If
        End If
    Next
    Set setup_rules = validation_setup
    Exit Function
    
missing_rule:
    validation_setup.SetupErrors.Add "couldn't find a rule for config >" & rule_name & "<"
    Set validation_rule = Nothing
    Resume Next
End Function

Public Property Get Log() As Logger
    
    If TypeName(mLogger) = "Nothing" Then
        Set mLogger = New Logger
    End If
    Set Log = mLogger
End Property
