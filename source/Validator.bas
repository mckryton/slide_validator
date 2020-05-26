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

Public Sub validate_slides(Optional pTargetPresentation, Optional pvarRules)

    Dim sldCurrent As Slide

    On Error GoTo error_handler
    If IsMissing(pTargetPresentation) Then
        Set pTargetPresentation = Validator.ValidationTarget
        If TypeName(pTargetPresentation) = "Nothing" Then
            MsgBox "Couldn't find any open presentation to apply validation rules.", vbExclamation + vbOKOnly, "No presentation for validation available"
            End
        End If
    End If
    If IsMissing(pvarRules) Then
        'TODO: add function to setup all rules
        pvarRules = Array()
    End If
    'remove comments from earlier validations because they may not reflect the current content
    Validator.cleanup_violation_comments
    For Each sldCurrent In pTargetPresentation.Slides
        'hidden slides contain most often discarded content and can be ignored
        If sldCurrent.SlideShowTransition.Hidden = msoFalse Then
            SystemLogger.Log "apply rules to slide " & sldCurrent.SlideIndex
            apply_rules pvarRules, sldCurrent
        Else
            SystemLogger.Log "skip hidden slide " & sldCurrent.SlideIndex
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
    'improve visibility by putting all comments for violation messages in a row
    lngCommentPosX = 10 * (psldValidatedSlide.Comments.Count + 1)
    psldValidatedSlide.Comments.Add lngCommentPosX, 10, COMMENT_AUTHOR, COMMENT_INITIALS, pstrViolationMessage
    Exit Sub

error_handler:
    SystemLogger.log_error "Validator.add_violation"
End Sub

Private Sub cleanup_violation_comments()

    Dim sldCurrent As Slide
    Dim comCurrentMsg As Comment
    Dim colOldMessages As Collection      'comment objects for old violation messages
    
    On Error GoTo error_handler
    SystemLogger.Log "delete old violation messages"
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
    SystemLogger.log_error "Validator.cleanup_violation_messages"
End Sub

Public Function read_config(pConfigSlide As Slide) As Collection
    
    Dim config_table As Table
    
    Set config_table = get_config_table(pConfigSlide)
    If TypeName(config_table) <> "Nothing" Then
        Set read_config = read_config_from_table(config_table)
    Else
        'return an empty collection to be able to count available settings in any case
        Set read_config = New Collection
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
    select_target_form.PresentationsInfo = get_target_presentations_info_info()
    Set get_validation_target_form = select_target_form
End Function

Public Function get_target_presentations_info_info() As Collection
    
    Dim target_presentation_names As Collection
    Dim open_presentation As Presentation
    
    Set target_presentation_names = New Collection
    For Each open_presentation In Application.Presentations
        If open_presentation.Name <> "SlideValidator.pptm" Then
            target_presentation_names.Add Array(open_presentation.Name, open_presentation.Path), open_presentation.Name
        End If
    Next
    Set get_target_presentations_info_info = target_presentation_names
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

    If pConfigSlide.Master.Name = "rule_config" Then
        is_config_slide = True
    Else
        is_config_slide = False
    End If
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
        If config_slide.Master.Name = CONFIG_TEMPLATE_NAME Then
            rule_name = Trim(config_slide.Shapes.Title.TextFrame.TextRange.Text)
            On Error GoTo missing_rule
            Set validation_rule = get_rule(rule_name)
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
