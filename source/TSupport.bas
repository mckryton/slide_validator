Attribute VB_Name = "TSupport"
'this module is for sharing helper functions between your test cases

Option Explicit

Public Function create_config_slide(pConfigPresentation As Presentation, Optional pRuleName) As Slide
    
    Dim validator_pres As Presentation
    Dim config_slide As Slide
    Dim rule_name As String
    Dim config_template_index As Integer
    
    If IsMissing(pRuleName) Then
        rule_name = "RuleName"
    Else
        rule_name = pRuleName
    End If
    Set validator_pres = Application.Presentations("SlideValidator.pptm")
    config_template_index = get_config_template_index(pConfigPresentation)
    Set config_slide = pConfigPresentation.Slides.AddSlide(1, pConfigPresentation.SlideMaster.CustomLayouts(config_template_index))
    config_slide.Shapes.Title.TextFrame.TextRange.Text = rule_name
    config_slide.Name = rule_name
    Set create_config_slide = config_slide
End Function

Public Function create_slide_validator_pres() As Presentation

    Dim config_presentation As Presentation
    Dim slide_validator As Presentation
    Dim master_slide As CustomLayout
    Dim config_template_index As Integer
        
    Set slide_validator = Application.Presentations("SlideValidator.pptm")
    Set config_presentation = Application.Presentations.Add
    
    config_template_index = get_config_template_index(slide_validator)
    Set master_slide = slide_validator.SlideMaster.CustomLayouts(config_template_index)
    master_slide.Copy
    config_presentation.SlideMaster.CustomLayouts.Paste
    
    Set create_slide_validator_pres = config_presentation
End Function

Private Function get_config_template_index(pConfigPresentation As Presentation) As Integer

    Dim custom_layout As CustomLayout
    
    For Each custom_layout In pConfigPresentation.SlideMaster.CustomLayouts
        If custom_layout.Name = Validator.CONFIG_TEMPLATE_NAME Then
            get_config_template_index = custom_layout.index
            Exit Function
        End If
    Next
    Err.raise Validator.ERR_ID_MISSING_CFG_MASTER_SLIDE, description:="couldn't find a custom layout named >" & _
                Validator.CONFIG_TEMPLATE_NAME & "< in presentation " & pConfigPresentation.Name
End Function

