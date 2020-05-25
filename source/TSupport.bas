Attribute VB_Name = "TSupport"
Option Explicit

Public Function create_config_slide(pConfigPresentation As Presentation, pRuleName As String) As Slide
    
    Dim config_slide As Slide
    
    Set config_slide = pConfigPresentation.Slides.AddSlide(1, get_config_template())
    config_slide.Shapes.Title.TextFrame.TextRange.Text = pRuleName
    config_slide.Name = pRuleName
    Set create_config_slide = config_slide
End Function

Private Function get_config_template() As CustomLayout

    Dim validator_pres As Presentation
    Dim custom_layout As CustomLayout
    
    Set validator_pres = Application.Presentations("SlideValidator.pptm")
    
    For Each custom_layout In ActivePresentation.SlideMaster.CustomLayouts
        If custom_layout.Name = Validator.CONFIG_TEMPLATE_NAME Then
            Set get_config_template = custom_layout
            Exit Function
        End If
    Next
    Err.raise Validator.ERR_ID_MISSING_CFG_MASTER_SLIDE, description:="couldn't find a custom layout named " & _
                Validator.CONFIG_TEMPLATE_NAME & " in slide master of SlideValidator"
End Function
