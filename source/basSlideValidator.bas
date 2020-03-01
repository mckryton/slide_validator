Attribute VB_Name = "basRuleChecker"
'------------------------------------------------------------------------
' Description  : apply rules to active presentation, add violations as comment
'------------------------------------------------------------------------
'
'Declarations
Const mcRuleCheckAuthor = "ECC Presentation Rules Checker"
Const mcRuleCheckInitials = "ERC"

'Declare variables

'Options
Option Explicit
'-------------------------------------------------------------
' Description   : add rule violation message as ppt comment
' Parameter     : psldValidatedSlide    - the current validated slide
'                 pstrViolationMsg      - feedback from the rule
'-------------------------------------------------------------
Public Sub addViolation(psldValidatedSlide As Slide, pstrViolationMsg As String)
    
    On Error GoTo error_handler
    psldValidatedSlide.Comments.Add 10, 10, mcRuleCheckAuthor, mcRuleCheckInitials, pstrViolationMsg
    Exit Sub

error_handler:
    basSystem.log_error "basRuleChecker.addViolation"
End Sub
