VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidatorDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IValidator

Private pAllowPastDate As Boolean
Private pAllowFutureDate As Boolean
Private pAllowPresentDay As Boolean
Private pIssueMessage As String

Public Sub Init(ByVal AllowPast As Boolean, _
                 ByVal AllowFuture As Boolean, _
                 ByVal AllowPresent As Boolean, _
                 ByVal issueMessage As String)
                 
    pAllowPastDate = AllowPast
    pAllowFutureDate = AllowFuture
    pAllowPresentDay = AllowPresent
    pIssueMessage = issueMessage
                 
End Sub

Public Function IValidator_Validate(ByVal aValue As String) As String

    Dim valueDate As Date
    Dim currentDate As Date
    
    Dim Result As String
    
    Dim AllowPastNotPresentNotFuture As Boolean
    Dim AllowPastPresentNotFuture As Boolean
    Dim AllowNotPastPresentFuture As Boolean
    Dim AllowNotPastNotPresentFuture As Boolean
    Dim AllowsNothing As Boolean
    
    Result = sysAdlStringConstantsEmpty

On Error GoTo ConversionError:
    
    valueDate = CDate(aValue)
    currentDate = VBA.Date()
        
    AllowPastNotPresentNotFuture = pAllowPastDate And Not pAllowPresentDay And Not pAllowFutureDate
    
    AllowPastPresentNotFuture = pAllowPastDate And pAllowPresentDay And Not pAllowFutureDate
    
    AllowNotPastPresentFuture = Not pAllowPastDate And pAllowPresentDay And pAllowFutureDate
    
    AllowNotPastNotPresentFuture = Not pAllowPastDate And Not pAllowPresentDay And pAllowFutureDate
    
    AllowsNothing = Not pAllowPastDate And Not pAllowPresentDay And Not pAllowFutureDate
    
    If AllowPastNotPresentNotFuture And valueDate >= currentDate Then
    
        Result = pIssueMessage
    
    ElseIf AllowPastPresentNotFuture And valueDate > currentDate Then
    
        Result = pIssueMessage
        
    ElseIf AllowNotPastPresentFuture And valueDate < currentDate Then
    
        Result = pIssueMessage
        
    ElseIf AllowNotPastNotPresentFuture And valueDate <= currentDate Then
    
        Result = pIssueMessage
        
    ElseIf AllowsNothing Then
        
        Result = pIssueMessage
        
    End If

ContinueValidate:

    IValidator_Validate = Result

Exit Function

ConversionError:

    Result = pIssueMessage
    
    Resume ContinueValidate:

End Function
