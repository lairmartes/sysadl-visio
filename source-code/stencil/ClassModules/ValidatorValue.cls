VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidatorValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IValidator

Private pMinValue As Double
Private pMaxValue As Double
Private pOnlyInteger As Boolean
Private pIssueMessage As String

Public Sub Init(ByVal minValue As Double, _
                ByVal maxValue As Double, _
                ByVal isOnlyInteger As Boolean, _
                ByVal issueMessage As String)
                
                
        pMinValue = minValue
        pMaxValue = maxValue
        pOnlyInteger = isOnlyInteger
        pIssueMessage = issueMessage

End Sub

Public Function IValidator_Validate(ByVal aValue As String) As String

    Dim value As Double
    Dim Result As String
    
    Result = sysAdlStringConstantsEmpty

On Error GoTo ConversionError:
    
    value = CDbl(aValue)
    
    If value < pMinValue Or value > pMaxValue Then
    
        Result = pIssueMessage
        
    ElseIf pOnlyInteger And Not IsInteger(value) Then
    
        Result = pIssueMessage
                
    End If

ContinueValidate:

    IValidator_Validate = Result

Exit Function

ConversionError:

    Result = pIssueMessage
    
    Resume ContinueValidate:

End Function

Private Function IsInteger(ByVal value As Double) As Boolean

    Dim Result As Boolean
    Dim valueInteger As Integer
    Dim valueMinusInteger As Boolean
    
    Result = False ' first, I think that number is not integer
    
    valueInteger = value \ 1
    valueMinusInteger = value - valueInteger
    
    If valueMinusInteger = 0 Then
    
        Result = True 'is integer because it has no decimal places
        
    End If
    
    IsInteger = Result
    
End Function
