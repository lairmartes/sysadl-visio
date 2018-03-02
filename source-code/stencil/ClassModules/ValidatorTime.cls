VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidatorTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IValidator

Private pMinimumTime As Date
Private pMaximumtime As Date
Private pIssueMessage As String

Public Sub Init(ByVal minTime As Date, _
                 ByVal maxTime As Date, _
                 ByVal issueMessage As String)
                 
    pMinimumTime = minTime
    pMaximumtime = maxTime
    pIssueMessage = issueMessage
                 
End Sub

Public Function IValidator_Validate(ByVal aValue As String) As String

    Dim valueTime As Date
    Dim Result As String
    
    Result = sysAdlStringConstantsEmpty

On Error GoTo ConversionError:
    
    valueTime = CDate(aValue)
    
    If valueTime < pMinimumTime Or valueTime > pMaximumtime Then
    
        Result = pIssueMessage
            
    End If

ContinueValidate:

    IValidator_Validate = Result

Exit Function

ConversionError:

    Result = pIssueMessage
    
    Resume ContinueValidate:

End Function


