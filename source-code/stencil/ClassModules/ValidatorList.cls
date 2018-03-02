VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidatorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IValidator

Private pIssueMessage As String


Public Sub Init(ByVal issueMessage As String)

    pIssueMessage = issueMessage
    
End Sub

Public Function IValidator_Validate(ByVal aValue As String) As String

    Dim isValidValue As Boolean
    Dim Result As String

    Result = sysAdlStringConstantsEmpty
    
    isValidValue = (aValue <> sysAdlStringConstantsEmpty)
    
    If Not isValidValue Then
    
        Result = pIssueMessage
        
    End If
    
    IValidator_Validate = Result

    Exit Function
    
End Function


