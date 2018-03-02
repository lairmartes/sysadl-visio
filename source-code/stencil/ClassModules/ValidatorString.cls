VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidatorString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IValidator

Private pRegExp As RegExp
Private pIssueMessage As String
Private pStringSizeLimitVisio2003Compatibility As String


Public Sub Init(ByVal Pattern As String, ByVal issueMessage As String)

    Dim ErrorDetectedMessage As String
    
    ErrorDetectedMessage = "A syntax error in regular expression '" + Pattern + _
                        "' has been detected.  Correct regular expression and try again."
    
    On Error GoTo ValidateStringRegExpErrorHandle
    
    Set pRegExp = New RegExp
    
    pRegExp.Pattern = Pattern
    
    pIssueMessage = issueMessage
    
    pRegExp.Test (sysAdlStringConstantsEmpty)
    
    pStringSizeLimitVisio2003Compatibility = CalculateSizeLimit(Pattern)
    
    If pStringSizeLimitVisio2003Compatibility > sysAdlIntegerConstantsMaxStringSizeForField Then
    
        ErrorDetectedMessage = "Field validator: String limit cannot exceed" + Str(sysAdlIntegerConstantsMaxStringSizeForField) + " characters." + vbCrLf + vbCrLf + _
                               "Please, reduce size limit value in regular expression pattern " + Pattern + " defined for this field."
        
        MsgBox ErrorDetectedMessage
        
        GUIServices.LogMessage ErrorDetectedMessage
        
        GoTo ValidateStringRegExpErrorHandle
        
    End If
    
    Exit Sub
    
ValidateStringRegExpErrorHandle:

    Err.Raise 9000, "", ErrorDetectedMessage

    
End Sub

Public Function IValidator_Validate(ByVal aValue As String) As String

    Dim isValidValue As Boolean
    Dim Result As String

    Result = sysAdlStringConstantsEmpty
    
    isValidValue = pRegExp.Test(aValue)
    
    If Not isValidValue Then ' RegExp should validate size limit
    
        Result = pIssueMessage
    
    ElseIf pStringSizeLimitVisio2003Compatibility > 0 Then
    
        ' if it reach this node then Visio2003 does not validated string upper limit correctly
    
        If Len(aValue) > pStringSizeLimitVisio2003Compatibility Then
        
            Result = pIssueMessage
            
        End If
        
    End If
    
    IValidator_Validate = Result

    Exit Function
    
End Function

Private Function CalculateSizeLimit(ByVal RegExpPattern As String) As Integer

    Dim LimitCalculatorPattern As String
    Dim LimitCalculatorRegExp As RegExp
    Dim PatternSize As Integer
    Dim I As Integer
    Dim PatternCharFound As String
    Dim SizeStringFound As String
    Dim Result As Integer
    
    LimitCalculatorPattern = "({)([0-9])+(,)([0-9])+(})"
    
    Set LimitCalculatorRegExp = New RegExp

    LimitCalculatorRegExp.Pattern = LimitCalculatorPattern
    
    PatternSize = Len(RegExpPattern)
    
    SizeStringFound = sysAdlStringConstantsEmpty
    
    For I = PatternSize - 1 To 1 Step -1
    
        PatternCharFound = Mid$(RegExpPattern, I, 1)
        
        If PatternCharFound = sysAdlStringConstantsComma Then
        
            Exit For
        
        End If
    
        SizeStringFound = PatternCharFound + SizeStringFound
        
    Next
    
    Result = Val(SizeStringFound)
    
    CalculateSizeLimit = Result
    
End Function
