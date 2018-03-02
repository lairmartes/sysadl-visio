VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidatorId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IValidator


Public Function IValidator_Validate(ByVal value As String) As String

    Dim Result As String
    Dim fieldLenght As Integer
    
    Result = sysAdlStringConstantsEmpty
    
    fieldLenght = Len(value)
        
    If fieldLenght < sysAdlValidateMinLengthId Or fieldLenght > sysAdlValidateMaxLengthId Then
        
            Result = "Provide Id value from " + Str(sysAdlValidateMinLengthId) + " to " _
                    + Str(sysAdlValidateMaxLengthId) + " characters."
                    
    ElseIf Not IsIdValid(value) Then
    
        Result = "Id field only accepts leters or numbers"
        
    End If
        
    IValidator_Validate = Result

End Function

Private Function IsIdValid(ByVal anId As String) As Boolean

    Dim IdRegExp As RegExp
    Dim Result As Boolean
    
    Set IdRegExp = New RegExp

    IdRegExp.Pattern = sysAdlStringConstantsRegExpAllowedPathQualifier
    
    Result = IdRegExp.Test(anId)

    IsIdValid = Result

End Function

