VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidatorNamespace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IValidator


Public Function IValidator_Validate(ByVal aNamespace As String) As String

    Dim Result As String
    Dim fieldLenght As Integer
    Dim IsSequentialPointFound As Boolean
    Dim IsEndWithPoint As Boolean
    Dim IsStartWithPoint As Boolean
    Dim IsCaracterInvalidFound As Boolean
    Dim IsFieldLenghOutOfLimits As Boolean
    
    Result = sysAdlStringConstantsEmpty
    
    fieldLenght = Len(aNamespace)
    
    IsFieldLenghOutOfLimits = (fieldLenght < sysAdlValidateMinLengthNamespace Or fieldLenght > sysAdlValidateMaxLengthNamespace)
    IsSequentialPointFound = InStr(aNamespace, sysAdlStringConstantsNamespacePointSequenceNotAllowed)
    IsStartWithPoint = (Left(aNamespace, 1) = sysAdlStringConstantsNamespacePathSeparator)
    IsEndWithPoint = (Right(aNamespace, 1) = sysAdlStringConstantsNamespacePathSeparator)
    
    If IsFieldLenghOutOfLimits Then
        
            Result = "Provide Namespace value from " + Str(sysAdlValidateMinLengthNamespace) + " to " _
                     + Str(sysAdlValidateMaxLengthNamespace) + " characters."
                     
    ElseIf IsSequentialPointFound Then
    
        Result = "It's not allowed to include the sequence '" + sysAdlStringConstantsNamespacePointSequenceNotAllowed + "'" + _
                 " in Namespace field"
                 
    ElseIf IsStartWithPoint Then
    
        Result = "It's not allowed to start Namespace with '" + sysAdlStringConstantsNamespacePathSeparator + "' character."
        
    ElseIf IsEndWithPoint Then
    
        Result = "It's not allowed to end Namespace field with '" + sysAdlStringConstantsNamespacePathSeparator + "' character."
        
    ElseIf Not IsAllNamespacePathValid(aNamespace) Then
    
        Result = "Namespace field only accepts leters, numbers and dots."
    
    End If
        
    IValidator_Validate = Result

End Function

Private Function IsAllNamespacePathValid(ByVal aNamespace As String) As Boolean

    Dim NamespaceRegExp As RegExp
    Dim IsNamespacePieceValid As Boolean
    Dim Result As Boolean
    Dim CurrentNamespacePiece As String
    Dim NamespacePathList As Variant
    Dim I As Integer
    
    Set NamespaceRegExp = New RegExp
    
    Result = True
    
    NamespaceRegExp.Pattern = sysAdlStringConstantsRegExpAllowedPathQualifier
    
    NamespacePathList = Split(aNamespace, sysAdlStringConstantsNamespacePathSeparator)
        
    For I = 0 To UBound(NamespacePathList)
            
        CurrentNamespacePiece = NamespacePathList(I)
        
        IsNamespacePieceValid = NamespaceRegExp.Test(CurrentNamespacePiece)
        
        If Not IsNamespacePieceValid Then
        
            Result = False
            
            Exit For
            
        End If
        
                
    Next I

    IsAllNamespacePathValid = Result

End Function


