VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidatorElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IValidator

Private pIssueMessage As String
Private pElementType As String

Public Sub Init(ByVal elementType As String, ByVal issueMessage As String)

    pElementType = elementType
    pIssueMessage = issueMessage

End Sub

Public Function IValidator_Validate(ByVal aValue As String) As String

    Dim elementNamespace As String
    Dim elementId As String
    Dim elementParameter As SysAdlElement
    Dim elementExists As Boolean
    Dim keyIsWellFormed As Boolean
    Dim Result As String
    
    Result = sysAdlStringConstantsEmpty
            
    elementNamespace = GUIServices.GetNamespaceFromString(aValue)
    elementId = GUIServices.GetIdFromString(aValue)
        
    Set elementParameter = New SysAdlElement

    elementParameter.Init pElementType, sysAdlNoStereotype
    elementParameter.InitBaseData elementNamespace, elementId, sysAdlStringConstantsEmpty
        
    elementExists = ElementServiceCache.IsElementExistsWithType(elementParameter)
        
    If Not elementExists Then
        
        Result = pIssueMessage
            
    End If
    
    IValidator_Validate = Result

End Function
