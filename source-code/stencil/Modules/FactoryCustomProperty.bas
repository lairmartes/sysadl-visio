Attribute VB_Name = "FactoryCustomProperty"
Option Explicit

Private CUSTOM_PROPERTIES_LIST As Collection
Private IsCustomPropertiesInitialized As Boolean
'generic custom properties
Private CustomPropertyBaseId As New CustomProperty
Private CustomPropertyBaseNamespace As New CustomProperty
Private CustomPropertyURLInfo As New CustomProperty
Private VALIDATOR_NAMESPACE As New ValidatorNamespace
Private VALIDATOR_ID As New ValidatorId
Private VALIDATOR_STEREOTYPE As New ValidatorNotMandatory
Private VALIDATOR_URLINFO As New ValidatorNotMandatory
Private EMPTY_CUSTOM_PROPERTY_SET As New CustomPropertySet
Private Const FIELD_IS_MANDATORY = True
Private Const FIELD_IS_NOT_MANDATORY = False

Public Function Create(ByVal sysadlType As String, ByVal SysADLStereotype As String) As CustomPropertySet

    Dim Result As CustomPropertySet
    Dim CreateCustomPropertiesForNoRelationElement As Boolean
    
    On Error GoTo HandlCreateError
    
    CreateCustomPropertiesForNoRelationElement = Not GUIServices.IsRelationElementType(sysadlType)
    
    If CreateCustomPropertiesForNoRelationElement Then
    
        If SysADLStereotype = sysAdlStringConstantsEmpty Then
            
            SysADLStereotype = sysAdlNoStereotype
        
        End If
        
        Set Result = GetCustomProperties(sysadlType, SysADLStereotype)
        
    Else
    
        Set Result = EMPTY_CUSTOM_PROPERTY_SET
        
    End If
        
    Set Create = Result.Clone
    
    Exit Function
    
HandlCreateError:

    MsgBox Err.Description

End Function


Private Function GetCustomProperties(ByVal elementType As String, ByVal ElementStereotype As String) As CustomPropertySet
    
    Dim RecoveredDefinitionFields As Collection
    Dim ElementDefinitionVO As VOElementDefinition
    Dim ElementDefinitionFieldVO As VOElementDefinitionField
    
    Dim FieldDefinitionVO As VOFieldDefinition
    
    Dim ElementFieldName As String
    Dim ElementFieldLabel As String
    Dim ElementFieldType As String
    Dim ElementFieldDescription As String
    
    Dim ElementFieldMandatory As Boolean
    
    Dim DesignOrder As Integer
    Dim DesignParenthesis As Boolean
    Dim CommentOrder As Integer
    Dim CommentParenthesis As Boolean
    
    Dim VisioFieldType As Integer
    
    Dim CurrentProperty As CustomProperty
    Dim Result As New CustomPropertySet

    Dim CustomPropertyValidator As IValidator
    
    
    InitializeCustomProperties
    
    Result.AddCustomProperty CreateStereotypeCustomProperty(elementType)
    
    Result.AddCustomProperty CustomPropertyBaseNamespace
    
    Result.AddCustomProperty CustomPropertyBaseId

    
    Set ElementDefinitionVO = FactoryDefinitionElement.GetElementDefinition(elementType, ElementStereotype)
    
    Set RecoveredDefinitionFields = ElementDefinitionVO.GetElementFieldList
    
    For Each ElementDefinitionFieldVO In RecoveredDefinitionFields
    
        ElementFieldName = ElementDefinitionFieldVO.FieldName
    
        Set FieldDefinitionVO = FactoryDefinitionField.GetFieldDefinition(ElementFieldName)
    
        Set CurrentProperty = New CustomProperty
           
        ElementFieldLabel = FieldDefinitionVO.Label
        ElementFieldType = FieldDefinitionVO.FieldType
        ElementFieldDescription = FieldDefinitionVO.Description
        
        Select Case ElementFieldType
        
            Case sysAdlFieldTypeDate
                VisioFieldType = visPropTypeString
        
            Case sysAdlFieldTypeList
                VisioFieldType = visPropTypeListFix
                
            Case sysAdlFieldTypeValue
                VisioFieldType = visPropTypeString
                
            Case Else
                VisioFieldType = visPropTypeString
                
        End Select
      
        
        If VisioFieldType <> visPropTypeListFix Then
        
            CurrentProperty.Init ElementFieldName, ElementFieldLabel, VisioFieldType, ElementFieldDescription
            
        Else
            CurrentProperty.Init ElementFieldName, ElementFieldLabel, VisioFieldType, ElementFieldDescription, FieldDefinitionVO.GetListDomain
        
        End If
        
       
        DesignOrder = ElementDefinitionFieldVO.ShowDesignOrder
        DesignParenthesis = ElementDefinitionFieldVO.ShowDesignParenthesis
            
        CommentOrder = ElementDefinitionFieldVO.ShowCommentsOrder
        CommentParenthesis = ElementDefinitionFieldVO.ShowCommentsParenthesis
        
        ElementFieldMandatory = ElementDefinitionFieldVO.FieldMandatory
        
        CurrentProperty.InitDesignData DesignOrder, DesignParenthesis, CommentOrder, CommentParenthesis
        
        InitCustomPropertiesValidationData elementType, _
                                           ElementStereotype, _
                                           ElementFieldMandatory, _
                                           ElementFieldName, _
                                           ElementFieldType, _
                                           CurrentProperty

        
        Result.AddCustomProperty CurrentProperty
       
    Next
         
    Result.AddCustomProperty CustomPropertyURLInfo

    Set GetCustomProperties = Result

End Function

Private Sub InitCustomPropertiesValidationData(ByVal elementType As String, _
                                               ByVal ElementStereotype As String, _
                                               ByVal ElementFieldMandatory As String, _
                                               ByVal ElementFieldName As String, _
                                               ByVal ElementFieldType As String, _
                                               ByRef CustomPropertyReferenced As CustomProperty)

    Dim CustomPropertyValidator As IValidator

    
    On Error GoTo ErrorGettingValidationData
    
        'If ElementFieldMandatory Then
        
            If ElementFieldType = sysAdlFieldTypeString Then
            
                Set CustomPropertyValidator = GetValidatorString(elementType, ElementStereotype, ElementFieldName)
            
            ElseIf ElementFieldType = sysAdlFieldTypeDate Then
            
                Set CustomPropertyValidator = GetValidatorDate(ElementFieldName)
                
            ElseIf ElementFieldType = sysAdlFieldTypeElement Then
            
                Set CustomPropertyValidator = GetValidatorElement(elementType, ElementStereotype, ElementFieldName)
                
            ElseIf ElementFieldType = sysAdlFieldTypeList Then
            
                Set CustomPropertyValidator = GetValidatorList(elementType, ElementStereotype, ElementFieldName)
                
            ElseIf ElementFieldType = sysAdlFieldTypeTime Then
            
                Set CustomPropertyValidator = GetValidatorTime(elementType, ElementStereotype, ElementFieldName)
                
            ElseIf ElementFieldType = sysAdlFieldTypeValue Then
            
                Set CustomPropertyValidator = GetValidatorValue(elementType, ElementStereotype, ElementFieldName)
            
            End If
            
        'Else
        
         '   Set CustomPropertyValidator = New ValidatorNotMandatory
            
        'End If
        
        CustomPropertyReferenced.InitValidator CustomPropertyValidator, ElementFieldMandatory
        
    
    
    Exit Sub

ErrorGettingValidationData:

    Dim errorHandleMessage As String
    
    GUIServices.LogMessage ("Error: " + Err.Description)
    
     errorHandleMessage = "An error occurred while getting validation data for " + vbCrLf + vbCrLf + _
                          "- SysADL Type: " + elementType + vbCrLf + _
                          "- Stereotype: " + ElementStereotype + vbCrLf + _
                          "- Field Name: " + ElementFieldName + vbCrLf + vbCrLf

    If Err.Number = 9000 Then
    
        errorHandleMessage = errorHandleMessage + Err.Description

    Else
        errorHandleMessage = errorHandleMessage + _
                             "It was not possible to get validation data from configuration database. " + _
                             "Please, check data validation configuration."
                             
    End If
    
    Err.Raise 9000, "", errorHandleMessage
    
    Exit Sub


End Sub



Private Sub InitializeCustomProperties()
        
    If Not IsCustomPropertiesInitialized Then
        
        InitializeCustomPropertiesForKeyProperties

        IsCustomPropertiesInitialized = True
    End If

End Sub

Private Sub InitializeCustomPropertiesForKeyProperties()
    CustomPropertyBaseId.Init sysAdlKeyCustPropRowNameId, sysAdlKeyCustPropLabelId, visPropTypeString, "Enter the identifier of the element"
    CustomPropertyBaseId.InitValidator VALIDATOR_ID, FIELD_IS_MANDATORY
    CustomPropertyBaseNamespace.Init sysAdlKeyCustPropRowNameNamespace, sysAdlKeyCustPropLabelNamespace, visPropTypeString, "Enter the namespace of the element"
    CustomPropertyBaseNamespace.InitValidator VALIDATOR_NAMESPACE, FIELD_IS_MANDATORY
    CustomPropertyURLInfo.Init sysAdlKeyCustPropRowNameURLInfo, sysAdlKeyCustPropLabelURLInfo, visPropTypeString, "Enter the URL to access a page that has further information about this element."
    CustomPropertyURLInfo.InitValidator VALIDATOR_URLINFO, FIELD_IS_NOT_MANDATORY
End Sub




Private Function CreateStereotypeCustomProperty(ByVal ListKey As String) As CustomProperty
    Dim Result As New CustomProperty
    Dim StereotypeList As Collection
    
    Set StereotypeList = GetStereotypeList(ListKey)
    
    If StereotypeList.Count < 1 Then
    
        Set StereotypeList = New Collection
        StereotypeList.Add sysAdlStringConstantsEmpty
        StereotypeList.Add sysAdlNoStereotype
        
    End If
    
    Result.Init sysAdlKeyCustPropRowNameStereotype, sysAdlKeyCustPropLabelStereotype, visPropTypeListFix, "Enter the stereotype of the element", StereotypeList
    Result.InitValidator VALIDATOR_STEREOTYPE, FIELD_IS_MANDATORY

    Set CreateStereotypeCustomProperty = Result
End Function

Private Function GetStereotypeList(ByVal sysadlType As String) As Collection

    Dim Result As Collection

    Set Result = FactoryDefinitionElement.GetStereotypeList(sysadlType)
    
    Set GetStereotypeList = Result

End Function



Private Function GetValidatorDate(ByVal ElementFieldName As String) As IValidator
                                  
    Dim Result As ValidatorDate
    Dim ElementFieldMessageError As String
    Dim ElementFieldDateAllowPast As Boolean
    Dim ElementFieldDateAllowPresent As Boolean
    Dim ElementFieldDateAllowFuture As Boolean
    
    Dim FieldDefinitionVO As VOFieldDefinition
    
    Set FieldDefinitionVO = FactoryDefinitionField.GetFieldDefinition(ElementFieldName)

    Set Result = New ValidatorDate
    
        ElementFieldMessageError = FieldDefinitionVO.ErrorMessage
        ElementFieldDateAllowPast = FieldDefinitionVO.DateAllowPast
        ElementFieldDateAllowPresent = FieldDefinitionVO.DateAllowPresent
        ElementFieldDateAllowFuture = FieldDefinitionVO.DateAllowFuture
    
        Result.Init ElementFieldDateAllowPast, _
                    ElementFieldDateAllowPresent, _
                    ElementFieldDateAllowFuture, _
                    ElementFieldMessageError
    
    Set GetValidatorDate = Result
                                  
End Function
Private Function GetValidatorElement(ByVal elementType As String, _
                                     ByVal ElementStereotype As String, _
                                     ByVal ElementFieldName As String) As IValidator
                                  
    Dim Result As ValidatorElement
    Dim ElementFieldMessageError As String
    Dim ElementFieldElementType As String
    
    Dim FieldDefinitionVO As VOFieldDefinition
    
    Set FieldDefinitionVO = FactoryDefinitionField.GetFieldDefinition(ElementFieldName)
    
    Set Result = New ValidatorElement
    
        ElementFieldMessageError = FieldDefinitionVO.ErrorMessage
        ElementFieldElementType = FieldDefinitionVO.elementType
    
        Result.Init ElementFieldElementType, ElementFieldMessageError
    
    
    Set GetValidatorElement = Result
                                  
End Function
                                  
Private Function GetValidatorString(ByVal elementType As String, _
                                     ByVal ElementStereotype As String, _
                                     ByVal ElementFieldName As String) As IValidator
                                  
    Dim Result As ValidatorString
    Dim ElementFieldMessageError As String
    Dim ElementFieldStringRegExp As String
    
    Dim FieldDefinitionVO As VOFieldDefinition
    
    Set FieldDefinitionVO = FactoryDefinitionField.GetFieldDefinition(ElementFieldName)
    
    Set Result = New ValidatorString
    
    
        ElementFieldMessageError = FieldDefinitionVO.ErrorMessage
        ElementFieldStringRegExp = FieldDefinitionVO.StringRegExp
    
        Result.Init ElementFieldStringRegExp, ElementFieldMessageError
    
    
    Set GetValidatorString = Result
                                  
End Function

Private Function GetValidatorTime(ByVal elementType As String, _
                                     ByVal ElementStereotype As String, _
                                     ByVal ElementFieldName As String) As IValidator
                                  
    Dim Result As ValidatorTime
    Dim ElementFieldMessageError As String
    Dim ElementFieldTimeMinimum As Date
    Dim ElementFieldTimeMaximum As Date
    
    Dim FieldDefinitionVO As VOFieldDefinition
    
    Set FieldDefinitionVO = FactoryDefinitionField.GetFieldDefinition(ElementFieldName)
    
    Set Result = New ValidatorTime
    
        ElementFieldMessageError = FieldDefinitionVO.ErrorMessage
        ElementFieldTimeMinimum = FieldDefinitionVO.TimeMinimum
        ElementFieldTimeMaximum = FieldDefinitionVO.TimeMaximum
    
        Result.Init ElementFieldTimeMinimum, ElementFieldTimeMaximum, ElementFieldMessageError
    
    
    Set GetValidatorTime = Result
                                  
End Function

Private Function GetValidatorValue(ByVal elementType As String, _
                                     ByVal ElementStereotype As String, _
                                     ByVal ElementFieldName As String) As IValidator
                                  
    Dim Result As ValidatorValue
    Dim ElementFieldMessageError As String
    Dim ElementFieldValueMinimum As Double
    Dim ElementFieldValueMaximum As Double
    Dim ElementFieldValueOnlyInteger As Boolean

    Dim FieldDefinitionVO As VOFieldDefinition
    
    Set FieldDefinitionVO = FactoryDefinitionField.GetFieldDefinition(ElementFieldName)
    
    Set Result = New ValidatorValue
    
        ElementFieldMessageError = FieldDefinitionVO.ErrorMessage
        ElementFieldValueMinimum = FieldDefinitionVO.ValueMinimum
        ElementFieldValueMaximum = FieldDefinitionVO.ValueMaximum
        ElementFieldValueOnlyInteger = FieldDefinitionVO.ValueOnlyInteger
    
        Result.Init ElementFieldValueMinimum, _
                    ElementFieldValueMaximum, _
                    ElementFieldValueOnlyInteger, _
                    ElementFieldMessageError
    
    
    Set GetValidatorValue = Result
                                  
End Function

Private Function GetValidatorList(ByVal elementType As String, _
                                     ByVal ElementStereotype As String, _
                                     ByVal ElementFieldName As String) As IValidator
                                  
    Dim Result As ValidatorList
    Dim ElementFieldMessageError As String
    Dim ElementFieldStringRegExp As String
    
    Dim FieldDefinitionVO As VOFieldDefinition
    
    Set FieldDefinitionVO = FactoryDefinitionField.GetFieldDefinition(ElementFieldName)
    
    Set Result = New ValidatorList
    
    ElementFieldMessageError = FieldDefinitionVO.ErrorMessage
    
    Result.Init ElementFieldMessageError
    
    Set GetValidatorList = Result
                                  
End Function


