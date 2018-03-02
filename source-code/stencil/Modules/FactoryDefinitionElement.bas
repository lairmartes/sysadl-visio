Attribute VB_Name = "FactoryDefinitionElement"

Private pChannelDefinitionList As Collection
Private pFormatDefinitionList As Collection
Private pLayerDefinitionList As Collection
Private pNodeDefinitionList As Collection
Private pObjectiveDefinitionList As Collection
Private pQualityDefinitionList As Collection
Private pReceiverDefinitionList As Collection
Private pSenderDefinitionList As Collection
Private pStakeholderDefinitionList As Collection
Private pSystemDefinitionList As Collection
Private pTransitionDefinitionList As Collection
Private pRoleDefinitionList As Collection
Private pDecisionDefinitionList As Collection
Private pAssumptionDefinitionList As Collection

Private Const sysAdlAlwaysRefreshDefinitions = "True"



Private Sub ImportElementDefinition(ByVal aSysAdlType As String)

    Dim FilePath As String
    Dim ElementDefinitionXML As New DOMDocument
    
    Dim TagDefinitions As IXMLDOMElement
    Dim NodeDefinitions As IXMLDOMNodeList
    Dim TagDefinition As IXMLDOMElement
    Dim NodeDefinition As IXMLDOMNodeList
    Dim NodeDefitinionData As IXMLDOMNodeList
    Dim TagDefinitionData As IXMLDOMElement
    Dim TagDefinitionPreferences As IXMLDOMElement
    
    Dim ElementDefinitionVO As VOElementDefinition
    Dim ElementDefinitionFieldVO As VOElementDefinitionField
    
    Dim Stereotype As String
    Dim IsDeprecated As Boolean
    Dim FileSysAdlTypeSufix As String
    
    Dim FieldName As String
    Dim FieldMandatory As String
    Dim FieldOrder As String


    'initialize list of definitions
    
    InitDefinitionList aSysAdlType
    
    FileSysAdlTypeSufix = StrConv(aSysAdlType, vbLowerCase)
    
    FilePath = FactoryConfigProperty.GetProperty(sysAdlConfigPropertyElementDefinitionPath) + _
               "element-definition-" + _
                FileSysAdlTypeSufix + _
                ".xml"
    
    ElementDefinitionXML.async = False
    
    ElementDefinitionXML.Load (FilePath)
    
    If ElementDefinitionXML.parseError.reason <> sysAdlStringConstantsEmpty Then
    
        'Err.Raise 7000, "Importing definitions", ElementDefinitionXML.parseError.reason
        Err.Raise 7001, _
                  "Factory Definition Element", _
                  "The file " + FilePath + " that contains important definitions that are required by this tool was not found. Provide this file to use Sys-ADL for Visio."
        
    End If
    
    On Error GoTo HandleErrorXML
    
    Set NodeDefinitions = ElementDefinitionXML.ChildNodes
    
    Set TagDefinitions = NodeDefinitions.NextNode
    
    Set NodeDefinition = TagDefinitions.ChildNodes
    
    For I = 0 To NodeDefinition.Length - 1
    
        Set TagDefinition = NodeDefinition.Item(I)
        
        Set ElementDefinitionVO = New VOElementDefinition
        
        Stereotype = TagDefinition.getAttribute("stereotype")
        IsDeprecated = CBool(TagDefinition.getAttribute("deprecated"))
        
        ElementDefinitionVO.Init aSysAdlType, Stereotype, IsDeprecated
        
        Set NodeDefinitionData = TagDefinition.ChildNodes
    
        For J = 0 To NodeDefinitionData.Length - 1
        
            Set ElementDefinitionFieldVO = New VOElementDefinitionField
            
            Set TagDefinitionData = NodeDefinitionData.Item(J)
            
            FieldName = TagDefinitionData.getAttribute("name")
            FieldMandatory = TagDefinitionData.getAttribute("mandatory")
            FieldOrder = TagDefinitionData.getAttribute("order")
            
            ElementDefinitionFieldVO.Init FieldName, CBool(FieldMandatory), CDbl(FieldOrder)
            
            InitPreferences ElementDefinitionFieldVO, TagDefinitionData
            
            ElementDefinitionVO.AddField ElementDefinitionFieldVO
            
        Next
        
        AddDefinitionListItem aSysAdlType, ElementDefinitionVO
        
    Next
    
    Exit Sub
    
HandleErrorXML:

    Err.Raise 7001, "Factory Definition Element", "The file " + FilePath + " that contains the element fields definition is required. Please, provide this file to use Sys-ADL for Visio."

End Sub

Private Sub InitPreferences(ByRef ElementDefinitionFieldVO As VOElementDefinitionField, _
                            ByVal TagDefinitionPreferences As IXMLDOMElement)
                            
    
    Dim FieldPreferencesShowDesignOrder As String
    Dim FieldPreferencesShowDesignParenthesis As String
    
    Dim FieldPreferencesShowCommentsOrder As String
    Dim FieldPreferencesShowCommentsParenthesis As String
    
    Dim PreferenceTag As IXMLDOMElement
    Dim PreferenceNodes As IXMLDOMNodeList
    
    Set PreferenceNodes = TagDefinitionPreferences.ChildNodes
    
    For I = 0 To PreferenceNodes.Length - 1
    
        Set PreferenceTag = PreferenceNodes.Item(I)
                            
            If PreferenceTag.nodeName = "show-design" Then
            
                FieldPreferencesShowDesignOrder = PreferenceTag.getAttribute("order")
                FieldPreferencesShowDesignParenthesis = PreferenceTag.getAttribute("parenthesis")
                
                ElementDefinitionFieldVO.InitDesignPreferences CDbl(FieldPreferencesShowDesignOrder), _
                                                               CBool(FieldPreferencesShowDesignParenthesis)
                
            ElseIf PreferenceTag.nodeName = "show-comments" Then
            
                FieldPreferencesShowCommentsOrder = PreferenceTag.getAttribute("order")
                FieldPreferencesShowCommentsParenthesis = PreferenceTag.getAttribute("parenthesis")
                
                ElementDefinitionFieldVO.InitCommentsPreferences CDbl(FieldPreferencesShowCommentsOrder), _
                                                                 CBool(FieldPreferencesShowCommentsParenthesis)
                
            End If
            
    Next

End Sub

Private Sub InitDefinitionList(ByVal aSysAdlType)

    Select Case aSysAdlType
    
        Case sysAdlTypeSetChannel: Set pChannelDefinitionList = New Collection
        Case sysAdlTypeSetFormat: Set pFormatDefinitionList = New Collection
        Case sysAdlTypeSetLayer: Set pLayerDefinitionList = New Collection
        Case sysAdlTypeSetNode: Set pNodeDefinitionList = New Collection
        Case sysAdlTypeSetObjective: Set pObjectiveDefinitionList = New Collection
        Case sysAdlTypeSetQuality: Set pQualityDefinitionList = New Collection
        Case sysAdlTypeSetReceiver: Set pReceiverDefinitionList = New Collection
        Case sysAdlTypeSetSender: Set pSenderDefinitionList = New Collection
        Case sysAdlTypeSetStakeholder: Set pStakeholderDefinitionList = New Collection
        Case sysAdlTypeSetSystem: Set pSystemDefinitionList = New Collection
        Case sysAdlTypeSetTransition: Set pTransitionDefinitionList = New Collection
        Case sysAdlTypeSetRole: Set pRoleDefinitionList = New Collection
        Case sysAdlTypeSetDecision: Set pDecisionDefinitionList = New Collection
        Case sysAdlTypeSetAssumption: Set pAssumptionDefinitionList = New Collection

    
    End Select
        
End Sub

Private Sub AddDefinitionListItem(ByVal aSysAdlType As String, ByVal aDefinitionItem As VOElementDefinition)

    Select Case aSysAdlType
    
        Case sysAdlTypeSetChannel: pChannelDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetFormat: pFormatDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetLayer: pLayerDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetNode: pNodeDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetObjective: pObjectiveDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetQuality: pQualityDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetReceiver: pReceiverDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetSender: pSenderDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetStakeholder: pStakeholderDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetSystem: pSystemDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetTransition: pTransitionDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetRole: pRoleDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetDecision: pDecisionDefinitionList.Add aDefinitionItem
        Case sysAdlTypeSetAssumption: pAssumptionDefinitionList.Add aDefinitionItem
    
    End Select
        
End Sub

Public Function GetElementDefinition(ByVal aSysAdlType, ByVal aStereotype) As VOElementDefinition

    Dim Result As VOElementDefinition
    Dim ListToSeach As Collection
    
    Set ListToSeach = GetElementDefinitionList(aSysAdlType)

    For Each Result In ListToSeach
    
        If Result.Stereotype = aStereotype Then Exit For
        
    Next
    
    
    Set GetElementDefinition = Result

End Function

Public Function GetElementDefinitionField(ByVal aSysAdlType, ByVal aStereotype, ByVal aFieldName) As VOElementDefinitionField

    Dim Result As VOElementDefinitionField
    Dim ListToSeach As Collection
    Dim ElementDefinitionVO As VOElementDefinition
    Dim FieldList As Collection

    Set ListToSeach = GetElementDefinitionList(aSysAdlType)

    For Each ElementDefinitionVO In ListToSeach
    
        If ElementDefinitionVO.Stereotype = aStereotype Then
        
            Set FieldList = ElementDefinitionVO.GetElementFieldList
        
            For Each Result In FieldList
            
                If Result.FieldName = aFieldName Then Exit For
                
            Next
        
        End If
        
    Next
    
    Set GetElementDefinitionField = Result

End Function

Public Function GetStereotypeList(ByVal aSysAdlType) As Collection

    Dim Result As New Collection
    Dim ListToSeach As Collection
    Dim ElementDefinitionFound As VOElementDefinition

    Set ListToSeach = GetElementDefinitionList(aSysAdlType)

    For Each ElementDefinitionFound In ListToSeach
    
        If (ElementDefinitionFound.IsDeprecated = False) Then
            
            Result.Add ElementDefinitionFound.Stereotype
        
        End If
    Next
    
    Set GetStereotypeList = Result

End Function

Private Function GetElementDefinitionList(ByVal aSysAdlType) As Collection

    Dim Result As Collection
    Dim RefreshPolicy As String

    RefreshPolicy = FactoryConfigProperty.GetProperty(sysAdlConfigPropertyElementDefinitionAlwaysRefresh)
    
    DefinitionMustBeRefreshed = (RefreshPolicy = sysAdlAlwaysRefreshDefinitions)
    
    Set Result = GetListByType(aSysAdlType)
    
    If ((Result Is Nothing) Or DefinitionMustBeRefreshed) Then
    
        ImportElementDefinition aSysAdlType
        
        Set Result = GetListByType(aSysAdlType)
        
    End If
    
    Set GetElementDefinitionList = Result

End Function

Private Function GetListByType(ByVal aSysAdlType) As Collection

    Dim Result As Collection

    Select Case aSysAdlType
    
        Case sysAdlTypeSetChannel: Set Result = pChannelDefinitionList
        Case sysAdlTypeSetFormat: Set Result = pFormatDefinitionList
        Case sysAdlTypeSetLayer: Set Result = pLayerDefinitionList
        Case sysAdlTypeSetNode: Set Result = pNodeDefinitionList
        Case sysAdlTypeSetObjective: Set Result = pObjectiveDefinitionList
        Case sysAdlTypeSetQuality: Set Result = pQualityDefinitionList
        Case sysAdlTypeSetReceiver: Set Result = pReceiverDefinitionList
        Case sysAdlTypeSetSender: Set Result = pSenderDefinitionList
        Case sysAdlTypeSetStakeholder: Set Result = pStakeholderDefinitionList
        Case sysAdlTypeSetSystem: Set Result = pSystemDefinitionList
        Case sysAdlTypeSetTransition: Set Result = pTransitionDefinitionList
        Case sysAdlTypeSetRole: Set Result = pRoleDefinitionList
        Case sysAdlTypeSetDecision: Set Result = pDecisionDefinitionList
        Case sysAdlTypeSetAssumption: Set Result = pAssumptionDefinitionList
        
    End Select
    
    Set GetListByType = Result
    
End Function
