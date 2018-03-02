Attribute VB_Name = "FactoryDefinitionField"


Private pFieldDefinitionList As Collection
Private Const sysAdlAlwaysRefreshDefinitions = "True"
Private Const DATE_1950_1_1 = "18264"
Private Const MINIMUM_DEFAULT_DOUBLE = "-1E-200"
Private Const MAXIMUM_DEFAULT_DOBULE = "1E200"

Private Sub ImportFieldDefinition()

    Dim FilePath As String
    Dim FieldDefinitionXML As New DOMDocument
    
    Dim TagFields As IXMLDOMElement
    Dim NodeFields As IXMLDOMNodeList
    Dim TagField As IXMLDOMElement
    Dim NodeField As IXMLDOMNodeList
    Dim NodeFieldData As IXMLDOMNodeList
    Dim TagFieldData As IXMLDOMElement
    Dim TagFieldTypeData As IXMLDOMElement
    
    Dim FieldDefinitionVO As VOFieldDefinition
    
    Dim FieldName As String
    Dim FieldLabel As String
    Dim FieldDescription As String
    Dim FieldErrorMessage As String
    Dim FieldType As String
    
    'initialize list of fields
    
    Set pFieldDefinitionList = New Collection
    
    FilePath = FactoryConfigProperty.GetProperty(sysAdlConfigPropertyElementDefinitionPath) + "element-fields.xml"
    
    FieldDefinitionXML.async = False
    
    FieldDefinitionXML.Load (FilePath)
    
    If FieldDefinitionXML.parseError.reason <> sysAdlStringConstantsEmpty Then
    
        'Err.Raise 7000, "Importing definitions", FieldDefinitionXML.parseError.reason
        Err.Raise 7001, _
                  "Factory Definition Field", _
                  "The file " + FilePath + " that contains data about fields and domains is not available. Provide this file to use Sys-ADL in Visio."
        
    End If
    
    On Error GoTo HandleErrorXML
    
    Set NodeFields = FieldDefinitionXML.ChildNodes
    
    Set TagFields = NodeFields.NextNode
    
    Set NodeField = TagFields.ChildNodes
    
    For I = 0 To NodeField.Length - 1
    
        Set TagField = NodeField.Item(I)
        
        Set NodeFieldData = TagField.ChildNodes
        
        FieldName = TagField.getAttribute("name")
    
        For J = 0 To NodeFieldData.Length - 1
        
            Set TagFieldData = NodeFieldData.Item(J)
                
            If TagFieldData.nodeName = "label" Then
            
                FieldLabel = TagFieldData.nodeTypedValue
                
            ElseIf TagFieldData.nodeName = "description" Then
            
                FieldDescription = TagFieldData.nodeTypedValue
                
            ElseIf TagFieldData.nodeName = "error-message" Then
            
                FieldErrorMessage = TagFieldData.nodeTypedValue
                
            ElseIf TagFieldData.nodeName = "type" Then
            
                FieldType = TagFieldData.getAttribute("value")
                
                Set TagFieldTypeData = TagFieldData
                
            End If
            
        Next
        
        Set FieldDefinitionVO = New VOFieldDefinition
            
        FieldDefinitionVO.Init FieldName, FieldLabel, FieldDescription, FieldErrorMessage, FieldType
            
        If FieldType = "String" Then
            
            ImportFieldDefinitionTypeString FieldDefinitionVO, TagFieldTypeData
                
        ElseIf FieldType = "List" Then
                
            ImportFieldDefinitionTypeList FieldDefinitionVO, TagFieldTypeData
                
        ElseIf FieldType = "Date" Then
            
            ImportFieldDefinitionTypeDate FieldDefinitionVO, TagFieldTypeData
                
        ElseIf FieldType = "Time" Then
            
            ImportFieldDefinitionTypeTime FieldDefinitionVO, TagFieldTypeData
                
        ElseIf FieldType = "Value" Then
            
            ImportFieldDefinitionTypeValue FieldDefinitionVO, TagFieldTypeData
                
        ElseIf FieldType = "Element" Then
            
            ImportFieldDefinitionTypeElement FieldDefinitionVO, TagFieldTypeData
                
        End If
            
        pFieldDefinitionList.Add FieldDefinitionVO
        
    Next
    
    Exit Sub
    
HandleErrorXML:

    Err.Raise 7001, "Factory Definition Field", "The file " + FilePath + " that contains data about fields and domains is not available. This file is required to use Sys-ADL in Visio."

End Sub

Private Sub ImportFieldDefinitionTypeString(ByRef FieldVO As VOFieldDefinition, _
                                            ByVal FieldXML As IXMLDOMElement)
                                            
    Dim FieldTypeStringRegExp As String
    Dim StringTypeTag As IXMLDOMElement
    
    Set StringTypeTag = FieldXML.ChildNodes.NextNode
    
    FieldTypeStringRegExp = StringTypeTag.nodeTypedValue
    
    FieldVO.InitTypeString FieldTypeStringRegExp
                                            
End Sub

Private Sub ImportFieldDefinitionTypeElement(ByRef FieldVO As VOFieldDefinition, _
                                            ByVal FieldXML As IXMLDOMElement)
                                            
    Dim FieldTypeElement As String
    Dim ElementTypeTag As IXMLDOMElement
    
    Set ElementTypeTag = FieldXML.ChildNodes.NextNode
    
    FieldTypeElement = ElementTypeTag.nodeTypedValue
    
    FieldVO.InitTypeElement FieldTypeElement
                                            
End Sub

Private Sub ImportFieldDefinitionTypeList(ByRef FieldVO As VOFieldDefinition, _
                                            ByVal FieldXML As IXMLDOMElement)
                                            
    Dim FieldTypeListItem As String
    Dim ListTypeTag As IXMLDOMElement
    Dim ListTypeNodes As IXMLDOMNodeList
    
    Set ListTypeNodes = FieldXML.ChildNodes
    
    For I = 0 To ListTypeNodes.Length - 1
    
        Set ListTypeTag = ListTypeNodes.Item(I)
    
        FieldTypeListItem = ListTypeTag.nodeTypedValue
    
        FieldVO.AddListDomainItem FieldTypeListItem
        
    Next
                                            
End Sub

Private Sub ImportFieldDefinitionTypeTime(ByRef FieldVO As VOFieldDefinition, _
                                            ByVal FieldXML As IXMLDOMElement)
                                            
    Dim TimeMinimum As String
    Dim TimeMaximum As String
    Dim TimeTypeTag As IXMLDOMElement
    Dim TimeTypeNodes As IXMLDOMNodeList
    
    TimeMinimum = "00:00"
    TimeMaximum = "23:59:59"
    
    Set TimeTypeNodes = FieldXML.ChildNodes
    
    For I = 0 To TimeTypeNodes.Length - 1
    
        Set TimeTypeTag = TimeTypeNodes.Item(I)
        
        If TimeTypeTag.nodeName = "minimum" Then
        
            TimeMinimum = TimeTypeTag.nodeTypedValue
            
        ElseIf TimeTypeTag.nodeName = "maximum" Then
        
            TimeMaximum = TimeTypeTag.nodeTypedValue

        End If
        
    Next
    
    FieldVO.InitTypeTime CDate(TimeMinimum), CDate(TimeMaximum)
                                            
End Sub

Private Sub ImportFieldDefinitionTypeDate(ByRef FieldVO As VOFieldDefinition, _
                                            ByVal FieldXML As IXMLDOMElement)
                                            
    Dim AllowPast As String
    Dim AllowPresent As String
    Dim AllowFuture As String
    Dim DateTypeTag As IXMLDOMElement
    Dim DateTypeNodes As IXMLDOMNodeList
    
    AllowPast = "False"
    AllowPresent = "False"
    AllowFuture = "False"
    
    Set DateTypeNodes = FieldXML.ChildNodes
    
    For I = 0 To DateTypeNodes.Length - 1
    
        Set DateTypeTag = DateTypeNodes.Item(I)
        
        If DateTypeTag.nodeName = "date-allow-past" Then
        
            AllowPast = DateTypeTag.nodeTypedValue
            
        ElseIf DateTypeTag.nodeName = "date-allow-present" Then
        
            AllowPresent = DateTypeTag.nodeTypedValue
            
        ElseIf DateTypeTag.nodeName = "date-allow-future" Then
        
            AllowFuture = DateTypeTag.nodeTypedValue

        End If
        
    Next
    
    FieldVO.InitTypeDate CBool(AllowPast), CBool(AllowPresent), CBool(AllowFuture)
                                            
End Sub

Private Sub ImportFieldDefinitionTypeValue(ByRef FieldVO As VOFieldDefinition, _
                                            ByVal FieldXML As IXMLDOMElement)
                                            
    Dim ValueMinimum As String
    Dim ValueMaximum As String
    Dim ValueOnlyInteger As String
    Dim ValueTypeTag As IXMLDOMElement
    Dim ValueTypeNodes As IXMLDOMNodeList
    
    Set ValueTypeNodes = FieldXML.ChildNodes
    
    ValueMinimum = MINIMUM_DEFAULT_DOUBLE
    ValueMaximum = MAXIMUM_DEFAULT_DOBULE
    ValueOnlyInteger = "False"
    
    For I = 0 To ValueTypeNodes.Length - 1
    
        Set ValueTypeTag = ValueTypeNodes.Item(I)
        
        If ValueTypeTag.nodeName = "minimum" Then
        
            ValueMinimum = ValueTypeTag.nodeTypedValue
            
        ElseIf ValueTypeTag.nodeName = "maximum" Then
        
            ValueMaximum = ValueTypeTag.nodeTypedValue
            
        ElseIf ValueTypeTag.nodeName = "only-integer" Then
        
            ValueOnlyInteger = ValueTypeTag.nodeTypedValue

        End If
        
    Next
    
    FieldVO.InitTypeValue CDbl(ValueMinimum), CDbl(ValueMaximum), CBool(ValueOnlyInteger)
                                            
End Sub

Public Function GetFieldDefinition(ByVal aFieldName) As VOFieldDefinition

    Dim NameFound As String
    Dim Result As VOFieldDefinition
    Dim DefinitionMustBeRefreshed
    Dim RefreshPolicy As String
    
    RefreshPolicy = FactoryConfigProperty.GetProperty(sysAdlConfigPropertyElementDefinitionAlwaysRefresh)
    
    DefinitionMustBeRefreshed = (RefreshPolicy = sysAdlAlwaysRefreshDefinitions)
    
    If ((pFieldDefinitionList Is Nothing) Or DefinitionMustBeRefreshed) Then
    
        ImportFieldDefinition
        
    End If

    For Each Result In pFieldDefinitionList
    
        If Result.Name = aFieldName Then Exit For
        
    Next
    
    If Result Is Nothing Then
    
        Err.Raise 7002, "Factory Definition Field", "Field named " + aFieldName + " was not found.  Check in definition files if ELEMENT FIELD DEFINED and FIELD DEFINITION names match"
        
    End If
    
    Set GetFieldDefinition = Result

End Function

