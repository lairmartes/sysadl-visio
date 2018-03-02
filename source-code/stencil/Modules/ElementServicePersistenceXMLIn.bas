Attribute VB_Name = "ElementServicePersistenceXMLIn"

Public Function OpenSysADLElement(ByVal anElement As SysAdlElement) As SysAdlElement


    'declare path to reach sadl document
    Dim SADLPath As String
    'declare XML holder
    Dim SADLDocument As New DOMDocument
    
    Dim IsDocumentCannotBeOpened As Boolean
    
    Dim Result As SysAdlElement
    
    GUIServices.LogMessage ">>>>>>>>>>>>>>>>>>>> IMPORTING ELEMENT <<<<<<<<<<<<<<<<<<<<<<<<<<<<"

    GUIServices.LogMessage "Importing element: " + anElement.namespace + "." + anElement.Id

    'convert namespace and diagram name to get SADL file path
    SADLPath = GUIServices.GetFullFileNameForElement(anElement.namespace, anElement.Id)
    
    GUIServices.LogMessage "Importing Element File: " + SADLPath
    
    'configure DOM Object
    SADLDocument.async = False
    
    'Load file
    Call SADLDocument.Load(SADLPath)
    
    'Treat error (if any)
    IsDocumentCannotBeOpened = SADLDocument.parseError.reason <> sysAdlStringConstantsEmpty
    
    If IsDocumentCannotBeOpened Then
    
        GUIServices.LogMessage "An error has occurred while loading file: " + SADLDocument.parseError.reason
        
        GUIServices.LogMessage "Error line  : " + Conversion.CStr(SADLDocument.parseError.Line)
        GUIServices.LogMessage "Error column: " + Conversion.CStr(SADLDocument.parseError.linepos)
        GUIServices.LogMessage "Text error  : " + SADLDocument.parseError.srcText
        
    Else
    
        'Import data of diagram
        Set Result = New SysAdlElement
        RecoverFileContent SADLDocument, Result
    
    End If
    
    Set OpenSysADLElement = Result

End Function

Public Function IsElementExists(ByVal anElement As SysAdlElement)

    Dim Result As Boolean
    Dim ElementFound As SysAdlElement
    
    Result = True
    
    Set ElementFound = OpenSysADLElement(anElement)
    
    If (ElementFound Is Nothing) Then
    
        Result = False
        
    End If
    
    IsElementExists = Result

End Function

Public Function IsElementExistsWithType(ByVal element As SysAdlElement) As Boolean

    Dim Result As Boolean
    Dim ElementFound As SysAdlElement
    Dim ElementTypeReceived As String
    
    ElementTypeReceived = element.sysadlType
    
    Result = False
    
    Set ElementFound = OpenSysADLElement(anElement)
    
    If Not (ElementFound Is Nothing) Then
    
        If ElementFound.sysadlType = ElementTypeReceived Then
        
            Result = True
            
        End If
        
    End If
    
    IsElementExistsWithType = Result

End Function
Private Sub RecoverFileContent(ByVal SADLDocument As DOMDocument, ByRef sysadlElementInFile As SysAdlElement)

    ' Declare children list for <elements> tag. Expected is <element>
    Dim NodeElement As IXMLDOMNodeList
    ' declare tag to hold <element> data
    Dim TagElementData As IXMLDOMElement
    Dim TagElementGroup As IXMLDOMElement
    Dim TagElement As IXMLDOMElement

    ' get children from whole diagram, in this case is <diagram>
    Set NodeElement = SADLDocument.ChildNodes

    Set TagElement = NodeElement.NextNode
    
    ImportElement TagElement, sysadlElementInFile

End Sub

Private Sub ImportElement(ByVal TagElement As IXMLDOMElement, ByRef sysadlElementInFile As SysAdlElement)

    'declare variables to hold xml data
    Dim elementType As String
    Dim ElementStereotype As String
    Dim elementNamespace As String
    Dim elementId As String
    Dim ElementURLInfo As String
    
    'declare children of TagElement
    Dim NodeElementItems As IXMLDOMNodeList
    'declare tag to hold item data
    Dim TagItem As IXMLDOMElement
    
    'get element basic data
    elementType = TagElement.getAttribute("type")
    ElementStereotype = TagElement.getAttribute("stereotype")
    elementNamespace = TagElement.getAttribute("namespace")
    elementId = TagElement.getAttribute("id")
    ElementURLInfo = GUIServices.PredefinedXMLToField(TagElement.getAttribute("url-info"))
    
    GUIServices.LogMessage " --------------------- ELEMENT -------------------------------"
    
    sysadlElementInFile.Init elementType, ElementStereotype
    sysadlElementInFile.InitBaseData elementNamespace, elementId, ElementURLInfo
    
    'get element items.  <attribute> is expected
    Set NodeElementItems = TagElement.ChildNodes
    
    GUIServices.LogMessage "    ............... atributes ................."
    
    'iterate withing element items
    For I = 0 To NodeElementItems.Length - 1
    
        Set TagItem = NodeElementItems.Item(I)
        
        'import element attribute (tag <attribute>)
        If sysAdlTagSectionElementAttribute = TagItem.nodeName Then
        
            ImportAttribute elementNamespace, elementId, TagItem, sysadlElementInFile

        End If
    
    Next


End Sub

Private Sub ImportAttribute(ByVal elementNamespace, _
                             ByVal elementId, _
                             ByVal TagAttribute As IXMLDOMElement, _
                             ByRef sysadlElementInFile As SysAdlElement)

    Dim ElementAttributeName As String
    Dim ElementAttributeValue As String
    
    'recover from xml attribute's name and value
    ElementAttributeName = TagAttribute.getAttribute("name")
    ElementAttributeValue = GUIServices.PredefinedXMLToField(TagAttribute.getAttribute("value"))
    
    'include attributes in element
    sysadlElementInFile.AddField ElementAttributeName, ElementAttributeValue

End Sub

