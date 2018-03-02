Attribute VB_Name = "DiagramServicePersistenceXMLIn"
Option Explicit

    Dim ImportedAnalysisList As New Collection


Public Sub ProcessOpenDiagram(ByVal fileName As String)

    'declare path to reach sadl document
    Dim SADLPath As String
    'declare XML holder
    Dim SADLDocument As New DOMDocument
    'diagram qualifier
    Dim diagramQualifier As String

    GUIServices.LogMessage ">>>>>>>>>>>>>>>>>>>> IMPORTING DIAGRAM <<<<<<<<<<<<<<<<<<<<<<<<<<<<"

    GUIServices.LogMessage "Importing diagram: " + fileName

    'convert namespace and diagram name to get SADL file path
    SADLPath = GUIServices.ChangeDocumentExtension(fileName, sysadlstringconstantsExtensionSysAdl)
    diagramQualifier = GUIServices.GetQualifierFromDocumentName(fileName)
    
    GUIServices.LogMessage "Importing SADL File: " + SADLPath
    
    'configure DOM Object
    SADLDocument.async = False
    
    'Load file
    Call SADLDocument.Load(SADLPath)
    
    'Treat error (if any)
    If SADLDocument.parseError.reason <> sysAdlStringConstantsEmpty Then
    
        Dim messageError As String
       
        GUIServices.LogMessage "Error line  : " + Conversion.CStr(SADLDocument.parseError.Line)
        GUIServices.LogMessage "Error column: " + Conversion.CStr(SADLDocument.parseError.linepos)
        GUIServices.LogMessage "Text error  : " + SADLDocument.parseError.srcText
       
        messageError = "An error has occurred while loading file: " + SADLDocument.parseError.reason
        
        MsgBox messageError, vbCritical, "Diagram can't be opened"

        
    Else

        'Import data of diagram
        ProcessImport SADLDocument, diagramQualifier
    
    End If


End Sub


Public Function GetDiagramElementByShapeId(ByVal ElementShapeId As String, ByVal elementType As String) As SysAdlElement

    Dim CurrentAnalysis As DiagramAnalysisResult
    Dim DiagramElementList As Collection
    Dim CurrentDiagramElement As DiagramElement
    Dim Result As SysAdlElement
    
    For Each CurrentAnalysis In ImportedAnalysisList
    
        Set DiagramElementList = CurrentAnalysis.DiagramElements
        
        For Each CurrentDiagramElement In DiagramElementList
        
            If CurrentDiagramElement.IsSameShapeId(ElementShapeId) Then
            
                If CurrentDiagramElement.elementType = elementType Then
                
                    Set Result = New SysAdlElement
                    
                    Result.Init CurrentDiagramElement.elementType, CurrentDiagramElement.ElementStereotype
                    Result.InitBaseData CurrentDiagramElement.elementNamespace, _
                                        CurrentDiagramElement.elementId, _
                                        CurrentDiagramElement.ElementURLInfo
                                        
                    Set Result = ElementServicePersistence.OpenPersistedElement(Result)
                    
                    Exit For
                
                End If
                
            End If
            
        Next
        
        If Not (Result Is Nothing) Then Exit For
    
    Next
    
    Set GetDiagramElementByShapeId = Result

End Function

Private Sub ProcessImport(ByVal SADLDocument As DOMDocument, _
                          ByVal fileName As String)
                          

    ' declare tags
    ' tags have been divided in hierarch
    '  Tag Diagram - All diagram information started in <diagram> tag
    '  Tag Group - Groups belonged to diagram. Types are:
    '                                                     <elements>
    '                                                     <objectives>
    '                                                     <structures>
    '                                                     <communications>
    '                                                     <transitions>
    Dim TagDiagram As IXMLDOMElement
    Dim TagGroup As IXMLDOMElement
    Dim I As Integer
    
    ' declare nodes
    ' nodes are childs of a Tag
    ' for example, children of tag <diagram> are <elements>, <objectives>, <structures> etc
    Dim NodeDiagram As IXMLDOMNodeList
    Dim NodeGroups As IXMLDOMNodeList
    
    ' get children from whole diagram, in this case is <diagram>
    Set NodeDiagram = SADLDocument.ChildNodes
    
    ' get <diagram> tag
    Set TagDiagram = NodeDiagram.NextNode
    
    ' get children from <diagram> tag
    Set NodeGroups = TagDiagram.ChildNodes
    
    'iterate withing <diagram> tag children
    For I = 0 To NodeGroups.Length - 1
    
        ' get current tag group. Are expected <elements>, <objectives>, <structures>, <communications> and <transitions>
        Set TagGroup = NodeGroups.Item(I)
    
        ' handle <element> tag
        If sysAdlTagSectionSysAdlElements = TagGroup.nodeName Then
            
            ' handle <elements>
            ImportElements TagGroup, fileName
            
        End If
    
    Next
    
End Sub



Private Sub ImportElements(ByVal TagElementGroup As IXMLDOMElement, ByVal fileName As String)
    ' Declare children list for <elements> tag. Expected is <element>
    Dim NodeElementGroups As IXMLDOMNodeList
    ' declare tag to hold <element> data
    Dim TagElementData As IXMLDOMElement
    
    Dim I As Integer
    
    ' Get elements from <elements>
    Set NodeElementGroups = TagElementGroup.ChildNodes
    
    For I = 0 To NodeElementGroups.Length - 1
    
        ' get xml data for current element
        Set TagElementData = NodeElementGroups.Item(I)
        
        'import basic data of element
        ImportElement TagElementData, fileName
        
    Next
    
    

End Sub



Private Sub ImportElement(ByVal TagElement As IXMLDOMElement, ByVal fileName As String)

    'declare variables to hold xml data
    Dim elementType As String
    Dim ElementStereotype As String
    Dim elementNamespace As String
    Dim elementId As String
    Dim ElementURLInfo As String
    Dim ElementShapeId As String
    
    'declare children of TagElement
    Dim NodeElementItems As IXMLDOMNodeList
    'declare tag to hold item data
    Dim TagItem As IXMLDOMElement
    
    Dim I As Integer
    
    'get element basic data
    elementType = TagElement.getAttribute("type")
    ElementStereotype = TagElement.getAttribute("stereotype")
    elementNamespace = TagElement.getAttribute("namespace")
    elementId = TagElement.getAttribute("id")
   ElementURLInfo = GUIServices.PredefinedXMLToField(TagElement.getAttribute("url-info"))

    
    'convert to <NO_STEREOTYPE> if not provided
    If ElementStereotype = sysAdlStringConstantsEmpty Then
    
        ElementStereotype = sysAdlNoStereotype
        
    End If
    
    GUIServices.LogMessage " --------------------- ELEMENT -------------------------------"
    
    
    'get element items.  <attribute> and <shape> are expected
    Set NodeElementItems = TagElement.ChildNodes
    
    GUIServices.LogMessage "    ............... atributes & shapes ................."
    
    'iterate withing element items
    For I = 0 To NodeElementItems.Length - 1
    
        Set TagItem = NodeElementItems.Item(I)
        
        'import shape data
        If sysAdlTagSectionElementShape = TagItem.nodeName Then
        
            ElementShapeId = TagItem.getAttribute("id")
        
            ImportDiagramElement elementType, ElementStereotype, elementNamespace, elementId, ElementURLInfo, ElementShapeId, fileName
            
        End If
    
    Next


End Sub


Private Sub ImportDiagramElement(ByVal elementType As String, _
                                   ByVal ElementStereotype As String, _
                                   ByVal elementNamespace As String, _
                                   ByVal elementId As String, _
                                   ByVal ElementURLInfo As String, _
                                   ByVal ElementShapeId As String, _
                                   ByVal fileName As String)
                                   
    Dim importedDiagramElement As New DiagramElement
    Dim analysisResultFile As DiagramAnalysisResult

    GUIServices.LogMessage "Element Type......: " + elementType
    GUIServices.LogMessage "Element Stereotype: " + ElementStereotype
    GUIServices.LogMessage "Element Namespace.: " + elementNamespace
    GUIServices.LogMessage "Element Id........: " + elementId
    GUIServices.LogMessage "Element URL Info..: " + ElementURLInfo
    GUIServices.LogMessage "Element Shape Id..: " + ElementShapeId
    GUIServices.LogMessage "File Name.........: " + fileName

    importedDiagramElement.Init elementNamespace, elementId, elementType, ElementStereotype, ElementURLInfo, ElementShapeId

    Set analysisResultFile = GetDiagramAnalysisByFileName(fileName)
    
    
    analysisResultFile.AddDiagramElement importedDiagramElement
                                         
End Sub

Private Function GetDiagramAnalysisByFileName(ByVal fileName As String) As DiagramAnalysisResult

    Dim CurrentAnalysis As DiagramAnalysisResult
    
    Dim Result As DiagramAnalysisResult
    
    For Each CurrentAnalysis In ImportedAnalysisList
    
        If CurrentAnalysis.isDiagramNameMatch(fileName) Then
        
            Set Result = CurrentAnalysis
            Exit For
            
        End If
    
    Next
    
    If Result Is Nothing Then
    
        Set Result = New DiagramAnalysisResult
        
        Result.Init fileName
        
        ImportedAnalysisList.Add Result
        
    End If

    Set GetDiagramAnalysisByFileName = Result

End Function
