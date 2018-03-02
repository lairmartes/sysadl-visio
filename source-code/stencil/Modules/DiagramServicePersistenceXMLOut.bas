Attribute VB_Name = "DiagramServicePersistenceXMLOut"

Option Explicit

    'current stream
    
    Dim CurrentFileExportingStream As Stream

    'constants to collapse or not
    
    Private Const TAG_COLLAPSED = True



Public Sub ProcessDiagramSaving(ByVal fileName As String, ByVal analysisResult As DiagramAnalysisResult)

    PublishAnalysisResult fileName, analysisResult

End Sub
    

'publish the result of analysis in XML format (using sysad extension)
'it publishes a file with the same visio document file name with a different extension
Private Sub PublishAnalysisResult(ByVal fileName As String, _
                                 ByVal analysisResult As DiagramAnalysisResult)
                                 
    ' declare tag <sysadldiagram>
    Dim TagSectionSysAdlDiagram As New XMLUtilTag
    Dim TagPropertyName As New XMLUtilTagValue
    Dim ExportXMLFileName As String
    Dim diagramQualifier As String
    Dim FileExportingStream As Stream
    

    'initialize diagram qualifier
    diagramQualifier = GUIServices.GetQualifierFromDocumentName(fileName)
    ExportXMLFileName = GUIServices.ChangeDocumentExtension(fileName, sysadlstringconstantsExtensionSysAdl)

    
    'configure file object
    
    Set CurrentFileExportingStream = New Stream
    CurrentFileExportingStream.Open
    CurrentFileExportingStream.Position = 0
    CurrentFileExportingStream.CharSet = "UTF-8"
    
    ' initializing sysadldiagram tag
    TagSectionSysAdlDiagram.Init sysAdlTagSectionDiagram
    
    TagPropertyName.Init sysAdlTagPropertyQualifier, diagramQualifier
    
    TagSectionSysAdlDiagram.AddProperty TagPropertyName

    ' publish tag <sysadldiagram>
    Publish TagSectionSysAdlDiagram.OpenTag
    
    'publish element data filled in diagram
    PublishDiagramElements analysisResult
    ' publish list of objectives
    PublishAnalysisResultObjectives analysisResult
    'publish list of structures (ports, intefaces, hosts, concerns, devices etc.)
    PublishAnalysisResultStructures analysisResult
    'publish list of relations (is-a, represents, composed-by and depends-of)
    PublishAnalysisResultRelations analysisResult
    'publish list of communications (nets, conversations and installations)
    PublishAnalysisResultCommunications analysisResult
    ' publish list of transitions (between states and layers)
    PublishAnalysisResultTransitions analysisResult
    
    ' publish tag </sysadldiagram>
    Publish TagSectionSysAdlDiagram.CloseTag
    
    PublishSkipLine
    
    'save file
    CurrentFileExportingStream.SaveToFile ExportXMLFileName, adSaveCreateOverWrite
    CurrentFileExportingStream.Close
    
    'Close #pFileExportingNumber
    

End Sub

'publish element data

Private Sub PublishDiagramElements(ByVal analysisResult As DiagramAnalysisResult)

    'declare tag <sysadl-element>
    Dim TagSectionSysAdlElement As New XMLUtilTag
    'declare innertag <attribute>
    Dim TagElementAttribute As New XMLUtilTag
    'create element data object to be published
    Dim CurrentElement As SysAdlElement
    Dim elementList As Collection
    
    'initiate tag elements
    
    TagSectionSysAdlElement.Init sysAdlTagSectionSysAdlElements
    
    Publish TagSectionSysAdlElement.OpenTag
    
    PublishElementDataGroup analysisResult.Channels 'sysAdlTypeSetChannel
    PublishElementDataGroup analysisResult.Decisions 'sysAdlTypeSetDecision
    PublishElementDataGroup analysisResult.Formats 'sysAdlTypeSetFormat
    PublishElementDataGroup analysisResult.Layers 'sysAdlTypeSetLayer
    PublishElementDataGroup analysisResult.Nodes ' sysAdlTypeSetNode
    PublishElementDataGroup analysisResult.Objectives 'sysAdlTypeSetObjective
    PublishElementDataGroup analysisResult.Assumptions 'sysAdlTypeSetAssumption
    PublishElementDataGroup analysisResult.Qualities 'sysAdlTypeSetQuality
    PublishElementDataGroup analysisResult.Receivers 'sysAdlTypeSetReceiver
    PublishElementDataGroup analysisResult.Roles 'sysAdlTypeSetRole
    PublishElementDataGroup analysisResult.Senders 'sysAdlTypeSetSender
    PublishElementDataGroup analysisResult.Stakeholders 'sysAdlTypeSetStakeholder
    PublishElementDataGroup analysisResult.Systems 'sysAdlTypeSetSystem
    PublishElementDataGroup analysisResult.Transitions 'sysAdlTypeSetTransition
    
    Publish TagSectionSysAdlElement.CloseTag
    PublishSkipLine

End Sub

' publish objectives
Private Sub PublishAnalysisResultObjectives(ByVal analysisResult As DiagramAnalysisResult)

    'Objective list to be published
    Dim ObjectiveList As Collection

    'declare tag <objectives>
    Dim TagSectionObjectives As New XMLUtilTag
    ' objectives are shapecontrollers
    Dim CurrentShapeController As shapeController
    
    'get objective list
    Set ObjectiveList = analysisResult.Objectives
    
    'init objectives tag
    TagSectionObjectives.Init sysAdlTagSectionObjectives
    
    'publish tag <objectives>
    Publish TagSectionObjectives.OpenTag
    
    'iterate in objective list
    For Each CurrentShapeController In ObjectiveList
        'publish tag <objective>
        PublishElementReference sysAdlTagSectionObjective, CurrentShapeController
        
    Next
    'publish tag </objectives>
   Publish TagSectionObjectives.CloseTag
   PublishSkipLine

End Sub

Private Sub PublishAnalysisResultStructures(ByVal analysisResult As DiagramAnalysisResult)
    ' declare tag <structures>
    Dim TagSectionStructures As New XMLUtilTag
    ' init tag <structures>
    TagSectionStructures.Init sysAdlTagSectionStructures
    ' publish tag <structures>
    Publish TagSectionStructures.OpenTag
    
    ' publish tag <protocols>
    PublishAnalysisResultStructure sysAdlTagSectionProtocols, _
                                   sysAdlTagSectionProtocol, _
                                   sysAdlTagSectionElementSystem, _
                                   sysAdlTagSectionElementReceiver, _
                                   analysisResult.Protocols, _
                                   sysAdlTypeSetSystem, _
                                   sysAdlTypeSetReceiver
                                   
    'publish tag <interfaces>
    PublishAnalysisResultStructure sysAdlTagSectionInterfaces, _
                                   sysAdlTagSectionInterface, _
                                   sysAdlTagSectionElementSystem, _
                                   sysAdlTagSectionElementFormat, _
                                   analysisResult.Interfaces, _
                                   sysAdlTypeSetSystem, _
                                   sysAdlTypeSetFormat
    'publish tag <hosts>
    PublishAnalysisResultStructure sysAdlTagSectionHosts, _
                                    sysAdlTagSectionHost, _
                                    sysAdlTagSectionElementNode, _
                                    sysAdlTagSectionElementSystem, _
                                    analysisResult.Hosts, _
                                    sysAdlTypeSetNode, _
                                    sysAdlTypeSetSystem
    'publish tag <ports>
    PublishAnalysisResultStructure sysAdlTagSectionPorts, _
                                    sysAdlTagSectionPort, _
                                    sysAdlTagSectionElementNode, _
                                    sysAdlTagSectionElementReceiver, _
                                    analysisResult.Ports, _
                                    sysAdlTypeSetNode, _
                                    sysAdlTypeSetReceiver
    'publish tag <deposits>
    PublishAnalysisResultStructure sysAdlTagSectionDeposits, _
                                    sysAdlTagSectionDeposit, _
                                    sysAdlTagSectionElementFormat, _
                                    sysAdlTagSectionElementSender, _
                                    analysisResult.Deposits, _
                                    sysAdlTypeSetFormat, _
                                    sysAdlTypeSetSender
    'publish tag <concerns>
    PublishAnalysisResultStructure sysAdlTagSectionConcerns, _
                                    sysAdlTagSectionConcern, _
                                    sysAdlTagSectionElementStakeholder, _
                                    sysAdlTagSectionElement, _
                                    analysisResult.Concerns, _
                                    sysAdlTypeSetStakeholder, _
                                    sysAdlStringConstantsEmpty
                                    
    'publish tag <requirements>
                                    
    PublishAnalysisResultStructure sysAdlTagSectionRequirements, _
                                    sysAdlTagSectionRequirement, _
                                    sysAdlTagSectionElementQuality, _
                                    sysAdlTagSectionElement, _
                                    analysisResult.Requirements, _
                                    sysAdlTypeSetQuality, _
                                    sysAdlStringConstantsEmpty
                                    
    'publish tag <responsibilities>
    
    PublishAnalysisResultStructure sysAdlTagSectionResponsibilities, _
                                   sysAdlTagSectionResponsibility, _
                                   sysAdlTagSectionElementRole, _
                                   sysAdlTagSectionElementStakeholder, _
                                   analysisResult.Responsibilities, _
                                   sysAdlTypeSetRole, _
                                   sysAdlTypeSetStakeholder
                                     

    'publish tag <devices>
    PublishAnalysisResultStructure sysAdlTagSectionDevices, _
                                    sysAdlTagSectionDevice, _
                                    sysAdlTagSectionElementNode, _
                                    sysAdlTagSectionElementSender, _
                                    analysisResult.Devices, _
                                    sysAdlTypeSetNode, _
                                    sysAdlTypeSetSender
                                    
    'publish tag <rationales>
                                    
    PublishAnalysisResultStructure sysAdlTagSectionRationales, _
                                    sysAdlTagSectionRationale, _
                                    sysAdlTagSectionElementDecision, _
                                    sysAdlTagSectionElementAssumption, _
                                    analysisResult.Rationales, _
                                    sysAdlTypeSetDecision, _
                                    sysAdlTypeSetAssumption
                                    
    'publish tag <definition>
    
    PublishAnalysisResultStructure sysAdlTagSectionDefinitions, _
                                   sysAdlTagSectionDefinition, _
                                   sysAdlTagSectionElementDecision, _
                                   sysAdlTagSectionElement, _
                                   analysisResult.Definitions, _
                                   sysAdlTypeSetDecision, _
                                   sysAdlStringConstantsEmpty
                                     
    
    ' publish end tag </structures>
   Publish TagSectionStructures.CloseTag
   PublishSkipLine

End Sub
' generic for structure publishing
Private Sub PublishAnalysisResultStructure(ByVal structureTagSection As String, _
                                          ByVal structureTagType As String, _
                                          ByVal elementTag1 As String, _
                                          ByVal elementTag2 As String, _
                                          ByVal structureList As Collection, _
                                          ByVal elementType1 As String, _
                                          ByVal elementType2 As String)

    ' declare tag for structure type (ports, devices etc.)
    Dim TagSectionStructureType As New XMLUtilTag
    ' declare tag for structure data (port, device, concern etc.)
    Dim TagSectionStructure As New XMLUtilTag
    ' declare property for structure id
    Dim TagPropertyStructureId As New XMLUtilTagValue

    ' declare elemens for iterations (to publish each point data from connection)
    Dim FirstShapeController As shapeController
    Dim SecondShapeController As shapeController
    ' declare current structure being published
    Dim CurrentDiagramStructure As DiagramStructureSysAdl
    
    'initialize specific structure type (ports, devices, concerns, hosts etc.)
    TagSectionStructureType.Init structureTagSection
    
    ' publish structure father structure (structure name in plural)
    Publish TagSectionStructureType.OpenTag
    
    For Each CurrentDiagramStructure In structureList
        ' init structure type
        TagSectionStructure.Init structureTagType
        ' intiate id of structure (connector id)
        TagPropertyStructureId.Init sysAdlTagPropertyId, CurrentDiagramStructure.Id
        ' add property defined high above
        TagSectionStructure.AddProperty TagPropertyStructureId
        ' publish tag start <port>, <device>, <concern>, <host> etc.
        Publish TagSectionStructure.OpenTag
        
        ' get first shape controller to get publishing data
        Set FirstShapeController = CurrentDiagramStructure.GetShapeControllerBySysAdlType(elementType1)
        
        ' check if the second element is a defined structure
        ' a defined structure is when an element is associated only with another element
        ' for example, in a interface a system element only can be connected with a format element
        ' other example, in a deposit a sender element only can be connected with a format element
        ' but, there is a situation where this relations is not 1 by 1
        ' for example, in a concern relation, a Stakeholder can be connected to a variety of other kinds of elements.
        ' because this, I must to see if the tag for second element has been provided
        ' to know if the element is generic, check if the tag is <element> or <element-[type]>
        If (elementTag2 <> sysAdlTagSectionElement) Then
            ' recover the element connected with a defined type
            Set SecondShapeController = CurrentDiagramStructure.GetShapeControllerBySysAdlType(elementType2)
        
        Else
            ' recover the other element connected (type can varying)
            Set SecondShapeController = CurrentDiagramStructure.GetShapeControllerDifferentSysAdlType(elementType1)
        
        End If
        
        'publish elements data
        PublishElementReference elementTag1, FirstShapeController
        PublishElementReference elementTag2, SecondShapeController
        
        ' publish tag end ex.: </port>, </device>, </concern>, </host> etc.
        Publish TagSectionStructure.CloseTag
        PublishSkipLine
        
    Next

    ' publish father end tag </ports>, </devices> etc
   Publish TagSectionStructureType.CloseTag
   PublishSkipLine

End Sub


'publish data of a shapecontroller
Private Sub PublishElementReference(ByVal tag As String, ByVal shapeController As shapeController)

    ' declare tag data
    Dim TagSectionElement As New XMLUtilTag
    Dim TagPropertyNamespace As New XMLUtilTagValue
    Dim TagPropertyId As New XMLUtilTagValue
    Dim TagPropertyShapeId As New XMLUtilTagValue

        ' initiate a tag for element
        TagSectionElement.Init tag, TAG_COLLAPSED
    
        ' add data for namespace, id and shapeid
        TagPropertyNamespace.Init sysAdlTagPropertyNamespace, shapeController.SysAdlElement.namespace
        TagPropertyId.Init sysAdlTagPropertyId, shapeController.SysAdlElement.Id
        TagPropertyShapeId.Init sysAdlTagPropertyShapeId, shapeController.GetShapeUniqueId
        ' add properties
        TagSectionElement.AddProperty TagPropertyNamespace
        TagSectionElement.AddProperty TagPropertyId
        TagSectionElement.AddProperty TagPropertyShapeId
        
        'public tag <element-[type]> (or only <element>
        Publish TagSectionElement.OpenTag

End Sub
Private Sub PublishElementDataGroup(ByVal ShapeControllerList As Collection)

    Dim CurrentShapeController As shapeController
    
    For Each CurrentShapeController In ShapeControllerList
    
        PublishElementData CurrentShapeController.SysAdlElement
        
    Next

End Sub

Private Sub PublishElementData(ByVal anElement As SysAdlElement)

    ' declare tag data for element head
    Dim TagSectionSysAdlElement As New XMLUtilTag
    Dim TagPropertySysAdlType As New XMLUtilTagValue
    Dim TagPropertyStereotype As New XMLUtilTagValue
    Dim TagPropertyNamespace As New XMLUtilTagValue
    Dim TagPropertyId As New XMLUtilTagValue
    Dim TagPropertyUrlInfo As New XMLUtilTagValue
    Dim StereotypePublished As String
    
    ' declare tag data for shape id
    Dim TagSectionShapeId As XMLUtilTag
    Dim TagPropertyShapeIdValue As XMLUtilTagValue
    
    
    'shapes from element
    Dim ElementShapeViewerList As Collection
    Dim CurrentElementShapeViewer As ShapeViewer
    
    'initiate <sysadl-element> tag
    TagSectionSysAdlElement.Init sysAdlTagSectionSysAdlElement
    
    'initializae stereotype value
    StereotypePublished = anElement.Stereotype
    'change to blank if equals <NO_STEROTYPE>
    If StereotypePublished = sysAdlNoStereotype Then _
        StereotypePublished = sysAdlStringConstantsEmpty
    
    ' add data for type, stereotype, namespace, id and UrlInfo
    TagPropertySysAdlType.Init sysAdlTagPropertyType, anElement.sysadlType
    TagPropertyStereotype.Init sysAdlTagPropertyStereotype, StereotypePublished
    TagPropertyNamespace.Init sysAdlTagPropertyNamespace, anElement.namespace
    TagPropertyId.Init sysAdlTagPropertyId, anElement.Id
    TagPropertyUrlInfo.Init sysAdlTagPropertyUrlInfo, GUIServices.PrepareFieldForXML(anElement.URLInfo)
        
    ' add properties
    TagSectionSysAdlElement.AddProperty TagPropertySysAdlType
    TagSectionSysAdlElement.AddProperty TagPropertyStereotype
    TagSectionSysAdlElement.AddProperty TagPropertyNamespace
    TagSectionSysAdlElement.AddProperty TagPropertyId
    TagSectionSysAdlElement.AddProperty TagPropertyUrlInfo
        
    'public tag <element-[type]> (or only <element>
    Publish TagSectionSysAdlElement.OpenTag
        
    Set ElementShapeViewerList = anElement.ShapeViewerList
    
    For Each CurrentElementShapeViewer In ElementShapeViewerList
    
        Set TagSectionShapeId = New XMLUtilTag
        Set TagPropertyShapeIdValue = New XMLUtilTagValue
        
        TagSectionShapeId.Init sysAdlTagSectionElementShape, TAG_COLLAPSED
        TagPropertyShapeIdValue.Init sysAdlTagPropertyId, CurrentElementShapeViewer.GetShapeId
        
        TagSectionShapeId.AddProperty TagPropertyShapeIdValue
        
        Publish TagSectionShapeId.OpenTag
    
    Next
    
    
    Publish TagSectionSysAdlElement.CloseTag
    PublishSkipLine


End Sub


' publish tag <relations>
Private Sub PublishAnalysisResultRelations(ByVal analysisResult As DiagramAnalysisResult)

    Dim TagSectionRelation As New XMLUtilTag
    
    TagSectionRelation.Init sysAdlTagSectionRelations
    
    Publish TagSectionRelation.OpenTag
    
    PublishAnalysisResultRelation analysisResult.IsARelations, sysAdlTypeSetIsA
    PublishAnalysisResultRelation analysisResult.RepresentsRelations, sysAdlTypeSetRepresents
    PublishAnalysisResultRelation analysisResult.ComposedByRelations, sysAdlTypeSetComposedBy
    PublishAnalysisResultInnerRelation analysisResult.ComposedByInnerRelations, sysAdlTypeSetComposedBy
    PublishAnalysisResultRelation analysisResult.DependsOnRelations, sysAdlTypeSetDependsOn
    
   Publish TagSectionRelation.CloseTag
   PublishSkipLine

End Sub


'publish tag <relation>
Private Sub PublishAnalysisResultRelation(ByVal relationList As Collection, ByVal relationType As String)

    
    Dim TagSectionRelation As New XMLUtilTag
    
    Dim TagPropertyRelationType As New XMLUtilTagValue
    Dim TagPropertyRelationId As New XMLUtilTagValue
    Dim CurrentDiagramRelation As DiagramRelationSysAdl
    
    TagPropertyRelationType.Init sysAdlTagPropertyType, relationType
    
    For Each CurrentDiagramRelation In relationList
    
        TagSectionRelation.Init sysAdlTagSectionRelation
    
        TagPropertyRelationId.Init sysAdlTagPropertyId, CurrentDiagramRelation.Id
        
        TagSectionRelation.AddProperty TagPropertyRelationType
        TagSectionRelation.AddProperty TagPropertyRelationId
        
        Publish TagSectionRelation.OpenTag
        
        PublishElementReference sysAdlTagSectionElementSource, CurrentDiagramRelation.RelationSource
        PublishElementReference sysAdlTagSectionElementDestiny, CurrentDiagramRelation.RelationDestiny
        
        Publish TagSectionRelation.CloseTag
        PublishSkipLine
    
    Next
    


End Sub

'publish tag <relation> for inner shapes
Private Sub PublishAnalysisResultInnerRelation(ByVal innerRelationList As Collection, ByVal relationType As String)

    
    Dim TagSectionRelation As New XMLUtilTag
    
    Dim TagPropertyRelationType As New XMLUtilTagValue
    Dim TagPropertyRelationId As New XMLUtilTagValue
    Dim CurrentDiagramRelation As DiagramInnerRelationSysADL
    
    TagPropertyRelationType.Init sysAdlTagPropertyType, relationType
    
    For Each CurrentDiagramRelation In innerRelationList
    
        TagSectionRelation.Init sysAdlTagSectionRelation
    
        TagPropertyRelationId.Init sysAdlTagPropertyId, CurrentDiagramRelation.Id
        
        TagSectionRelation.AddProperty TagPropertyRelationType
        TagSectionRelation.AddProperty TagPropertyRelationId
        
        Publish TagSectionRelation.OpenTag
        
        PublishElementReference sysAdlTagSectionElementSource, CurrentDiagramRelation.RelationSource
        PublishElementReference sysAdlTagSectionElementDestiny, CurrentDiagramRelation.RelationDestiny
        
        Publish TagSectionRelation.CloseTag
        PublishSkipLine
    
    Next
    


End Sub

'publish tag <communications>
Private Sub PublishAnalysisResultCommunications(ByVal analysisResult As DiagramAnalysisResult)


    Dim TagSectionCommunications As New XMLUtilTag
    Dim TagSectionNets As New XMLUtilTag
    Dim TagSectionInstallations As New XMLUtilTag
    Dim TagSectionConversations As New XMLUtilTag
    
    TagSectionCommunications.Init sysAdlTagSectionCommunications
    
    TagSectionNets.Init sysAdlTagSectionNets
    TagSectionInstallations.Init sysAdlTagSectionInstallations
    TagSectionConversations.Init sysAdlTagSectionConversations
    
    '<communications>
    Publish TagSectionCommunications.OpenTag
    
    '<nets>
    Publish TagSectionNets.OpenTag
    '<net>
    PublishAnalysisResultCommunication analysisResult.NetCommunications, sysAdlTagSectionNet
    '</nets>
    Publish TagSectionNets.CloseTag
    PublishSkipLine
    '<installations>
    Publish TagSectionInstallations.OpenTag
    '<installation>
    PublishAnalysisResultCommunication analysisResult.InstallationCommuncations, sysAdlTagSectionInstallation
    '</installations>
    Publish TagSectionInstallations.CloseTag
    PublishSkipLine
    '<conversations>
    Publish TagSectionConversations.OpenTag
    '<conversation>
    PublishAnalysisResultConversations analysisResult.ConversationCommunications
    '</conversations>
    Publish TagSectionConversations.CloseTag
    PublishSkipLine
    
    '</communications>
    Publish TagSectionCommunications.CloseTag
    PublishSkipLine


End Sub

'publish tag <net> or <installation>
Private Sub PublishAnalysisResultCommunication(ByVal communicationList As Collection, ByVal communicationTag As String)
    
    Dim TagSection As New XMLUtilTag
    Dim TagPropertyId As New XMLUtilTagValue
    Dim TagPropertyStructure As New XMLUtilTag
    Dim TagPropertyStructureId As New XMLUtilTagValue
    Dim TagPropertyNodeId As New XMLUtilTagValue
    Dim CurrentDiagramCommunication As DiagramCommunicationSysAdl
    
    For Each CurrentDiagramCommunication In communicationList
    
        TagSection.Init communicationTag
    
        TagPropertyId.Init sysAdlTagPropertyId, CurrentDiagramCommunication.Id
        
       If (communicationTag = sysAdlTagSectionNet) Then

            TagPropertyStructure.Init sysAdlTagSectionPort, True
            TagPropertyStructureId.Init sysAdlTagPropertyId, CurrentDiagramCommunication.GetStructureByType(sysAdlStructurePort).Id
            TagPropertyStructure.AddProperty TagPropertyStructureId
            
        ElseIf (communicationTag = sysAdlTagSectionInstallation) Then
        
            TagPropertyStructure.Init sysAdlTagSectionDevice, True
            TagPropertyStructureId.Init sysAdlTagPropertyId, CurrentDiagramCommunication.GetStructureByType(sysAdlStructureDevice).Id
            TagPropertyStructure.AddProperty TagPropertyStructureId
            
        End If
            
        TagPropertyNodeId.Init sysAdlTagPropertyNodeId, CurrentDiagramCommunication.GetStructureByType(sysAdlStructureSingleElement).Id
        
        TagSection.AddProperty TagPropertyId
        
        Publish TagSection.OpenTag
        
        Publish TagPropertyStructure.OpenTag
        
        PublishElementReference sysAdlTagSectionElementNode, CurrentDiagramCommunication.GetStructureByType(sysAdlStructureSingleElement).GetSingleElementController
        
        PublishElementReference sysAdlTagSectionElementChannel, CurrentDiagramCommunication.CommunicationShapeController
        
        Publish TagSection.CloseTag
        PublishSkipLine
    
    Next

End Sub

'publish tag <conversation>
Private Sub PublishAnalysisResultConversations(ByVal conversationList As Collection)

    Dim TagSection As New XMLUtilTag
    Dim TagPropertyId As New XMLUtilTagValue
    Dim TagPropertyType As New XMLUtilTagValue
    
    Dim TagSectionDeposit As New XMLUtilTag
    Dim TagSectionInterface As New XMLUtilTag
    Dim TagSectionProtocol As New XMLUtilTag
    
    Dim TagPropertyInterfaceId As New XMLUtilTagValue
    Dim TagPropertyProtocolId As New XMLUtilTagValue
    Dim TagPropertyDepositId As New XMLUtilTagValue

    
    Dim conversationType As String

    Dim CurrentConversation As DiagramCommunicationSysAdl
    
    For Each CurrentConversation In conversationList
    
        conversationType = CurrentConversation.GetConversationType
    
        TagPropertyId.Init sysAdlTagPropertyId, CurrentConversation.Id
        TagPropertyType.Init sysAdlTagPropertyType, conversationType
        
        TagSection.Init sysAdlTagSectionConversation
            
        TagSection.AddProperty TagPropertyId
        TagSection.AddProperty TagPropertyType
        
        Publish TagSection.OpenTag
        
        If conversationType = sysAdlStructureDeposit Then
        
            TagSectionDeposit.Init sysAdlTagSectionDeposit, True
        
            TagPropertyDepositId.Init sysAdlTagPropertyId, CurrentConversation.GetStructureByType(sysAdlStructureDeposit).Id
            
            TagSectionDeposit.AddProperty TagPropertyDepositId
            
            Publish TagSectionDeposit.OpenTag
            
            PublishElementReference sysAdlTagSectionElementSystem, CurrentConversation.GetStructureByType(sysAdlStructureSingleElement).GetSingleElementController
            
            PublishElementReference sysAdlTagSectionElementChannel, CurrentConversation.CommunicationShapeController
            
        ElseIf conversationType = sysAdlStructureProtocol Then
        
            TagSectionProtocol.Init sysAdlTagSectionProtocol, True
            TagSectionInterface.Init sysAdlTagSectionInterface, True
            
            TagPropertyProtocolId.Init sysAdlTagPropertyId, CurrentConversation.GetStructureByType(sysAdlStructureProtocol).Id
            
            TagPropertyInterfaceId.Init sysAdlTagPropertyId, CurrentConversation.GetStructureByType(sysAdlStructureInterface).Id
            
            TagSectionProtocol.AddProperty TagPropertyProtocolId
            
            TagSectionInterface.AddProperty TagPropertyInterfaceId
            
            Publish TagSectionProtocol.OpenTag
            Publish TagSectionInterface.OpenTag
            
            PublishElementReference sysAdlTagSectionElementChannel, CurrentConversation.CommunicationShapeController
            
        End If
        
        Publish TagSection.CloseTag
        PublishSkipLine
    
    Next
    

End Sub

'publish tag <transitions>
Private Sub PublishAnalysisResultTransitions(ByVal analysisResult As DiagramAnalysisResult)

    Dim TagSection As New XMLUtilTag
    
    TagSection.Init sysAdlTagSectionTransitions
    
    Publish TagSection.OpenTag
    
    PublishAnalysisResultTransition analysisResult.TransitionRelations
    
   Publish TagSection.CloseTag
   PublishSkipLine

End Sub

'publish tag <transition>
Private Sub PublishAnalysisResultTransition(ByVal TransitionList As Collection)

    
    Dim TagSection As New XMLUtilTag
    
    Dim TagPropertyType As New XMLUtilTagValue
    Dim TagPropertyId As New XMLUtilTagValue
    Dim CurrentDiagramTransition As DiagramRelationSysAdl
    
    For Each CurrentDiagramTransition In TransitionList
    
        TagSection.Init sysAdlTagSectionTransition
    
        TagPropertyId.Init sysAdlTagPropertyId, CurrentDiagramTransition.Id
        
        TagPropertyType.Init sysAdlTagPropertyType, CurrentDiagramTransition.RelationSource.ShapeSysAdlType
        
        TagSection.AddProperty TagPropertyId
        TagSection.AddProperty TagPropertyType
        
        Publish TagSection.OpenTag
        
        PublishElementReference sysAdlTagSectionElementSource, CurrentDiagramTransition.RelationSource
        PublishElementReference sysAdlTagSectionElementDestiny, CurrentDiagramTransition.RelationDestiny
        PublishElementReference sysAdlTagSectionElementTransition, CurrentDiagramTransition.RelationConnector
        
        Publish TagSection.CloseTag
        
        PublishSkipLine
    
    Next
    


End Sub
'print line in file
Private Sub Publish(ByVal aLine As String)

    CurrentFileExportingStream.WriteText aLine

End Sub

Private Sub PublishSkipLine()

    CurrentFileExportingStream.SkipLine
    
End Sub



