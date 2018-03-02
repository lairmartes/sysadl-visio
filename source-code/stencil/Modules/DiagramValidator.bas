Attribute VB_Name = "DiagramValidator"
Option Explicit

    'list of elements bound in structures
    Private ProtocolShapeList As UtilSysAdlList
    Private InterfaceShapeList As UtilSysAdlList
    Private HostShapeList As UtilSysAdlList
    Private PortShapeList As UtilSysAdlList
    Private DepositShapeList As UtilSysAdlList
    Private ConcernShapeList As UtilSysAdlList
    Private RequirementShapeList As UtilSysAdlList
    Private ResponsibilityShapeList As UtilSysAdlList

    'diagram analysis result
    
    Private analysisResult As DiagramAnalysisResult

    
    'list of analyzed channels
    Private AnalyzedChannelsList As Collection
    
    'number of elements that a connector must connect
    Private Const MINIMUN_CONNECTION_COUNT = 2
    
    'Page in analysis
    Private VisioDocInAnalysis As IVDocument
    

    

Public Function ValidateDiagram(ByVal Doc As IVDocument) As DiagramAnalysisResult
    
    Dim DiagramIsValid As Boolean
    
    Dim ConnectorList As Collection
    
    Dim ObjectiveList As Collection
    
    Dim DocElementList As Collection
    
    Dim CurrentObjective As shapeController
    

    Set analysisResult = New DiagramAnalysisResult
    
    analysisResult.Init Doc.FullName
    
    Set VisioDocInAnalysis = Doc
    
    'initializing list of elements bound in structures
    Set ProtocolShapeList = New UtilSysAdlList
    Set InterfaceShapeList = New UtilSysAdlList
    Set HostShapeList = New UtilSysAdlList
    Set PortShapeList = New UtilSysAdlList
    Set DepositShapeList = New UtilSysAdlList
    Set ConcernShapeList = New UtilSysAdlList
    Set ResponsibilityShapeList = New UtilSysAdlList
    Set RequirementShapeList = New UtilSysAdlList
    
    'initializing list of analyzed channels
    Set AnalyzedChannelsList = New Collection
    
    'get list of connectors that has built structures
    Set ConnectorList = GetConnectorControllers()
            
    'first I valid structrures looking for inconsitencies
    ValidateStructures ConnectorList, sysAdlTypeSetConnector
    
    DiagramIsValid = Not analysisResult.HasError

    If DiagramIsValid Then
            
        'check if relations depends of are ok
        ValidateRelations GetRelationControllers(sysAdlTypeSetDependsOn)
        ' check if representes are ok
        ValidateRelations GetRelationControllers(sysAdlTypeSetRepresents)
        ' check if is a are ok
        ValidateRelations GetRelationControllers(sysAdlTypeSetIsA)
        ' check if transitions are ok
        ValidateRelations GetRelationControllers(sysAdlTypeSetTransition)
        ' check if composed by are ok
        ValidateRelations GetRelationControllers(sysAdlTypeSetComposedBy)
        ' check inner relations
        ValidateInnerRelations GetRelationControllers(sysAdlTypeSetLayer)
        ValidateInnerRelations GetRelationControllers(sysAdlTypeSetSystem)
        ValidateInnerRelations GetRelationControllers(sysAdlTypeSetNode)
        
        
        DiagramIsValid = Not analysisResult.HasError
        
        If DiagramIsValid Then
             
            ' validate communications
            ValidateCommunications

        End If
        
    End If
    
    ' include objectives
    Set ObjectiveList = GetShapeControllersFromDoc(sysAdlTypeSetObjective)
        
    For Each CurrentObjective In ObjectiveList
        
        analysisResult.AddObjective CurrentObjective

    Next

    'include other elements in analysis result
    
    IncludeElementsInAnalysis sysAdlTypeSetChannel
    IncludeElementsInAnalysis sysAdlTypeSetDecision
    IncludeElementsInAnalysis sysAdlTypeSetFormat
    IncludeElementsInAnalysis sysAdlTypeSetLayer
    IncludeElementsInAnalysis sysAdlTypeSetNode
    IncludeElementsInAnalysis sysAdlTypeSetAssumption
    IncludeElementsInAnalysis sysAdlTypeSetQuality
    IncludeElementsInAnalysis sysAdlTypeSetReceiver
    IncludeElementsInAnalysis sysAdlTypeSetSender
    IncludeElementsInAnalysis sysAdlTypeSetStakeholder
    IncludeElementsInAnalysis sysAdlTypeSetSystem
    IncludeElementsInAnalysis sysAdlTypeSetTransition
    IncludeElementsInAnalysis sysAdlTypeSetRole
    
    ' return result of analisys (True if diagram ok or False if not ok
    Set ValidateDiagram = analysisResult
    

End Function

Private Sub IncludeElementsInAnalysis(ByVal sysadlType As String)

    Dim CurrentElement As shapeController
    
    Dim elementList As Collection

    Set elementList = GetShapeControllersFromDoc(sysadlType)
        
    For Each CurrentElement In elementList
        
        analysisResult.AddElement CurrentElement

    Next


End Sub


' checking structures built by user
Private Sub ValidateStructures(ByVal SysAdlConnectorsController As Collection, _
                                   ByVal SysAdlConnectorType As String)

    ' element viewer that has been found in sysadl element (it keeps a reference for the shape and viewer has facilities to access it)
    Dim ElementShapeViewer As ShapeViewer
    ' Shape that is in evaluation
    Dim CurrentShape As shape
    ' ShapeController that has facilities like to discover the type of shape
    Dim CurrentShapeController As shapeController
    ' List of elements that are connected in Connector
    Dim ShapeConnectorsList As Visio.Connects
       
    ' Diagram structure built here (can be valid or not - if valid, it is added on structure list)
    Dim StructureToCheck As DiagramStructureSysAdl
    
    ' the first shape connected in Connector
    Dim ConnectedShapeController1st As shapeController
    ' the second shape connected in Connector
    Dim ConnectedShapeController2nd As shapeController
    ' the shape controller of Connector
    Dim ConnectedShapeControllerConnector As shapeController
    
    ' the shapeviewer of 1st shape controller connected
    Dim ConnectedShapeViewer1st As ShapeViewer
    ' the shapeviewer of 2nd shape controller connected
    Dim ConnectedShapeViewer2nd As ShapeViewer
    ' the shapeviewer of shape controller of connector
    Dim ConnectedShapeViewerConnector As ShapeViewer
    
    ' the sysadltype of 1st element connected
    Dim ConnectedElementType1st As String
    ' the sysadltype of 2nd element connected
    Dim ConnectedElementType2nd As String
    
    ' object Visio.Connect of first element
    Dim Connector1st As Connect
    ' object Visio.Connect of second element
    Dim Connector2nd As Connect
    
    ' object Visio.Shape of first element
    Dim ShapeConnected1st As shape
    ' object Visio.Shape of first element
    Dim ShapeConnected2nd As shape

    ' error message to be shown in shape comments
    Dim ErrorMessage As String
    

        ' iterate list of connectors
        For Each CurrentShapeController In SysAdlConnectorsController
        
                ' get shape controller's viewer
                Set ElementShapeViewer = CurrentShapeController.ShapeViewer
                
                ' current connector of element
                Set CurrentShape = ElementShapeViewer.shape
            
                ' list of elements connected in this connector
                Set ShapeConnectorsList = CurrentShape.Connects
                
                ' the connector must be connected with, at least, two shapes
                If ShapeConnectorsList.Count < MINIMUN_CONNECTION_COUNT Then
                    
                    'of course, if it is not connected with at least two shapes...
                    PublishError CurrentShapeController, "" + SysAdlConnectorType + " must connect at least two elements"
                
                Else
                
                    ' give name to our pieces...
                    Set Connector1st = ShapeConnectorsList.Item(1)
                    Set Connector2nd = ShapeConnectorsList.Item(2)
                    
                    Set ShapeConnected1st = Connector1st.ToSheet
                    Set ShapeConnected2nd = Connector2nd.ToSheet

                    Set ConnectedShapeController1st = FactoryShapeController.GetShapeControllerByShape(ShapeConnected1st)
                    Set ConnectedShapeController2nd = FactoryShapeController.GetShapeControllerByShape(ShapeConnected2nd)
                    
                    ConnectedElementType1st = ConnectedShapeController1st.ShapeSysAdlType
                    ConnectedElementType2nd = ConnectedShapeController2nd.ShapeSysAdlType
                    
                    Set ConnectedShapeViewer1st = ConnectedShapeController1st.ShapeViewer
                    Set ConnectedShapeViewer2nd = ConnectedShapeController2nd.ShapeViewer

                    ' building structure
                    Set StructureToCheck = New DiagramStructureSysAdl
                    
                    ' initializing structure with elements discovered
                    StructureToCheck.Init ConnectedShapeController1st, _
                                          ConnectedShapeController2nd, _
                                          CurrentShapeController
                
                    ' get structure type and check if it is valid (checking if structure type is not an empty value)
                    If StructureToCheck.StructureType = sysAdlStringConstantsEmpty Then
                    
                        'if structure is not valid, then we have a problem...
                        ErrorMessage = "It is not possible to use " + SysAdlConnectorType + " between " + ConnectedElementType1st + " and " + ConnectedElementType2nd
                        
                        ' publish error in shapes
                        PublishError CurrentShapeController, ErrorMessage
                        PublishError ConnectedShapeController1st, ErrorMessage
                        PublishError ConnectedShapeController2nd, ErrorMessage
                        
                    Else
                        ' add elements discovered in the list of elements that belongs to structures
                        AddElementsInStructureInList StructureToCheck
                        
                        ' add structure in the list of structures
                        AddStructureInList StructureToCheck
                        
                    End If
                    
                End If
            
        Next
    

End Sub


' add structure in the correct list depending of its type
Private Sub AddStructureInList(ByVal aStructure As DiagramStructureSysAdl)

    analysisResult.AddStructure aStructure

End Sub
' checking structures built by user
Private Sub ValidateRelations(ByVal SysAdlConnectorsController As Collection)

    ' element viewer that has been found in sysadl element (it keeps a reference for the shape and viewer has facilities to access it)
    Dim ElementShapeViewer As ShapeViewer
    ' Shape that is in evaluation
    Dim CurrentShape As shape
    ' ShapeController that has facilities like to discover the type of shape
    Dim CurrentShapeController As shapeController
    ' List of elements that are connected in Connector
    Dim ShapeConnectorsList As Visio.Connects
       
    ' Diagram structure built here (can be valid or not - if valid, it is added on structure list)
    Dim relationFound As DiagramRelationSysAdl

    ' error message to be shown in shape comments
    Dim ErrorMessage As String

        ' iterate list of connectors
        For Each CurrentShapeController In SysAdlConnectorsController
        
                ' get shape controller's viewer
                Set ElementShapeViewer = CurrentShapeController.ShapeViewer
                
                ' current connector of element
                Set CurrentShape = ElementShapeViewer.shape
            
                ' list of elements connected in this connector
                Set ShapeConnectorsList = CurrentShape.Connects
                
                ' the connector must be connected with, at least, two shapes
                If ShapeConnectorsList.Count < MINIMUN_CONNECTION_COUNT Then
                    
                    'of course, if it is not connected with at least two shapes...
                    PublishError CurrentShapeController, "" + ElementShapeViewer.GetSysAdlTypeOfShape + " must connect at least two elements"

                
                Else
                
                    Set relationFound = New DiagramRelationSysAdl
                    
                    relationFound.Init CurrentShapeController
                    
                    If relationFound.Relation = sysAdlStringConstantsEmpty Then
                    
                        ErrorMessage = "It is not possible to define a relation " + ElementShapeViewer.GetSysAdlTypeOfShape + _
                                        " from " + relationFound.RelationSource.ShapeSysAdlType + _
                                        " to " + relationFound.RelationDestiny.ShapeSysAdlType
                                        
                        PublishError CurrentShapeController, ErrorMessage
                        PublishError relationFound.RelationSource, ErrorMessage
                        PublishError relationFound.RelationDestiny, ErrorMessage

                        
                    Else
                        
                        ' add relation found in analysis result
                        analysisResult.AddRelation relationFound
                        
                    End If
                    
                End If
            
        Next

End Sub

' check the elements that are inside layer to build Composed-by relations
Private Sub ValidateInnerRelations(ByVal SysAdlOuterControllerList As Collection)
    
    Dim currentOuterController As shapeController
    Dim elementCollection As Collection
    Dim ElementController As shapeController
    Dim relationFound As DiagramInnerRelationSysADL

    
    For Each currentOuterController In SysAdlOuterControllerList

        
        ' add inner relations in System, Layers, Nodes and Stakeholders controllers
        
        AddInnerRelations currentOuterController, GetShapeControllersFromDoc(sysAdlTypeSetSystem) ' systems
        AddInnerRelations currentOuterController, GetShapeControllersFromDoc(sysAdlTypeSetLayer) ' layers
        AddInnerRelations currentOuterController, GetShapeControllersFromDoc(sysAdlTypeSetNode) ' nodes
        AddInnerRelations currentOuterController, GetShapeControllersFromDoc(sysAdlTypeSetStakeholder) ' stakeholders
      
    Next


End Sub

Private Sub AddInnerRelations(ByVal OuterController As shapeController, ByVal InnerControllerList As Collection)

        Dim ElementController As shapeController
        Dim relationFound As DiagramInnerRelationSysADL

        For Each ElementController In InnerControllerList
        
            Set relationFound = New DiagramInnerRelationSysADL
        
            relationFound.Init OuterController, ElementController
            
            If (relationFound.Relation <> sysAdlStringConstantsEmpty) Then
            
                analysisResult.AddInnerRelation relationFound
                
            End If
        
        Next

End Sub

' validate communications built by user
Private Sub ValidateCommunications()
    
    ' check communications for protocols
    ValidateCommunicationsStructures analysisResult.Protocols, sysAdlStructureProtocol
    
    ' what is applicable for protocols, are for interfaces...
    ValidateCommunicationsStructures analysisResult.Interfaces, sysAdlStructureInterface
    
    ' ... and ports...
    ValidateCommunicationsStructures analysisResult.Ports, sysAdlStructurePort
    
    ' ... and deposits...
    ValidateCommunicationsStructures analysisResult.Deposits, sysAdlStructureDeposit
    
    ' ... and (recently I confess) devices!
    ValidateCommunicationsStructures analysisResult.Devices, sysAdlStructureDevice

    ' validate communications that have no structures
    ValidateCommunicationsWithoutStructures

End Sub

Private Sub ValidateCommunicationsStructures(ByVal structureList As Collection, ByVal StructureType As String)
    
    ' current structure being checked
    Dim CurrentStructureSysAdl As DiagramStructureSysAdl
    
    ' channel that is connected in structure
    Dim ShapeConnectableInChannel As shapeController
    
    ' channels that are connected to structure
    Dim ConnectedChannels As Collection
    
    ' shape viewer of connected channel
    Dim ConnectorChannel As shapeController
    
    ' list of shapes connected in channel
    Dim ShapesConnectedInChannel As Visio.Connects
    
    ' shapes connected in channel
    Dim ShapeConnected1st As shape
    Dim ShapeConnected2nd As shape
    
    ' viewers of shapes connected
    Dim ShapeControllerConnected1st As shapeController
    Dim ShapeControllerConnected2nd As shapeController
    
    ' viewer of channel
    Dim ShapeControllerConnected As shapeController
    
    ' communication discovered here (it can be valid or not)
    Dim Communication As DiagramCommunicationSysAdl
    
    ' find the structure of an element connected to the channel
    Dim StructureFound As DiagramStructureSysAdl
    
    ' list of Visio elements connected to the channel
    Dim Connector1st As Connect
    Dim Connector2nd As Connect

    
    ' iterate in structure list received in parameter
    For Each CurrentStructureSysAdl In structureList
    
        'initializing the shape that is connected
        Set ShapeConnectableInChannel = Nothing
    
        ' if the type of structure is Protocol then...
        If (StructureType = sysAdlStructureProtocol) Then
         
            ' ...I must find the receiver (that is the element the user is supposed to connect in the channel)
            Set ShapeConnectableInChannel = CurrentStructureSysAdl.GetShapeControllerBySysAdlType(sysAdlTypeSetReceiver)
        
        ' if the type of structure is Port then...
        ElseIf (StructureType = sysAdlStructurePort) Then
        
            ' ...I must find the receiver (that is the element the user is supposed to connect in the channel)
            Set ShapeConnectableInChannel = CurrentStructureSysAdl.GetShapeControllerBySysAdlType(sysAdlTypeSetReceiver)
            
        ' if the type of structure is Device then...
        ElseIf (StructureType = sysAdlStructureDevice) Then
        
            ' ...I must find the sender (that is the element the user is supposed to connect in the channel)
            Set ShapeConnectableInChannel = CurrentStructureSysAdl.GetShapeControllerBySysAdlType(sysAdlTypeSetSender)
            
        ' if  the type of structure is Interface then...
        ElseIf (StructureType = sysAdlStructureInterface) Then
        
            ' ...I must find the format (that is the... ow! you know this history!)
            Set ShapeConnectableInChannel = CurrentStructureSysAdl.GetShapeControllerBySysAdlType(sysAdlTypeSetFormat)
            
        ' if the type of structure is Deposit then...
        ElseIf (StructureType = sysAdlStructureDeposit) Then
        
            ' now the element is the sender!
            Set ShapeConnectableInChannel = CurrentStructureSysAdl.GetShapeControllerBySysAdlType(sysAdlTypeSetSender)
        
        End If
        
        ' If I have found an element to be connected, I can check the rest.
        ' obs.: This element must exist because it has been validated in first check
        If Not (ShapeConnectableInChannel Is Nothing) Then
        
            ' Get list of channels connected to the element
            Set ConnectedChannels = GetChannelListFromElement(ShapeConnectableInChannel)
    
            ' For each channel connected in element
            For Each ConnectorChannel In ConnectedChannels
            
                ' add channel to the list of analyzed channels independently of analysis result
                AddShapeViewerToAnalyzedChannels ConnectorChannel
            
                ' Get shapes connected to the channel
                '(Did you get? First, I get the channels connected to the element and, then, I get the shapes connected to this channel)
                Set ShapesConnectedInChannel = ConnectorChannel.shape.Connects
                
                ' If this channel is not connected to at least two elements I have a problem...
                If (ShapesConnectedInChannel.Count < MINIMUN_CONNECTION_COUNT) Then
                
                    ' show to user that communication is not ok...
                    PublishError ConnectorChannel, "A Channel must communicate at least two elements"
                
                Else
                
                    ' initalize elements discovered
                    Set Connector1st = ShapesConnectedInChannel.Item(1)
                    Set Connector2nd = ShapesConnectedInChannel.Item(2)
                
                    Set ShapeConnected1st = Connector1st.ToSheet
                    Set ShapeConnected2nd = Connector2nd.ToSheet
                    
                    Set ShapeControllerConnected1st = FactoryShapeController.GetShapeControllerByShape(ShapeConnected1st)
                    Set ShapeControllerConnected2nd = FactoryShapeController.GetShapeControllerByShape(ShapeConnected2nd)
                    
                    Set StructureFound = Nothing
                    
                    ' Discover wich element is the other end of this communication
                    Set ShapeControllerConnected = GetShapeToEvaluate(ShapeControllerConnected1st, ShapeControllerConnected2nd, ShapeConnectableInChannel)
                    
                    ' if structure is an interface or a protocol...
                    If (StructureType = sysAdlStructureInterface Or StructureType = sysAdlStructureProtocol) Then
                        
                        'if structure is a interface...
                        If (StructureType = sysAdlStructureInterface) Then
                                
                            ' ... I must check if the other end is a protocol
                            Set StructureFound = GetProtocol(ShapeControllerConnected)
                            
                        ' if structure is a protocol...
                        ElseIf (StructureType = sysAdlStructureProtocol) Then
                        
                            '... I must check if the other end is an interface
                            Set StructureFound = GetInterface(ShapeControllerConnected)
                            
                        End If
                        
                        ' If the structure is not an Interface neither a Protocol (depending of the case)...
                        If StructureFound Is Nothing Then
                        
                            '... I build a structure that has a single element (the other end's element)
                            Set StructureFound = New DiagramStructureSysAdl
                            StructureFound.InitSingleElement ShapeControllerConnected
                            
                        End If
                        
                    ' if the structure type is another...
                    Else
                    
                        ' just create the single element structure that is in the other end
                        Set StructureFound = New DiagramStructureSysAdl
                        StructureFound.InitSingleElement ShapeControllerConnected
                        
                    End If
                        
                    ' check if connection is valid between discovered elements
                    ProcessCommunication CurrentStructureSysAdl, _
                                          StructureFound, _
                                          ConnectorChannel
                
                End If
            
            Next
        
        End If
    
    Next

End Sub

' check if there is channels that are not connected to structures

Private Function ValidateCommunicationsWithoutStructures()

    Dim Result As Boolean
    
    Dim CurrentChannelViewer As ShapeViewer
    Dim CurrentChannelShape As shape
    Dim CurrentChannelConnect As Visio.Connect
    Dim ChannelConnectorsList As Visio.Connects
    
    Dim CurrentChannelController As shapeController
    Dim ChannelControllerList As Collection
    
    Dim ConnectedShape As shape
    Dim ConnectedShapeController As shapeController
    
    Dim ErrorMessage As String
    
    Result = True
    
    ErrorMessage = "A communication must have at least one structure such as Interface, Protocol, Port, Device or Deposit"
    
    Set ChannelControllerList = GetChannelShapeControllers
    
    For Each CurrentChannelController In ChannelControllerList
    
        If Not ChannelHasBeenAnalyzed(CurrentChannelController) Then
            
            Result = False
    
            Set CurrentChannelViewer = CurrentChannelController.ShapeViewer
            
            Set CurrentChannelShape = CurrentChannelViewer.shape
            
            Set ChannelConnectorsList = CurrentChannelShape.Connects
            
            PublishError CurrentChannelController, ErrorMessage
            
            For Each CurrentChannelConnect In ChannelConnectorsList
                    
                Set ConnectedShape = CurrentChannelConnect.ToSheet
                
                Set ConnectedShapeController = FactoryShapeController.GetShapeControllerByShape(ConnectedShape)
                
                PublishError ConnectedShapeController, ErrorMessage
            
            Next
            
        End If
        
    Next
    
    ValidateCommunicationsWithoutStructures = Result

End Function


Private Function ChannelHasBeenAnalyzed(ByVal aChannelController As shapeController)

    Dim Result As Boolean
    Dim CurrentChannelShapeController As shapeController
    
    Result = False
    
    For Each CurrentChannelShapeController In AnalyzedChannelsList
    
        If (CurrentChannelShapeController.GetShapeUniqueId = aChannelController.GetShapeUniqueId) Then
            
            Result = True
            
            Exit For
            
        End If
    
    Next
    
    ChannelHasBeenAnalyzed = Result

End Function

Private Sub AddShapeViewerToAnalyzedChannels(ByVal aShapeController As shapeController)

    If aShapeController.ShapeSysAdlType = sysAdlTypeSetChannel Then
    
        AnalyzedChannelsList.Add aShapeController
    
    End If

End Sub

' put elements in the list of elements that belong to a structure
Private Sub AddElementsInStructureInList(ByVal aStructure As DiagramStructureSysAdl)

    Dim StructureElementItem1 As New UtilSysAdlItem
    Dim StructureElementItem2 As New UtilSysAdlItem
    
    Dim ElementKey1 As String
    Dim ElementKey2 As String
    
    ElementKey1 = aStructure.FirstShapeController.GetShapeUniqueId
    ElementKey2 = aStructure.SecondShapeController.GetShapeUniqueId
    
    StructureElementItem1.Init ElementKey1, aStructure
    StructureElementItem2.Init ElementKey2, aStructure

    If (aStructure.StructureType = sysAdlStructureConcern) Then
    
        ConcernShapeList.Add StructureElementItem1
        ConcernShapeList.Add StructureElementItem2
    
    ElseIf (aStructure.StructureType = sysAdlStructureDeposit) Then
    
        DepositShapeList.Add StructureElementItem1
        DepositShapeList.Add StructureElementItem2
    
    ElseIf (aStructure.StructureType = sysAdlStructureHost) Then
    
        HostShapeList.Add StructureElementItem1
        HostShapeList.Add StructureElementItem2
    
    ElseIf (aStructure.StructureType = sysAdlStructureInterface) Then
    
        InterfaceShapeList.Add StructureElementItem1
        InterfaceShapeList.Add StructureElementItem2
        
    ElseIf (aStructure.StructureType = sysAdlStructurePort) Then
    
        PortShapeList.Add StructureElementItem1
        PortShapeList.Add StructureElementItem2
        
    ElseIf (aStructure.StructureType = sysAdlStructureProtocol) Then
    
        ProtocolShapeList.Add StructureElementItem1
        ProtocolShapeList.Add StructureElementItem2
        
    ElseIf (aStructure.StructureType = sysAdlStructureRequirement) Then
    
        RequirementShapeList.Add StructureElementItem1
        RequirementShapeList.Add StructureElementItem2
        
    ElseIf (aStructure.StructureType = sysAdlStructureResponsibility) Then
    
        ResponsibilityShapeList.Add StructureElementItem1
        ResponsibilityShapeList.Add StructureElementItem2
        
    End If


End Sub

' Get the Port of an receiver
Private Function GetPort(ByVal aShapeViewer As ShapeViewer) As DiagramStructureSysAdl

    Dim Result As DiagramStructureSysAdl

    Set Result = Nothing

    If (aShapeViewer.GetSysAdlTypeOfShape = sysAdlTypeSetReceiver) Then
        
        Set Result = PortShapeList.Item(aShapeViewer.GetShapeId)
    
    End If

    Set GetPort = Result

End Function

' get the Protocol of an receiver
Private Function GetProtocol(ByVal aShapeController As shapeController) As DiagramStructureSysAdl

    Dim Result As DiagramStructureSysAdl
    
    Set Result = Nothing
    
    If (aShapeController.ShapeSysAdlType = sysAdlTypeSetReceiver) Then
    
        Set Result = ProtocolShapeList.Item(aShapeController.GetShapeUniqueId)
        
    End If
    
    Set GetProtocol = Result
    
End Function

' get the interface of an format
Private Function GetInterface(ByVal aShapeController As shapeController) As DiagramStructureSysAdl

    Dim Result As DiagramStructureSysAdl
    
    Set Result = Nothing
    
    If (aShapeController.ShapeSysAdlType = sysAdlTypeSetFormat) Then
    
        Set Result = InterfaceShapeList.Item(aShapeController.GetShapeUniqueId)
        
    End If
    
    Set GetInterface = Result

End Function

' get the deposit of a sender
Private Function GetDepositDeck(ByVal aShapeController As shapeController) As DiagramStructureSysAdl

    Dim Result As DiagramStructureSysAdl
    
    Set Result = Nothing
    
    If (aShapeController.ShapeSysAdlType = sysAdlTypeSetSender) Then
    
        Set Result = DepositShapeList.Item(aShapeController.GetShapeUniqueId)
        
    End If
    
    Set GetDepositDeck = Result

End Function

' get the list of connectors in diagrams (from elements)
Private Function GetConnectorControllers() As Collection

    Dim Result As Collection

    Set Result = GetShapeControllersFromDoc(sysAdlTypeSetConnector)

    Set GetConnectorControllers = Result
    
End Function

Private Function GetChannelShapeControllers() As Collection

    Dim Result As Collection
    
    Set Result = GetShapeControllersFromDoc(sysAdlTypeSetChannel)
    
    Set GetChannelShapeControllers = Result

End Function

Private Function GetRelationControllers(ByVal relationType As String)

    Dim Result As Collection
    
    Set Result = GetShapeControllersFromDoc(relationType)
    
    Set GetRelationControllers = Result

End Function

Private Function GetShapeControllersFromDoc(ByVal elementType As String) As Collection

    Dim Result As Collection
    Dim PageList As Visio.Pages
    Dim CurrentPage As Visio.page
    Dim ShapeList As Visio.Shapes
    Dim CurrentShape As Visio.shape
    Dim ShapeControllerFound As shapeController
    
    Set Result = New Collection

    Set PageList = VisioDocInAnalysis.Pages
    
    For Each CurrentPage In PageList
    
        Set ShapeList = CurrentPage.Shapes
        
        For Each CurrentShape In ShapeList
        
            Set ShapeControllerFound = FactoryShapeController.GetShapeControllerByShape(CurrentShape)
            
            If ShapeControllerFound.ShapeSysAdlType = elementType Then
            
                Result.Add ShapeControllerFound
            
            End If
        
        Next
    
    Next
    
    Set GetShapeControllersFromDoc = Result

End Function

' get the list of channels that are connected to an element
Private Function GetChannelListFromElement(ByVal aShapeController As shapeController) As Collection

    ' the list of discovered channels
    Dim Result As New Collection
    
    ' the shapes connected to the element
    Dim ShapeConnectedTo As shape
    Dim ShapeConnectedFrom As shape
    
    ' the elements derived from the shape
    Dim ElementTypeConnectedTo As String
    Dim ElementTypeConnectedFrom As String
    
    ' the controllers derived from the shape
    Dim CurrentShapeControllerTo As shapeController
    Dim CurrentShapeControllerFrom As shapeController
    
    'Connections discovered from the element
    Dim ShapeConnections As Connects
    Dim ShapeConnection As Connect
    
    ' Shape that is the source of connections
    Dim ShapeSearchParameter As shape
    
    ' get the shape of the shapeviewer sent by function caller
    Set ShapeSearchParameter = aShapeController.ShapeViewer.shape

    ' get connections of this shape
    Set ShapeConnections = ShapeSearchParameter.FromConnects
    
    ' iterate in connections discovered
    For Each ShapeConnection In ShapeConnections
        
        ' shapes connected
        Set ShapeConnectedTo = ShapeConnection.ToSheet
        Set ShapeConnectedFrom = ShapeConnection.FromSheet
    
        ' controllers of shapes connected
        Set CurrentShapeControllerTo = FactoryShapeController.GetShapeControllerByShape(ShapeConnectedTo)
        Set CurrentShapeControllerFrom = FactoryShapeController.GetShapeControllerByShape(ShapeConnectedFrom)
    
        ' types of shapes connected
        ElementTypeConnectedTo = CurrentShapeControllerTo.ShapeSysAdlType
        ElementTypeConnectedFrom = CurrentShapeControllerFrom.ShapeSysAdlType
        
        ' if the type of element discovered in the source is Channel then...
        If (ElementTypeConnectedTo = sysAdlTypeSetChannel) Then
        
            '... I add this discovered channel in function's result
            Result.Add CurrentShapeControllerTo
            
        ' if the type of element discovered in the end is Channel then
        ElseIf (ElementTypeConnectedFrom = sysAdlTypeSetChannel) Then
                        
            '... I add this discovered channel in function's result
            Result.Add CurrentShapeControllerFrom
            
        End If
        
    Next
    
    ' return the result of this search
    Set GetChannelListFromElement = Result

End Function

' show error messages to user
Private Sub PublishStructureErrors(ByVal structureToMark As DiagramStructureSysAdl, ByVal ErrorMessage As String)

    PublishError structureToMark.FirstShapeController, ErrorMessage
    PublishError structureToMark.SecondShapeController, ErrorMessage
    PublishError structureToMark.ConnectorElementShapeController, ErrorMessage
    
End Sub


' return the element that is different of LinkElement (generally is used to know which element may be in other end of a communication)
Private Function GetShapeToEvaluate(ByVal FirstShapeController As shapeController, _
                                    ByVal SecondShapeController As shapeController, _
                                    ByVal LinkElement As shapeController) As shapeController
                                    
        Dim Result As shapeController
                        
        Set Result = Nothing
                
        If (LinkElement.ShapeViewer.IsSameShapeViewer(FirstShapeController.ShapeViewer)) Then
                    
            Set Result = SecondShapeController
                        
        ElseIf (LinkElement.ShapeViewer.IsSameShapeViewer(SecondShapeController.ShapeViewer)) Then
                    
            Set Result = FirstShapeController
                    
        End If
        
        Set GetShapeToEvaluate = Result

End Function

' check if connection is valid and puts communication found in communication list
Private Sub ProcessCommunication(ByVal aSourceStructure As DiagramStructureSysAdl, _
                                  ByVal aDestinyStructure As DiagramStructureSysAdl, _
                                  ByVal aConnectorChannel As shapeController)
                                  
        ' Communication found
        Dim Communication As DiagramCommunicationSysAdl
        ' error message
        Dim ErrorMessage As String
        ' type of structure on the end of communication (strucuture has been provided by caller function
        Dim DestinyStructureType As String
        
        'communication type
        Dim CommunicationType As String
        
        ' the result of this analysis
        Dim Result As Boolean
        
        'create and initialize the communication
        Set Communication = New DiagramCommunicationSysAdl
        Communication.Init aSourceStructure, aDestinyStructure, aConnectorChannel
                
        ' if communication type is valid (not empty string) then...
        CommunicationType = Communication.CommunicationType
        
        If CommunicationType <> sysAdlStringConstantsEmpty Then
                        
           '... add communication to the list of communications
           analysisResult.AddCommunication Communication
           
        ' but, if communication is not valid...
        Else
            
            ' ... well... first, I discover the type of structure...
            DestinyStructureType = aDestinyStructure.StructureType
            
            '... and, if this structure type is a Single Element...
            If (DestinyStructureType = sysAdlStructureSingleElement) Then
                
                '... I discover the type of element of this structure (I must be only one - System, Format, Receiver etc...)
                DestinyStructureType = aDestinyStructure.GetSingleElementType
                
                ' a little explain: I do this to show a message to the user (for instance: "An Interface can't communicates to a System")
                    
            End If
               
            ' Create the message error to be show to the user
            ErrorMessage = "It is not possible create a communication between " + aSourceStructure.StructureType + " and " + DestinyStructureType
                            
            ' Publish error in source structure
            PublishStructureErrors aSourceStructure, ErrorMessage
                            
            ' Publish error in the channel that connects these structures
            PublishError aConnectorChannel, ErrorMessage
            
            ' Publish error in other end structure...
            
            ' if it is a single element, then I publish the error only in the shape
            If (aDestinyStructure.StructureType = sysAdlStructureSingleElement) Then
                
                ' (because it has no connectos and if I try to publish erro in the structure it will cause an error)
                PublishError aDestinyStructure.FirstShapeController, ErrorMessage
                
            Else
                ' publish errors in all structure
                PublishStructureErrors aDestinyStructure, ErrorMessage
                
            End If
                    
             Result = False
                        
         End If
        
End Sub





' publish error in a single element (used by PublishStructureErrors)
Private Sub PublishError(ByVal aShapeController As shapeController, ByVal aMessage As String)
    
    Dim IssueFound As DiagramShapeIssues
    Dim ListItem As UtilSysAdlItem
    
    Set IssueFound = New DiagramShapeIssues
    Set ListItem = New UtilSysAdlItem
    
    IssueFound.Init aShapeController.GetShapeUniqueId, aShapeController.ShapeSysAdlType
    IssueFound.AddIssue aMessage
    
    ListItem.Init IssueFound.ShapeId, IssueFound
    
    analysisResult.AddError ListItem
    
End Sub

