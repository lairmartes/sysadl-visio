VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiagramAnalysisResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    'constants to collapse or not
    
    Private Const TAG_COLLAPSED = True
    
    'list of elements found
    Private ChannelList As Collection
    Private FormatList As Collection
    Private LayerList As Collection
    Private NodeList As Collection
    Private QualityList As Collection
    Private ReceiverList As Collection
    Private SenderList As Collection
    Private StakeholderList As Collection
    Private SystemList As Collection
    Private TransitionList As Collection
    Private RoleList As Collection
    Private DecisionList As Collection
    Private AssumptionList As Collection
    
    'list of related objectives
    Private ObjectiveList As Collection

    'list of structures found
    Private ProtocolStructureList As Collection
    Private InterfaceStructureList As Collection
    Private HostStructureList As Collection
    Private PortStructureList As Collection
    Private DepositStructureList As Collection
    Private ConcernStructureList As Collection
    Private DeviceStructureList As Collection
    Private RequirementStructureList As Collection
    Private ResponsibilityStructureList As Collection
    Private RationaleStructureList As Collection
    Private DefinitionStructureList As Collection

    'list of relations found
    Private IsARelationList As Collection
    Private ComposedByRelationList As Collection
    Private RepresentsRelationList As Collection
    Private DependsOnRelationList As Collection
    Private ComposedByInnerRelationList As Collection
    Private TransitionRelationList As Collection
        
    'list of communications found
    Private NetCommunicationList As Collection
    Private ConversationCommunicationList As Collection
    Private InstallationCommunicationList As Collection
    
    'conversations id already included (check it when including a communication)
    Private ConversationsIncluded As UtilSysAdlList
    
    'list of errors
    Private ErrorList As UtilSysAdlList
    
    'list of elements in diagram (no matter type - plus shape!)
    Private DiagramElementList As Collection
    
    'name of diagram
    Private pDiagramFileName As String
    Private pDiagramQualifier As String
    Private pExportXMLFileName

    'object file for exporting
    Dim pFileExportingStream As Stream


Public Sub Init(ByVal fileName As String)

    'initializing list of elements
    Set ChannelList = New Collection
    Set FormatList = New Collection
    Set LayerList = New Collection
    Set NodeList = New Collection
    Set QualityList = New Collection
    Set ReceiverList = New Collection
    Set SenderList = New Collection
    Set StakeholderList = New Collection
    Set SystemList = New Collection
    Set TransitionList = New Collection
    Set RoleList = New Collection
    Set DecisionList = New Collection
    Set AssumptionList = New Collection
    
    'initialize objective list
    Set ObjectiveList = New Collection
    
    'initializing list of structures
    Set ProtocolStructureList = New Collection
    Set InterfaceStructureList = New Collection
    Set HostStructureList = New Collection
    Set PortStructureList = New Collection
    Set DepositStructureList = New Collection
    Set ConcernStructureList = New Collection
    Set DeviceStructureList = New Collection
    Set RequirementStructureList = New Collection
    Set ResponsibilityStructureList = New Collection
    Set RationaleStructureList = New Collection
    Set DefinitionStructureList = New Collection
        
    'initializing list of relations
    Set IsARelationList = New Collection
    Set ComposedByRelationList = New Collection
    Set RepresentsRelationList = New Collection
    Set DependsOnRelationList = New Collection
    Set ComposedByInnerRelationList = New Collection
    Set TransitionRelationList = New Collection
    
    'initializing list of communications
    Set NetCommunicationList = New Collection
    Set ConversationCommunicationList = New Collection
    Set InstallationCommunicationList = New Collection
    
    'initializing list of conversations included
    Set ConversationsIncluded = New UtilSysAdlList
    
    'list of all elements (with shape)
    Set DiagramElementList = New Collection
    
    'initializing list of errors
    Set ErrorList = New UtilSysAdlList
    
    'initialize diagram qualifier
    pDiagramFileName = fileName
    pDiagramQualifier = GUIServices.GetQualifierFromDocumentName(fileName)
    pExportXMLFileName = GUIServices.ChangeDocumentExtension(fileName, sysadlstringconstantsExtensionSysAdl)

    
End Sub
    
    
    ' add structure in the correct list depending of its type
Public Sub AddStructure(ByVal aStructure As DiagramStructureSysAdl)
    
    If aStructure.StructureType = sysAdlStructureProtocol Then ProtocolStructureList.Add aStructure
    
    If aStructure.StructureType = sysAdlStructureInterface Then InterfaceStructureList.Add aStructure
    
    If aStructure.StructureType = sysAdlStructureConcern Then ConcernStructureList.Add aStructure
    
    If aStructure.StructureType = sysAdlStructureHost Then HostStructureList.Add aStructure
    
    If aStructure.StructureType = sysAdlStructurePort Then PortStructureList.Add aStructure
    
    If aStructure.StructureType = sysAdlStructureDeposit Then DepositStructureList.Add aStructure
    
    If aStructure.StructureType = sysAdlStructureDevice Then DeviceStructureList.Add aStructure
    
    If aStructure.StructureType = sysAdlStructureRequirement Then RequirementStructureList.Add aStructure
    
    If aStructure.StructureType = sysAdlStructureResponsibility Then ResponsibilityStructureList.Add aStructure
    
    If aStructure.StructureType = sysAdlStructureRationale Then RationaleStructureList.Add aStructure
    
    If aStructure.StructureType = sysAdlStructureDefinition Then DefinitionStructureList.Add aStructure


End Sub
' add communications depending of its type
Public Sub AddCommunication(ByVal aCommunication As DiagramCommunicationSysAdl)

    Dim ConversationIncluded As UtilSysAdlItem

    If aCommunication.CommunicationType = sysAdlCommunicationConversation Then
    
        If ConversationsIncluded.Item(aCommunication.Id) Is Nothing Then
        
            Set ConversationIncluded = New UtilSysAdlItem
            
            ConversationIncluded.Init aCommunication.Id, aCommunication
    
            ConversationCommunicationList.Add aCommunication
            
            ConversationsIncluded.Add ConversationIncluded
            
        End If
            
     End If
            
    If aCommunication.CommunicationType = sysAdlCommunicationNet Then NetCommunicationList.Add aCommunication
    If aCommunication.CommunicationType = sysAdlCommunicationInstallation Then InstallationCommunicationList.Add aCommunication

End Sub

'add relations and transitions (they are the same nature)
Public Sub AddRelation(ByVal aRelation As DiagramRelationSysAdl)

    If aRelation.Relation = sysAdlTypeSetDependsOn Then DependsOnRelationList.Add aRelation
    If aRelation.Relation = sysAdlTypeSetIsA Then IsARelationList.Add aRelation
    If aRelation.Relation = sysAdlTypeSetRepresents Then RepresentsRelationList.Add aRelation
    If aRelation.Relation = sysAdlTypeSetComposedBy Then ComposedByRelationList.Add aRelation
    ' add transitions here too
    If aRelation.Relation = sysAdlTypeSetTransition Then TransitionRelationList.Add aRelation


End Sub

'add inner relation
Public Sub AddInnerRelation(ByVal anInnerRelation As DiagramInnerRelationSysADL)

    ComposedByInnerRelationList.Add anInnerRelation

End Sub

' add error messages found in analysis for further use (like list of errors)
Public Sub AddError(ByVal anIssue As UtilSysAdlItem)

    ErrorList.Add anIssue
    
End Sub


Public Sub AddObjective(ByVal ObjectiveShapeController As shapeController)

    ObjectiveList.Add ObjectiveShapeController

End Sub
Public Sub AddElement(ByVal ElementController As shapeController)

    If (ElementController.ShapeSysAdlType = sysAdlTypeSetChannel) Then ChannelList.Add ElementController
    If (ElementController.ShapeSysAdlType = sysAdlTypeSetFormat) Then FormatList.Add ElementController
    If (ElementController.ShapeSysAdlType = sysAdlTypeSetLayer) Then LayerList.Add ElementController
    If (ElementController.ShapeSysAdlType = sysAdlTypeSetNode) Then NodeList.Add ElementController
    If (ElementController.ShapeSysAdlType = sysAdlTypeSetQuality) Then QualityList.Add ElementController
    If (ElementController.ShapeSysAdlType = sysAdlTypeSetReceiver) Then ReceiverList.Add ElementController
    If (ElementController.ShapeSysAdlType = sysAdlTypeSetSender) Then SenderList.Add ElementController
    If (ElementController.ShapeSysAdlType = sysAdlTypeSetStakeholder) Then StakeholderList.Add ElementController
    If (ElementController.ShapeSysAdlType = sysAdlTypeSetSystem) Then SystemList.Add ElementController
    If (ElementController.ShapeSysAdlType = sysAdlTypeSetTransition) Then TransitionList.Add ElementController
    If (ElementController.ShapeSysAdlType = sysAdlTypeSetRole) Then RoleList.Add ElementController
    If (ElementController.ShapeSysAdlType = sysAdlTypeSetDecision) Then DecisionList.Add ElementController
    If (ElementController.ShapeSysAdlType = sysAdlTypeSetAssumption) Then AssumptionList.Add ElementController

End Sub

Public Sub AddDiagramElement(ByVal DiagramElement As DiagramElement)

    DiagramElementList.Add DiagramElement

End Sub

Public Property Get Protocols() As Collection

    Set Protocols = ProtocolStructureList
    
End Property

Public Property Get Interfaces() As Collection

    Set Interfaces = InterfaceStructureList
    
End Property


Public Property Get Hosts() As Collection

    Set Hosts = HostStructureList
    
End Property

Public Property Get Ports() As Collection

    Set Ports = PortStructureList
    
End Property

Public Property Get Deposits() As Collection

    Set Deposits = DepositStructureList

End Property

Public Property Get Devices() As Collection

    Set Devices = DeviceStructureList

End Property

Public Property Get Requirements() As Collection

    Set Requirements = RequirementStructureList

End Property
Public Property Get Concerns() As Collection

    Set Concerns = ConcernStructureList

End Property

Public Property Get Objectives() As Collection

    Set Objectives = ObjectiveList

End Property

Public Property Get Responsibilities() As Collection

    Set Responsibilities = ResponsibilityStructureList
    
End Property

Public Property Get Rationales() As Collection

    Set Rationales = RationaleStructureList

End Property

Public Property Get Definitions() As Collection

    Set Definitions = DefinitionStructureList
    
End Property


Public Property Get Errors() As UtilSysAdlList
    
    Set Errors = ErrorList

End Property

'elements collected
Public Property Get Channels() As Collection

    Set Channels = ChannelList

End Property

Public Property Get Formats() As Collection

    Set Formats = FormatList

End Property

Public Property Get Layers() As Collection

    Set Layers = LayerList

End Property

Public Property Get Nodes() As Collection

    Set Nodes = NodeList
    
End Property



Public Property Get Qualities() As Collection

    Set Qualities = QualityList

End Property

Public Property Get Receivers() As Collection

    Set Receivers = ReceiverList

End Property

Public Property Get Roles() As Collection

    Set Roles = RoleList

End Property

Public Property Get Senders() As Collection

    Set Senders = SenderList

End Property

Public Property Get Stakeholders() As Collection

    Set Stakeholders = StakeholderList

End Property

Public Property Get Systems() As Collection

    Set Systems = SystemList

End Property

Public Property Get Transitions() As Collection

    Set Transitions = TransitionList

End Property

Public Property Get Decisions() As Collection

    Set Decisions = DecisionList

End Property

Public Property Get Assumptions() As Collection

    Set Assumptions = AssumptionList
    
End Property


Public Property Get IsARelations() As Collection

    Set IsARelations = IsARelationList

End Property

Public Property Get TransitionRelations() As Collection

    Set TransitionRelations = TransitionRelationList

End Property

Public Property Get ComposedByRelations() As Collection

    Set ComposedByRelations = ComposedByRelationList

End Property

Public Property Get ComposedByInnerRelations() As Collection

    Set ComposedByInnerRelations = ComposedByInnerRelationList

End Property

Public Property Get DependsOnRelations() As Collection

    Set DependsOnRelations = DependsOnRelationList

End Property

Public Property Get RepresentsRelations() As Collection

    Set RepresentsRelations = RepresentsRelationList

End Property

Public Property Get NetCommunications() As Collection

    Set NetCommunications = NetCommunicationList

End Property

Public Property Get ConversationCommunications() As Collection

    Set ConversationCommunications = ConversationCommunicationList

End Property

Public Property Get InstallationCommuncations() As Collection

    Set InstallationCommuncations = InstallationCommunicationList

End Property

Public Property Get DiagramElements() As Collection

    Set DiagramElements = DiagramElementList

End Property


' check if has been added errors messages
Public Property Get HasError() As Boolean

    Dim Result As Boolean
    
    Result = False
    
    If (ErrorList.Count > 0) Then Result = True
    
    HasError = Result

End Property

' check if diagram is ok (there's no errors and has at least one objective)
Public Property Get DiagramIsOk() As Boolean

    Dim Result As Boolean
    
    Result = True
    
    If (ErrorList.Count > 0 Or Not HasObjectives) Then Result = False
    
    
    DiagramIsOk = Result

End Property

'check if diagram has at least one objective
Public Property Get HasObjectives() As Boolean

    Dim Result As Boolean
    
    Result = False
    
    If (ObjectiveList.Count > 0) Then Result = True
    
    HasObjectives = Result

End Property

Public Function isDiagramNameMatch(ByVal fileName As String)

    Dim Result As Boolean
    
    Result = (pDiagramFileName = fileName)
    
    isDiagramNameMatch = Result
    
End Function

