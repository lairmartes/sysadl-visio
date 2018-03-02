VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiagramStructureSysAdl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pFirstShapeController As shapeController
    Private pSecondShapeController As shapeController
    Private pConnectorShapeController As shapeController
    Private isSingleElement As Boolean
    
    Private Const TYPE_SEPARATOR = ","
    
    Private Const ALLOWED_TYPE_CONNECTED_DEPOSIT = TYPE_SEPARATOR + sysAdlTypeSetSystem + TYPE_SEPARATOR
    Private Const ALLOWED_TYPE_CONNECTED_STAKEHOLDER = TYPE_SEPARATOR + _
                                                       sysAdlTypeSetFormat + TYPE_SEPARATOR + _
                                                       sysAdlTypeSetNode + TYPE_SEPARATOR + _
                                                       sysAdlTypeSetObjective + TYPE_SEPARATOR + _
                                                       sysAdlTypeSetReceiver + TYPE_SEPARATOR + _
                                                       sysAdlTypeSetSender + TYPE_SEPARATOR + _
                                                       sysAdlTypeSetSystem + _
                                                       TYPE_SEPARATOR
                                                       
    Private Const ALLOWED_TYPE_CONNECTED_CONCERN = TYPE_SEPARATOR + _
                                                   sysAdlTypeSetRole + TYPE_SEPARATOR + _
                                                   sysAdlTypeSetQuality + _
                                                   TYPE_SEPARATOR
                                                   
    Private Const ALLOWED_TYPE_CONNECTED_DECISION = TYPE_SEPARATOR + _
                                                    sysAdlTypeSetSystem + TYPE_SEPARATOR + _
                                                    sysAdlTypeSetNode + _
                                                    TYPE_SEPARATOR
                                                    
    Private Const ALLOWED_TYPE_CONNECTED_DEFINITION = TYPE_SEPARATOR + _
                                                   sysAdlTypeSetAssumption + _
                                                   TYPE_SEPARATOR
                                                   
                                                   
       
    Public Sub Init(ByVal aFirstShapeController As shapeController, _
                    ByVal aSecondShapeController As shapeController, _
                    ByVal aConnectorShapeController As shapeController)
    
        Set pFirstShapeController = aFirstShapeController
        Set pSecondShapeController = aSecondShapeController
        Set pConnectorShapeController = aConnectorShapeController
        
        isSingleElement = False
    
    End Sub
    
    Public Sub InitSingleElement(ByVal anElementShapecontroller As shapeController)
                                 
                                 
            Init anElementShapecontroller, anElementShapecontroller, Nothing
            
            isSingleElement = True
                                 
    End Sub
    
    Public Property Get Id() As String
    
        Dim Result As String
        
        If Not isSingleElement Then
        
            Result = pConnectorShapeController.GetShapeUniqueId
            
        Else
            
            Result = pFirstShapeController.GetShapeUniqueId
            
        End If
        
        Id = Result
    
    End Property
    
    Public Property Get FirstElementType() As String

         FirstElementType = pFirstShapeController.ShapeSysAdlType

    End Property
    
    Public Property Get FirstShapeController() As shapeController
    
        Set FirstShapeController = pFirstShapeController
    
    End Property
    
    Public Property Get SecondElementType() As String
    
         SecondElementType = pSecondShapeController.ShapeSysAdlType
    
    End Property
    
    Public Property Get SecondShapeController() As shapeController
    
        Set SecondShapeController = pSecondShapeController
    
    End Property
    
   
   Public Property Get ConnectorElementType() As String
   
         ConnectorElementType = pConnectorShapeController.ShapeSysAdlType
   
   End Property
   
   Public Property Get ConnectorElementShapeController() As shapeController

        Set ConnectorElementShapeController = pConnectorShapeController

   End Property
   
Public Property Get Key() As String

    Key = pConnectorShapeController.ShapeViewer.GetShapeId

End Property
   
Public Property Get StructureType() As String

    Dim Result As String
    Dim ConnectorElementType As String
    
    Result = sysAdlStringConstantsEmpty
    
    If Not isSingleElement Then
    
       ConnectorElementType = ConnectorElementShapeController.ShapeSysAdlType
        
        If (ConnectorElementType = sysAdlTypeSetConnector) Then
        
            Result = CalculateStructureTypeForConnections()
        
        End If
    Else
        
        Result = sysAdlStructureSingleElement
        
    End If
    
    StructureType = Result

End Property
   
Private Function CalculateStructureTypeForConnections()

    Dim shape1SysAdlType As String
    Dim shape2SysAdlType As String
    Dim Result As String
    
    shape1SysAdlType = FirstShapeController.ShapeSysAdlType
    shape2SysAdlType = SecondShapeController.ShapeSysAdlType
    
    Result = sysAdlStringConstantsEmpty
    
    If IsProtocol(shape1SysAdlType, shape2SysAdlType) Then
        
        Result = sysAdlStructureProtocol
        
    ElseIf IsInterface(shape1SysAdlType, shape2SysAdlType) Then
    
        Result = sysAdlStructureInterface
    
    ElseIf IsDeposit(shape1SysAdlType, shape2SysAdlType) Then
    
        Result = sysAdlStructureDeposit
        
    ElseIf IsHost(shape1SysAdlType, shape2SysAdlType) Then
    
        Result = sysAdlStructureHost
    
    ElseIf IsConcern(shape1SysAdlType, shape2SysAdlType) Then
    
        Result = sysAdlStructureConcern
        
    ElseIf IsPort(shape1SysAdlType, shape2SysAdlType) Then
    
        Result = sysAdlStructurePort
        
    ElseIf IsDevice(shape1SysAdlType, shape2SysAdlType) Then
    
        Result = sysAdlStructureDevice
        
    ElseIf IsRequirement(shape1SysAdlType, shape2SysAdlType) Then

        Result = sysAdlStructureRequirement
        
    ElseIf IsResponsibility(shape1SysAdlType, shape2SysAdlType) Then
    
        Result = sysAdlStructureResponsibility
        
    ElseIf IsRationale(shape1SysAdlType, shape2SysAdlType) Then
    
        Result = sysAdlStructureRationale
        
    ElseIf IsDefinition(shape1SysAdlType, shape2SysAdlType) Then
    
        Result = sysAdlStructureDefinition
    
    End If
    
    CalculateStructureTypeForConnections = Result
    

End Function


Private Function IsProtocol(ByVal elementType1 As String, ByVal elementType2 As String)

    Dim Result As Boolean
    
    Result = False
    
    If ((elementType1 = sysAdlTypeSetSystem And elementType2 = sysAdlTypeSetReceiver)) Then
    
        Result = True
        
    End If
    
    If ((elementType1 = sysAdlTypeSetReceiver And elementType2 = sysAdlTypeSetSystem)) Then
    
        Result = True
        
    End If
    
    IsProtocol = Result

End Function


Private Function IsInterface(ByVal elementType1 As String, ByVal elementType2 As String)

    Dim Result As Boolean
    
    Result = False
    
    If ((elementType1 = sysAdlTypeSetSystem And elementType2 = sysAdlTypeSetFormat)) Then
    
        Result = True
        
    End If
    
    If ((elementType1 = sysAdlTypeSetFormat And elementType2 = sysAdlTypeSetSystem)) Then
    
        Result = True
        
    End If
    
    IsInterface = Result

End Function


Private Function IsDeposit(ByVal elementType1 As String, ByVal elementType2 As String)

    Dim Result As Boolean
    
    Result = False
    
    If ((elementType1 = sysAdlTypeSetSender And elementType2 = sysAdlTypeSetFormat) And (IsConnectedTo(pSecondShapeController, ALLOWED_TYPE_CONNECTED_DEPOSIT))) Then
    
        Result = True
        
    End If
    
    If ((elementType1 = sysAdlTypeSetFormat And elementType2 = sysAdlTypeSetSender) And (IsConnectedTo(pFirstShapeController, ALLOWED_TYPE_CONNECTED_DEPOSIT))) Then
    
        Result = True
        
    End If
    
    IsDeposit = Result

End Function

Private Function IsPort(ByVal elementType1 As String, ByVal elementType2 As String)

    Dim Result As Boolean
    
    Result = False
    
    If ((elementType1 = sysAdlTypeSetReceiver And elementType2 = sysAdlTypeSetNode)) Then
    
        Result = True
        
    End If
    
    If ((elementType1 = sysAdlTypeSetNode And elementType2 = sysAdlTypeSetReceiver)) Then
    
        Result = True
        
    End If
    
    IsPort = Result

End Function
Private Function IsDevice(ByVal elementType1 As String, ByVal elementType2 As String)

    Dim Result As Boolean
    
    Result = False
    
    If ((elementType1 = sysAdlTypeSetSender And elementType2 = sysAdlTypeSetNode)) Then
    
        Result = True
        
    End If
    
    If ((elementType1 = sysAdlTypeSetNode And elementType2 = sysAdlTypeSetSender)) Then
    
        Result = True
        
    End If
    
    IsDevice = Result

End Function

Private Function IsHost(ByVal elementType1 As String, ByVal elementType2 As String)

    Dim Result As Boolean
    
    Result = False
    
    If ((elementType1 = sysAdlTypeSetSystem And elementType2 = sysAdlTypeSetNode)) Then
    
        Result = True
        
    End If
    
    If ((elementType1 = sysAdlTypeSetNode And elementType2 = sysAdlTypeSetSystem)) Then
    
        Result = True
        
    End If
    
    IsHost = Result

End Function

Private Function IsConcern(ByVal elementType1 As String, ByVal elementType2 As String)

    Dim Result As Boolean
    Dim ElementTypeConcerned As String
    Dim IsElementAbleToBeConcerned As Boolean
    Dim IsThereReasonForConcern As Boolean
    Dim StakeholderShapeController As shapeController
    
    Result = False
    
    ElementTypeConcerned = sysAdlStringConstantsEmpty
    IsElementAbleToBeConcerned = False
    IsThereReasonForConcern = False
    
    If (elementType1 = sysAdlTypeSetStakeholder) Then
    
        ElementTypeConcerned = elementType2
        Set StakeholderShapeController = pFirstShapeController
    
    ElseIf (elementType2 = sysAdlTypeSetStakeholder) Then
        
        ElementTypeConcerned = elementType1
        Set StakeholderShapeController = pSecondShapeController
        
    End If
    
    
    If ElementTypeConcerned <> sysAdlStringConstantsEmpty Then
    
        IsElementAbleToBeConcerned = IsValueInStringList(ElementTypeConcerned, ALLOWED_TYPE_CONNECTED_STAKEHOLDER)
        
    End If
    
    If IsElementAbleToBeConcerned Then
    
        IsThereReasonForConcern = (IsConnectedTo(StakeholderShapeController, ALLOWED_TYPE_CONNECTED_CONCERN))
        
        If (IsThereReasonForConcern) Then Result = True
        
    End If
    
    IsConcern = Result

End Function

Private Function IsResponsibility(ByVal elementType1 As String, ByVal elementType2 As String)

    Dim Result As Boolean
    Dim ElementTypeResponsible As String
    
    Result = False
    
    ElementTypeResponsible = sysAdlStringConstantsEmpty
    
    If (elementType1 = sysAdlTypeSetRole) Then
    
        ElementTypeResponsible = elementType2
    
    ElseIf (elementType2 = sysAdlTypeSetRole) Then
        
        ElementTypeResponsible = elementType1
        
    End If
    
    If ElementTypeResponsible <> sysAdlStringConstantsEmpty Then
  
        If (ElementTypeResponsible = sysAdlTypeSetStakeholder) Then Result = True
        
    End If
    
    IsResponsibility = Result

End Function

'========================================================

Private Function IsRequirement(ByVal elementType1 As String, ByVal elementType2 As String)

    Dim Result As Boolean
    Dim ElementTypeRequired As String
    Dim StakeholderShapeController As shapeController
    
    Result = False
    
    ElementTypeRequired = sysAdlStringConstantsEmpty
    
    If (elementType1 = sysAdlTypeSetQuality) Then
    
        ElementTypeRequired = elementType2
        Set StakeholderShapeController = pSecondShapeController
    
    ElseIf (elementType2 = sysAdlTypeSetQuality) Then
        
        ElementTypeRequired = elementType1
        Set StakeholderShapeController = pFirstShapeController
        
    End If
    
    If ElementTypeRequired <> sysAdlStringConstantsEmpty Then
  
        If ((ElementTypeRequired = sysAdlTypeSetStakeholder) And (IsConnectedTo(StakeholderShapeController, ALLOWED_TYPE_CONNECTED_STAKEHOLDER))) Then Result = True
        
    End If
    
    IsRequirement = Result

End Function

'========================================================

Private Function IsRationale(ByVal elementType1 As String, ByVal elementType2 As String)

    Dim Result As Boolean
    Dim ElementTypeRequired As String
    Dim DecisionShapeController As shapeController
    
    Result = False
    
    ElementTypeRequired = sysAdlStringConstantsEmpty
    
    If (elementType1 = sysAdlTypeSetAssumption) Then
    
        ElementTypeRequired = elementType2
        Set DecisionShapeController = pSecondShapeController
    
    ElseIf (elementType2 = sysAdlTypeSetAssumption) Then
        
        ElementTypeRequired = elementType1
        Set DecisionShapeController = pFirstShapeController
        
    End If
    
    If ElementTypeRequired <> sysAdlStringConstantsEmpty Then
  
        If ((ElementTypeRequired = sysAdlTypeSetDecision) And (IsConnectedTo(DecisionShapeController, ALLOWED_TYPE_CONNECTED_DECISION))) Then Result = True
        
    End If
    
    IsRationale = Result

End Function


'========================================================

Private Function IsDefinition(ByVal elementType1 As String, ByVal elementType2 As String)

    Dim Result As Boolean
    Dim ElementTypeDefined As String
    Dim IsElementAbleToBeDefined As Boolean
    Dim IsThereRationaleForDefinition As Boolean
    Dim DecisionShapeController As shapeController
    
    Result = False
    
    ElementTypeDefined = sysAdlStringConstantsEmpty
    IsElementAbleToBeDefined = False
    IsThereRationaleForDefinition = False
    
    If (elementType1 = sysAdlTypeSetDecision) Then
    
        ElementTypeDefined = elementType2
        Set DecisionShapeController = pFirstShapeController
    
    ElseIf (elementType2 = sysAdlTypeSetDecision) Then
        
        ElementTypeDefined = elementType1
        Set DecisionShapeController = pSecondShapeController
        
    End If
    
    
    If ElementTypeDefined <> sysAdlStringConstantsEmpty Then
    
        IsElementAbleToBeDefined = IsValueInStringList(ElementTypeDefined, ALLOWED_TYPE_CONNECTED_DECISION)
        
    End If
    
    If IsElementAbleToBeDefined Then
    
        IsThereRationaleForDefinition = (IsConnectedTo(DecisionShapeController, ALLOWED_TYPE_CONNECTED_DEFINITION))
        
        If (IsThereRationaleForDefinition) Then Result = True
        
    End If
    
    IsDefinition = Result

End Function

'========================================================

Public Function GetShapeControllerBySysAdlType(ByVal sysadlType As String) As shapeController


    Dim Result As shapeController
    
    Dim shapeType1st As String
    Dim shapeType2nd As String
    
    shapeType1st = pFirstShapeController.ShapeSysAdlType
    shapeType2nd = pSecondShapeController.ShapeSysAdlType
    
    Set Result = Nothing
    
    If (shapeType1st = sysadlType) Then
    
        Set Result = pFirstShapeController
        
    ElseIf (shapeType2nd = sysadlType) Then
    
        Set Result = pSecondShapeController
        
    End If
    
    Set GetShapeControllerBySysAdlType = Result

End Function

Public Function GetShapeControllerDifferentSysAdlType(ByVal sysadlType As String) As shapeController


    Dim Result As shapeController
    
    Dim shapeType1st As String
    Dim shapeType2nd As String
    
    shapeType1st = pFirstShapeController.ShapeSysAdlType
    shapeType2nd = pSecondShapeController.ShapeSysAdlType
    
    Set Result = Nothing
    
    If (shapeType1st = sysadlType) Then
    
        Set Result = pSecondShapeController
        
    Else
    
        Set Result = pFirstShapeController
        
    End If
    
    Set GetShapeControllerDifferentSysAdlType = Result

End Function

Public Function GetSingleElementController() As shapeController

    Dim Result As shapeController
    
    Set Result = Nothing
    
    If isSingleElement Then
    
        Set Result = pFirstShapeController
        
    End If
    
    Set GetSingleElementController = Result

End Function

Public Function GetSingleElementType() As String

    Dim Result As String
    
    Result = sysAdlStringConstantsEmpty
    
    If isSingleElement Then
    
        Result = pFirstShapeController.ShapeSysAdlType
        
    End If
    
    GetSingleElementType = Result

End Function

Private Function IsConnectedTo(ByVal aShapeController As shapeController, ByVal typeList As String) As Boolean

    Dim Result As Boolean

    Dim ShapeConnections As Visio.Connects
    Dim ShapesConnected As Visio.Connects
    
    Dim CurrentConnection As Visio.Connect
    Dim CurrentConnectedShape1 As IVShape
    Dim CurrentConnectedShape2 As IVShape
    
    Dim CurrentConnectionShape As IVShape
    
    
    Dim ShapeControlled As IVShape
    Dim ConnectedShapeController1 As shapeController
    Dim ConnectedShapeController2 As shapeController
    
    Dim IsConnector As Boolean
    Dim IsConnectedToAllowedType As Boolean
    
    Set ShapeControlled = aShapeController.shape
    

    
    Set ShapeConnections = ShapeControlled.FromConnects
    
    Result = False
    
    For Each CurrentConnection In ShapeConnections
    
        IsConnector = InStr("Connector", Left(CurrentConnection.FromSheet.NameU, Len("Connector"))) > 0
        
        If IsConnector Then
        
            Set CurrentConnectionShape = FactoryShapeController.GetShapeControllerByShape(CurrentConnection.FromSheet).shape
        
            Set ShapesConnected = CurrentConnectionShape.Connects
            
            If ShapesConnected.Count > 1 Then
            
                Set CurrentConnectedShape1 = ShapesConnected.Item(1).ToSheet
                Set CurrentConnectedShape2 = ShapesConnected.Item(2).ToSheet
                
                Set ConnectedShapeController1 = FactoryShapeController.GetShapeControllerByShape(CurrentConnectedShape1)
                Set ConnectedShapeController2 = FactoryShapeController.GetShapeControllerByShape(CurrentConnectedShape2)
                
                IsConnectedToAllowedType = IsValueInStringList(ConnectedShapeController1.ShapeSysAdlType, typeList)
                
                If Not IsConnectedToAllowedType Then  'I haven't found it yet... let's try again...
                
                    IsConnectedToAllowedType = IsValueInStringList(ConnectedShapeController2.ShapeSysAdlType, typeList)
                    
                End If
                
                If IsConnectedToAllowedType Then
                
                    Result = True
                    
                    Exit For
                    
                End If
                
            End If
            
        End If
            
        
    Next
    
    IsConnectedTo = Result


End Function


Private Function IsValueInStringList(ByVal SearchedValue As String, ByVal StringList As String) As Boolean

    Dim Result As Boolean
    Dim CurrentString As Variant
    Dim ValueSearchedPositionInString As Integer
    
    Result = False
    
    SearchedValue = TYPE_SEPARATOR + SearchedValue + TYPE_SEPARATOR
    
    ValueSearchedPositionInString = InStr(StringList, SearchedValue)
    
    If ValueSearchedPositionInString > 0 Then Result = True

    IsValueInStringList = Result

End Function

