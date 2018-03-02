VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandChangeKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICommand
    
    Private MyShapeController As shapeController
    Private NewNamespace As String
    Private NewId As String
    
Public Sub Init(ByVal aShapeController As shapeController, ByVal aNamespace As String, ByVal anId As String)
    
    Set MyShapeController = aShapeController
    NewNamespace = aNamespace
    NewId = anId

End Sub

Public Sub ICommand_Execute()

    Dim OpeningElement As SysAdlElement
    Dim ElementBeforeKeyChange As SysAdlElement
    Dim ElementFound As Boolean
    Dim FieldInvalidCorrectedEvent As EventSysAdl
    Dim ElementCanBeUsed As Boolean
    
    Dim ElementQualificationCompleted As Boolean
    
    Set ElementBeforeKeyChange = New SysAdlElement
    ElementBeforeKeyChange.InitWithSysAdlElement MyShapeController.SysAdlElement
    
    Set OpeningElement = New SysAdlElement
    
    'configure element key to be opened
    OpeningElement.Init MyShapeController.ShapeSysAdlType, sysAdlStringConstantsEmpty
    OpeningElement.ChangeFieldValue sysAdlKeyCustPropRowNameNamespace, NewNamespace
    OpeningElement.ChangeFieldValue sysAdlKeyCustPropRowNameId, NewId
    
    'Get element data
    Set OpeningElement = ElementServiceCache.OpenElement(OpeningElement)
    
    'check if element has been found
    ElementFound = Not (OpeningElement Is Nothing)
    
    If ElementFound Then
        'check if element is same type (I can't open a Receiver as a System)
        ElementCanBeUsed = KeyCanBeUsed(OpeningElement)
        If Not ElementCanBeUsed Then
            Dim OtherElementKeyUsedEvent As EventSysAdl
            Set OtherElementKeyUsedEvent = FactoryEvent.CreateEvent(sysAdlEventKeyUsedOtherType)
            
            ShowKeyIssueToUser
            
            MyShapeController.TriggerEvent OtherElementKeyUsedEvent
        Else
            Dim ElementRecoveredEvent As EventSysAdl
            MyShapeController.ChangeSysAdlElement OpeningElement
            ElementServiceCache.ProcessElementRemove ElementBeforeKeyChange

            'create and initialize event
            Set ElementRecoveredEvent = FactoryEvent.CreateEvent(sysAdlEventElementRecovered)
            'trigger element changed event
            MyShapeController.TriggerEvent ElementRecoveredEvent
        End If
    Else
        Dim ElementCreated As EventSysAdl
        Dim NewElement As SysAdlElement

        Set ElementCreated = FactoryEvent.CreateEvent(sysAdlEventCreatedSysAdlElement)
        
        Set NewElement = New SysAdlElement
        NewElement.InitWithSysAdlElement MyShapeController.SysAdlElement
                
        NewElement.ChangeFieldValue sysAdlKeyCustPropRowNameNamespace, NewNamespace
        NewElement.ChangeFieldValue sysAdlKeyCustPropRowNameId, NewId
        
        MyShapeController.ChangeSysAdlElement NewElement
        ElementServiceCache.ProcessElementRemove ElementBeforeKeyChange
        
        ElementQualificationCompleted = (NewElement.Id <> sysAdlStringConstantsEmpty And _
                                          NewElement.namespace <> sysAdlStringConstantsEmpty)
        
        If ElementQualificationCompleted Then
        
            ElementServiceCache.AddElement NewElement
                
            'trigger element changed event
            MyShapeController.TriggerEvent ElementCreated
            
        End If
        
    End If
    
        
End Sub

Private Function KeyCanBeUsed(ByVal OpeningElement As SysAdlElement) As Boolean

    Dim Result As Boolean

    If Not MyShapeController.SysAdlElement.IsSameSysAdlType(OpeningElement) Then
        Result = False
    Else
        Result = True
    End If
    
    KeyCanBeUsed = Result
    
End Function

Private Sub ShowKeyIssueToUser()

    Dim elementIssue As New DiagramShapeIssues
    Dim elementIssueList As New UtilSysAdlList
    Dim elementIssueItem As New UtilSysAdlItem

            
    Set elementIssue = New DiagramShapeIssues
    Set elementIssueList = New UtilSysAdlList
            
    elementIssue.Init MyShapeController.GetShapeUniqueId, MyShapeController.SysAdlElement.sysadlType
    elementIssue.AddIssue FactoryMessageText.GetMessageText(sysAdlMessageKeyAlreadyInUse)
    
    elementIssueItem.Init elementIssue.ShapeId, elementIssue
            
    elementIssueList.Add elementIssueItem
            
    DiagramMessagePublisher.ShowIssues Visio.ActiveDocument, elementIssueList

    'show message to user to user
    GUIServices.ShowWarnMessage sysAdlMessageKeyAlreadyInUse

End Sub


