VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandCreateNewShapeSysAdl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICommand
Private MyShapeController As shapeController

Public Sub ICommand_Execute()

    Dim CreateEvent As EventSysAdl
    Dim ChangeStereotypeEvent As EventSysAdl
    Dim ChangeCellEvent As EventSysAdl
    Dim NamespaceFromDoc As String
    
    ElementServiceCache.AddElement MyShapeController.SysAdlElement
    
    Set CreateEvent = FactoryEvent.CreateEvent(sysAdlEventCreatedSysAdlElement)
    Set ChangeStereotypeEvent = FactoryEvent.CreateEvent(sysAdlEventChangedStereotype)
    
    MyShapeController.TriggerEvent CreateEvent
    MyShapeController.TriggerEvent ChangeStereotypeEvent
    
    NamespaceFromDoc = GUIServices.GetDocumentNamespace()
    
    If NamespaceFromDoc <> sysAdlStringConstantsEmpty Then
    
        Set ChangeCellEvent = FactoryEvent.CreateEvent(sysAdlEventChangedCellValue)
        MyShapeController.SysAdlElement.ChangeFieldValue sysAdlKeyCustPropRowNameNamespace, NamespaceFromDoc
        
        MyShapeController.TriggerEvent ChangeCellEvent
        
    End If

End Sub

Public Sub Init(ByVal ShapeControllerToAdd As shapeController)
    
    
    Set MyShapeController = ShapeControllerToAdd

    
End Sub

