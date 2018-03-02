VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandOpenShapeSysAdl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Implements ICommand
Private MyShapeController As shapeController

Public Sub ICommand_Execute()

    Dim DocumentOpenedEvent As EventSysAdl
    
    Set DocumentOpenedEvent = FactoryEvent.CreateEvent(sysAdlEventDocumentOpened)
    
    ElementServiceCache.AddElement MyShapeController.SysAdlElement
    
    MyShapeController.TriggerEvent DocumentOpenedEvent


End Sub

Public Sub Init(ByVal ShapeControllerToAdd As shapeController)
    
    Set MyShapeController = ShapeControllerToAdd
    
End Sub

