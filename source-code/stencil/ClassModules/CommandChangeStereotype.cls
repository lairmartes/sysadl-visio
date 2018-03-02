VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandChangeStereotype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICommand
    
    Private MyShapeController As shapeController
    Private NewStereotypeValue As String
    
Public Sub Init(ByVal aShapeController As shapeController, ByVal ANewStereotypeValue As String)
    
    Set MyShapeController = aShapeController
    NewStereotypeValue = ANewStereotypeValue

End Sub

Public Sub ICommand_Execute()

    Dim CurrentElement As SysAdlElement
    Dim StereotypeChangedEvent As EventSysAdl
    Dim SysAdlElementChangedEvent As EventSysAdl
    
    MyShapeController.SysAdlElement.ChangeFieldValue sysAdlKeyCustPropRowNameStereotype, NewStereotypeValue
        
    Set StereotypeChangedEvent = FactoryEvent.CreateEvent(sysAdlEventChangedStereotype)
    MyShapeController.TriggerEvent StereotypeChangedEvent
    
End Sub
