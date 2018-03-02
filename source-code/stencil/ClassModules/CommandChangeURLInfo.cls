VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandChangeURLInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICommand
    Private MyShapeController As shapeController

    
Public Sub Init(ByVal aShapeController As shapeController)

    Set MyShapeController = aShapeController
    
End Sub

Public Sub ICommand_Execute()

    Dim ChangedURLInfo As EventSysAdl
    Dim SysAdlElementControlled As SysAdlElement
    
    Set SysAdlElementControlled = MyShapeController.SysAdlElement
    
    Set ChangedURLInfo = FactoryEvent.CreateEvent(sysAdlEventChangedURLInfo)
    
    MyShapeController.TriggerEvent ChangedURLInfo

End Sub


