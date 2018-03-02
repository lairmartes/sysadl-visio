VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandRemoveShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    
Implements ICommand
Private MyShapeController As shapeController

Public Sub ICommand_Execute()

    Dim aSysAdlElement As SysAdlElement
    
    Set aSysAdlElement = MyShapeController.SysAdlElement
    
    MyShapeController.ProcessControllerRemove
    
    ElementServiceCache.ProcessElementRemove aSysAdlElement

End Sub

Public Sub Init(ByVal ShapeControllerToAdd As shapeController)
    
    
    Set MyShapeController = ShapeControllerToAdd
    
End Sub
