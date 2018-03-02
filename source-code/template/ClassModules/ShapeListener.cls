VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Dim WithEvents MyShape As Shape
Attribute MyShape.VB_VarHelpID = -1
    Dim MyShapeController As ShapeController
    
Public Sub Init(ByVal SomeShapeController As ShapeController)
            
    Set MyShape = SomeShapeController.Shape
    Set MyShapeController = SomeShapeController

End Sub

Public Sub MyShape_CellChanged(ByVal Cell As IVCell)

    MyShapeController.FireCellChanged Cell

End Sub





