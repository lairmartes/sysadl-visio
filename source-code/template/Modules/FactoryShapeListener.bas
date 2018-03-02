Attribute VB_Name = "FactoryShapeListener"
Option Explicit
    Private ShapeListenerList As New Collection


Public Sub CreateShapeListener(ByVal SomeShapeController As ShapeController)
    
    Dim CreateShapeListener As ShapeListener
    
    Set CreateShapeListener = New ShapeListener
    
    CreateShapeListener.Init SomeShapeController
    
    ShapeListenerList.Add CreateShapeListener
    
End Sub

