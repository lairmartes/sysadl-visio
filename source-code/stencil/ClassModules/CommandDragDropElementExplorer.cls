VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandDragDropElementExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICommand
Implements IVisioEventLister

Private shapeController As shapeController
Private elementId As String
Private elementNamespace As String
Private isStencilAlreadyOpened As Boolean
Private visioStencil As Visio.Document
Private isElementDragging As Boolean

    


Public Sub ICommand_Execute()


    If (Not GetShapeAdded Is Nothing) Then

        Dim changeKeyCommand As CommandChangeKey
    
        Set changeKeyCommand = New CommandChangeKey
        
        changeKeyCommand.Init shapeController, elementNamespace, elementId
        
        changeKeyCommand.ICommand_Execute
        
        isElementDragging = False
    
    End If

End Sub

Public Sub Init(ByVal ElementKey As String)

    Dim selectedElement As SysAdlElement
    Dim elementType As String
    Dim visioMaster As Visio.Master
    Dim shapeAdded As Visio.shape
    
    elementId = GUIServices.GetIdFromString(ElementKey)
    elementNamespace = GUIServices.GetNamespaceFromString(ElementKey)
    
    Set selectedElement = New SysAdlElement
    
    selectedElement.InitBaseData elementNamespace, elementId, sysAdlStringConstantsEmpty
    
    Set selectedElement = ElementServiceCache.OpenElement(selectedElement)
    
    If Not selectedElement Is Nothing Then
    
        elementType = selectedElement.sysadlType
        
        isElementDragging = True
        
        Set shapeAdded = CreateSysADLShape(elementType)
        
        Set shapeController = FactoryShapeController.CreateShapeController(shapeAdded, selectedElement)
        
    Else
    
        Set shapeAdded = Nothing
        
    End If
    

End Sub

Private Function CreateSysADLShape(ByVal elementType As String) As Visio.shape


Dim visioStencil As Visio.Document
    Dim visioDocs As Visio.Documents

    Dim currentVisioPage As Visio.page
    Dim Result As Visio.shape
    
    Set visioDocs = ThisDocument.Application.Documents
   
   Set visioStencil = GUIServices.GetSysADLVisioStencil

    Set currentVisioPage = ThisDocument.Application.ActiveDocument.Pages(1)


    Dim visioSysADLMaster As Visio.Master
    Set visioSysADLMaster = visioStencil.Masters(elementType)
    
    
    Set Result = currentVisioPage.Drop(visioSysADLMaster, 5, 5)
    
    Set CreateSysADLShape = Result

End Function

Public Function GetShapeAdded() As Visio.shape

    Dim Result As Visio.shape
    
    If Not shapeController Is Nothing Then
    
        Set Result = shapeController.shape
        
    End If
    
    Set GetShapeAdded = Result
    
End Function


Public Sub IVisioEventLister_MouseDown(ByVal mouseEvent As VisioEventData)


    Debug.Print "Mouse downed..."

End Sub

Public Sub IVisioEventLister_MouseUp(ByVal mouseEvent As VisioEventData)


    Debug.Print "Mouse upped..."

End Sub

Public Sub IVisioEventLister_MouseMove(ByVal mouseEvent As VisioEventData)

    If isElementDragging Then
    
        Dim isMovingAChannel As Boolean
        Dim isMovingATransition As Boolean
    
        isMovingAChannel = (shapeController.ShapeSysAdlType = sysAdlTypeSetChannel)
        isMovingATransition = (shapeController.ShapeSysAdlType = sysAdlTypeSetTransition)
                              
        If isMovingAChannel Then
        
            GetShapeAdded.Cells("BeginY") = mouseEvent.MouseY
            GetShapeAdded.Cells("BeginX") = mouseEvent.MouseX
            GetShapeAdded.Cells("EndY") = mouseEvent.MouseY - (0.8)
            GetShapeAdded.Cells("EndX") = mouseEvent.MouseX + (0.8)
            
        ElseIf isMovingATransition Then
        
            GetShapeAdded.Cells("BeginY") = mouseEvent.MouseY
            GetShapeAdded.Cells("BeginX") = mouseEvent.MouseX
            GetShapeAdded.Cells("EndY") = mouseEvent.MouseY
            GetShapeAdded.Cells("EndX") = mouseEvent.MouseX + (1.5)
              
        Else

            GetShapeAdded.Cells("piny") = mouseEvent.MouseY
            GetShapeAdded.Cells("pinx") = mouseEvent.MouseX
            
        End If
        
    End If

End Sub

