VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim myMouseListener As MouseListener

Private Sub Document_BeforeDocumentClose(ByVal Doc As IVDocument)

    DocumentEventProcessor.ProcessDocumentClosing Doc
    
    UnloadRibbon
    
End Sub

Private Sub Document_BeforeSelectionDelete(ByVal Selection As IVSelection)

    DocumentEventProcessor.ProcessSelectedShapedRemoving Selection
    
End Sub

Private Sub Document_DocumentCreated(ByVal Doc As IVDocument)

    LoadRibbon
    
     Set myMouseListener = New MouseListener

End Sub

Private Sub Document_DocumentOpened(ByVal Doc As IVDocument)

    Dim ControllerList As Collection
    Dim CurrentController As ShapeController
    
    LoadRibbon
    
    Set myMouseListener = New MouseListener
    
    Set ControllerList = DocumentEventProcessor.ProcessDocumentOpening(Doc)

 
    For Each CurrentController In ControllerList
        
        FactoryShapeListener.CreateShapeListener CurrentController
            
    Next
    
End Sub

Private Sub Document_DocumentSaved(ByVal Doc As IVDocument)
    
    ProcessDocumentSaving Doc, False 'document only saved... not saved as

End Sub

Private Sub Document_DocumentSavedAs(ByVal Doc As IVDocument)
        
    ProcessDocumentSaving Doc, True 'document saved as...
    
End Sub

Private Sub ProcessDocumentSaving(ByVal Doc As IVDocument, ByVal IsSaveAs As Boolean)
        
    DocumentEventProcessor.ProcessDocumentSaving Doc, IsSaveAs

End Sub

Private Sub Document_ShapeAdded(ByVal Shape As IVShape)
    
    Dim NewShapeController As ShapeController
    
    Set NewShapeController = DocumentEventProcessor.ProcessShapeAdding(Shape)
    
    FactoryShapeListener.CreateShapeListener NewShapeController
    
End Sub


Private Sub LoadRibbon()
    FactoryRibbon.CustomUIStart ThisDocument
End Sub

Private Sub UnloadRibbon()
    FactoryRibbon.CustomUIStop ThisDocument
End Sub


