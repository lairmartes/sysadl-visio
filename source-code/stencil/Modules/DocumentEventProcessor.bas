Attribute VB_Name = "DocumentEventProcessor"
Option Explicit

Private Const REMOVE_FROM_DATABASE = False
Private Const DO_NOT_REMOVE_FROM_DATABASE = True


Public Sub ProcessDocumentSaving(ByVal Doc As IVDocument, ByVal IsSaveAs As Boolean)

    Dim ProblemsWithFields As Collection
    Dim ProblemsWithDiagram As Collection
    Dim analysisResult As DiagramAnalysisResult
    Dim DiagramIsValid As Boolean
    Dim DiagramHasObjective As Boolean
    Dim DiagramIssues As UtilSysAdlList
    Dim TotalSteps As Integer
    Dim diagramFileName As String
    
    ' if "Save As..." has been actionated all Shapes' IDs must be replaced
    If IsSaveAs Then ReplaceAllShapeUniqueIDs Doc
    
    TotalSteps = 4
    
    'initializing progress bar
    GUIServices.ProgressBarViewInit "Persisting Sys-ADL Elements"
    GUIServices.ProgressBarUpdate "Validating fields", TotalSteps, 1 'first step
    
    DiagramMessagePublisher.RemoveIssues Doc
        
    Set DiagramIssues = DiagramFieldsValidator.ValidateDiagramElementsFields(Doc)
    
    DiagramIsValid = DiagramIssues.Count < 1

    If DiagramIsValid Then
    
        GUIServices.ProgressBarUpdate "Validating diagram", TotalSteps, 2 'second step
    
        Set analysisResult = DiagramValidator.ValidateDiagram(Doc)
        
        DiagramIsValid = analysisResult.DiagramIsOk
        
        If DiagramIsValid Then
        
            GUIServices.ProgressBarUpdate "Saving Sys-ADL Elements", TotalSteps, 3 'third step
            
            'save elements data
            ElementServiceCache.ProcessElementPersistence Doc


            GUIServices.ProgressBarUpdate "Creating analysis result file", TotalSteps, 4 'fourth step
            ' publish result analysis
            diagramFileName = Doc.Path + Doc.Name
            
             DiagramServiceCache.ProcessDiagramSaving diagramFileName, analysisResult
            
        Else
        
            GUIServices.ProgressBarViewFinish
            
            DiagramHasObjective = analysisResult.HasObjectives
            
            If (DiagramHasObjective) Then
            
                DiagramMessagePublisher.ShowIssues Doc, analysisResult.Errors
        
                GUIServices.ShowWarnMessage sysAdlMessageInvalidDiagram
                
            Else
            
                GUIServices.ShowWarnMessage sysAdlMessageWithoutObjective
            
            End If
            
        End If
        
        GUIServices.ProgressBarViewFinish
        
    Else
    
        DiagramMessagePublisher.ShowIssues Doc, DiagramIssues
    
        GUIServices.ProgressBarViewFinish
        
        GUIServices.ShowWarnMessage sysAdlMessageProblemsSaving
        
    End If

End Sub

Public Function ProcessDocumentOpening(ByVal Doc As IVDocument) As Collection

    Dim PageList As Visio.Pages
    Dim ShapeList As Visio.Shapes
    Dim CurrentPage As Visio.page
    Dim CurrentShape As Visio.shape
    Dim CurrentStep As Integer
    Dim CurrentTaskDescription As String
    Dim TotalSteps As Integer
    Dim Result As New Collection
    
    DiagramServiceCache.ProcessDiagramOpening Doc.Path + Doc.Name
    
    Set PageList = Doc.Pages
 
    For Each CurrentPage In PageList
 
        Set ShapeList = CurrentPage.Shapes
        
        TotalSteps = ShapeList.Count
        
        If TotalSteps > 0 Then
            GUIServices.ProgressBarViewInit "Opening SysADL Diagram"
        End If
 
        For Each CurrentShape In ShapeList
        
            Dim NewShapeController As shapeController
 
            Set NewShapeController = FactoryShapeController.FireShapeAdded(CurrentShape)
            
            Result.Add NewShapeController
 
            CurrentStep = CurrentStep + 1
            
            CurrentTaskDescription = NewShapeController.SysAdlElement.Key + " opened"
            
            GUIServices.ProgressBarUpdate CurrentTaskDescription, TotalSteps, CurrentStep
            
        Next
        
        GUIServices.ProgressBarViewFinish
    Next
    
    Set ProcessDocumentOpening = Result

End Function


Public Sub ProcessDocumentClosing(ByVal Doc As IVDocument)

    Dim PageList As Visio.Pages
    Dim ShapeList As Visio.Shapes
    Dim CurrentPage As Visio.page
    Dim CurrentShape As Visio.shape
    
    Set PageList = Doc.Pages
 
    For Each CurrentPage In PageList
 
        Set ShapeList = CurrentPage.Shapes
 
        For Each CurrentShape In ShapeList
        
            FactoryShapeController.RemoveShapeController CurrentShape, DO_NOT_REMOVE_FROM_DATABASE
        
        Next
    Next

End Sub

Public Sub ProcessSelectedShapedRemoving(ByVal Selection As IVSelection)

    Dim CurrentShape As IVShape
    
    For Each CurrentShape In Selection
                                
        FactoryShapeController.RemoveShapeController CurrentShape, REMOVE_FROM_DATABASE
    
    Next
            
End Sub

Public Function ProcessShapeAdding(ByVal NewShape As IVShape) As shapeController

    Dim Result As shapeController
    
    Set Result = FactoryShapeController.FireShapeAdded(NewShape)

    Set ProcessShapeAdding = Result
    
End Function

Private Sub ReplaceAllShapeUniqueIDs(ByVal Doc As IVDocument)

    Dim PageList As Visio.Pages
    Dim ShapeList As Visio.Shapes
    Dim CurrentPage As Visio.page
    Dim CurrentShape As Visio.shape
    Dim CurrentShapeId As String
    Dim CurrentShapeController As shapeController
    
    Set PageList = Doc.Pages
 
    For Each CurrentPage In PageList
 
        Set ShapeList = CurrentPage.Shapes
 
        For Each CurrentShape In ShapeList
        
            CurrentShapeId = CurrentShape.UniqueID(visDeleteGUID)
            CurrentShapeId = CurrentShape.UniqueID(visGetOrMakeGUID)
            
            Set CurrentShapeController = FactoryShapeController.GetShapeControllerByShape(CurrentShape)
            
            'inform that element has changed
            CurrentShapeController.SysAdlElement.HandleEvent FactoryEvent.CreateEvent(sysAdlEventDocumentSavedAs)
            
        Next

    Next


End Sub
