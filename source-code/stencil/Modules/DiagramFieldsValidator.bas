Attribute VB_Name = "DiagramFieldsValidator"
Option Explicit


    Private Const NAMESPACE_MIN_LENGTH = 1
    Private Const NAMESPACE_MAX_LENGTH = 200
    Private Const ID_MIN_LENGTH = 1
    Private Const ID_MAX_LENGTH = 50
    Private issueList As UtilSysAdlList

Public Function ValidateDiagramElementsFields(ByVal aDocument As Visio.IVDocument) As UtilSysAdlList
    
        Dim ShapeControllersInDiagram As Collection
        Dim CurrentShapeController As shapeController
        Dim IssuesFound As DiagramShapeIssues
        Dim IssueItem As UtilSysAdlItem
        Dim Result As UtilSysAdlList
        
        Set Result = New UtilSysAdlList
        
        Set ShapeControllersInDiagram = GetShapeControllersFromDoc(aDocument)
        
        GUIServices.LogMessage "Validating fiedls from " + aDocument.Name + "..."
        
        For Each CurrentShapeController In ShapeControllersInDiagram
        
            GUIServices.LogMessage "Validation fields from element " + CurrentShapeController.SysAdlElement.Key
        
            Set IssuesFound = ValidateFields(CurrentShapeController)
            
            If Not IssuesFound Is Nothing Then
            
                Set IssueItem = New UtilSysAdlItem
                
                IssueItem.Init IssuesFound.ShapeId, IssuesFound
            
                Result.Add IssueItem
            
            End If
        
        Next

    Set ValidateDiagramElementsFields = Result

    End Function

Private Function ValidateFields(ByVal aShapeController As shapeController) As DiagramShapeIssues

    Dim element As SysAdlElement
    Dim errorMessages As Collection
    Dim currentErrorMessage As Variant
    Dim Result As DiagramShapeIssues

    Dim EventChangeNamespace As EventSysAdl
    Dim NamespaceExtractedFromDoc As String
    
    NamespaceExtractedFromDoc = GUIServices.GetDocumentNamespace()
        
    Set EventChangeNamespace = FactoryEvent.CreateEvent(sysAdlEventChangedCellValue)
    
    Set Result = Nothing
    
    Set element = aShapeController.SysAdlElement
    
    If element.namespace = sysAdlStringConstantsEmpty Then
        
        element.ChangeFieldValue sysAdlKeyCustPropRowNameNamespace, NamespaceExtractedFromDoc
        element.HandleEvent EventChangeNamespace

    End If
    
    
    Set errorMessages = element.ValidateFields
    
    If (errorMessages.Count > 0) Then
    
        Set Result = New DiagramShapeIssues
        
        Result.Init aShapeController.GetShapeUniqueId, aShapeController.ShapeSysAdlType
        
        For Each currentErrorMessage In errorMessages
        
            Result.AddIssue currentErrorMessage
    
        Next
    
    End If
    
    Set ValidateFields = Result

End Function

Private Function GetShapeControllersFromDoc(ByVal aDocument As Visio.IVDocument) As Collection

    Dim Result As Collection
    Dim PageList As Visio.Pages
    Dim CurrentPage As Visio.page
    Dim ShapeList As Visio.Shapes
    Dim CurrentShape As Visio.shape
    Dim ShapeControllerFound As shapeController
    
    Set Result = New Collection

    Set PageList = aDocument.Pages
    
    For Each CurrentPage In PageList
    
        Set ShapeList = CurrentPage.Shapes
        
        For Each CurrentShape In ShapeList
        
            Set ShapeControllerFound = FactoryShapeController.GetShapeControllerByShape(CurrentShape)
            
            Result.Add ShapeControllerFound
        
        Next
    
    Next
    
    Set GetShapeControllersFromDoc = Result

End Function

Public Function AreFieldsOk(ByVal aSysAdlElement As SysAdlElement) As Boolean

    Dim Result As Boolean
    Dim hasErrors As Boolean
    
    hasErrors = aSysAdlElement.ValidateFields.Count > 0
    
    Result = Not hasErrors
    
    AreFieldsOk = Result

End Function
