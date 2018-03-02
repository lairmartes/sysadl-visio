Attribute VB_Name = "ElementServiceCache"
Option Explicit

    Private CacheElements As New Collection
    
Public Sub AddElement(ByVal element As SysAdlElement)

    Dim CheckElementAdded As SysAdlElement
    
    Set CheckElementAdded = FindCachedElement(element)
        
    If (CheckElementAdded Is Nothing) Then
        
        CacheElements.Add element
        
    End If

End Sub


Public Function OpenElement(ByVal element As SysAdlElement) As SysAdlElement

    Dim Result As SysAdlElement
    
    Set Result = FindCachedElement(element)
    
    If Result Is Nothing Then
    
        Set Result = ElementServicePersistence.OpenPersistedElement(element)
        
        If Not (Result Is Nothing) Then
        
            CacheElements.Add Result
            
        End If
    
    End If
    
    Set OpenElement = Result
    

End Function


Public Sub ProcessElementRemove(ByVal element As SysAdlElement)

    ElementServicePersistence.ProcessElementDelete element
    
    If element.ShapeViewerList.Count = 0 Then
                    
       RemoveCachedElement element
        
    End If
    
End Sub
     
Private Function FindCachedElement(ByVal element As SysAdlElement, Optional ByVal elementType As String) As SysAdlElement

    Dim Result As SysAdlElement
    Dim CurrentElement As SysAdlElement
    
    Set Result = Nothing
    
    For Each CurrentElement In CacheElements
    
        If Not (CurrentElement Is Nothing) Then
    
            If CurrentElement.Equals(element) Then
            
                If elementType = sysAdlStringConstantsEmpty Then
            
                    Set Result = CurrentElement
                    
                Else
                    
                    If CurrentElement.sysadlType = elementType Then
                    
                        Set Result = CurrentElement
                        
                    End If
                
                End If
                
                Exit For
                
            End If
            
        End If
    
    Next
    
    Set FindCachedElement = Result

End Function

Public Function IsElementExists(ByVal element As SysAdlElement) As Boolean

    Dim Result As Boolean
    Dim foundInCache As Boolean

    foundInCache = Not (FindCachedElement(element) Is Nothing)
    
    If Not foundInCache Then
    
        Result = ElementServicePersistence.IsElementExists(element)
        
    Else
        
        Result = foundInCache
        
    End If
    
    IsElementExists = Result

End Function

Public Function IsElementExistsWithType(ByVal element As SysAdlElement) As Boolean

    Dim Result As Boolean
    Dim foundInCache As Boolean

    foundInCache = Not (FindCachedElement(element, element.sysadlType) Is Nothing)
    
    If Not foundInCache Then
    
        Result = ElementServicePersistence.IsElementExistsWithType(element)
        
    Else
        
        Result = foundInCache
        
    End If
    
    IsElementExistsWithType = Result

End Function


Private Sub RemoveCachedElement(ByVal element As SysAdlElement)

    Dim CurrentElement As SysAdlElement
    Dim I As Integer
    Dim CachedElementsCount As Integer
    
    CachedElementsCount = CacheElements.Count
    
    For I = 1 To CachedElementsCount
    
        Set CurrentElement = CacheElements.Item(I)
    
        If Not (CurrentElement Is Nothing) Then
    
            If CurrentElement.Equals(element) Then
            
                CacheElements.Remove I
                
                Exit For
                
            End If
            
        End If
    
    Next

End Sub

Public Sub ProcessElementPersistence(ByVal Doc As IVDocument)

    Dim CurrentElement As SysAdlElement
    Dim ElementPersistedEvent As EventSysAdl
    Dim ElementsToBePersisted As Collection
    Dim ElementMustBePersisted As Boolean
    
    Set ElementPersistedEvent = New EventSysAdl
    
    Set ElementsToBePersisted = GetElementsFromDoc(Doc)
    
    ElementPersistedEvent.Init sysAdlEventElementPersisted
    
    For Each CurrentElement In ElementsToBePersisted
    
        ElementMustBePersisted = Not GUIServices.IsRelationElementType(CurrentElement.sysadlType)

        If CurrentElement.IsDirty And ElementMustBePersisted Then
        
            ElementServicePersistence.ProcessElementSaving CurrentElement
            
            CurrentElement.HandleEvent ElementPersistedEvent
        
        End If
    
    Next
    
End Sub

Private Function GetElementsFromDoc(ByVal Doc As IVDocument) As Collection

    Dim Result As Collection
    Dim PageList As Visio.Pages
    Dim CurrentPage As Visio.page
    Dim ShapeList As Visio.Shapes
    Dim CurrentShape As Visio.shape
    Dim ShapeControllerFound As shapeController
    
    Set Result = New Collection

    Set PageList = Doc.Pages
    
    For Each CurrentPage In PageList
    
        Set ShapeList = CurrentPage.Shapes
        
        For Each CurrentShape In ShapeList
        
            Set ShapeControllerFound = FactoryShapeController.GetShapeControllerByShape(CurrentShape)

            Result.Add ShapeControllerFound.SysAdlElement
            
        Next
    
    Next
    
    Set GetElementsFromDoc = Result

End Function

