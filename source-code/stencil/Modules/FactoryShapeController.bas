Attribute VB_Name = "FactoryShapeController"
Option Explicit
    Private ShapeControllerList As New Collection


Public Function CreateShapeController(ByVal SomeShape As IVShape, ByVal SomeSysAdlElement As SysAdlElement) As shapeController
    
    Set CreateShapeController = New shapeController
    
    CreateShapeController.Init SomeShape, SomeSysAdlElement
    
    ShapeControllerList.Add CreateShapeController  ', CreateShapeController.GetShapeUniqueId
    
End Function

Public Function FireShapeAdded(ByVal NewShape As IVShape) As shapeController

    Dim NewElement As SysAdlElement
    Dim ShapeControllerToAdd As shapeController
    Dim IsOpenShapeCommand As Boolean
    Dim SysADLTypeFromShape As String
    Dim IsElementCached As Boolean
    
    Dim CreateNewShape As New CommandCreateNewShapeSysAdl
    Dim OpenShape As New CommandOpenShapeSysAdl

    Dim QueueCommand As New QueueCommandExecution
    
    Dim NewShapeId As String
    
    NewShapeId = NewShape.UniqueID(visGetGUID)
    
    SysADLTypeFromShape = GUIServices.PrepareMasterNameForElement(NewShape.Master.Name)
    
        Set NewElement = DiagramServiceCache.GetDiagramElementByShapeId(NewShapeId, SysADLTypeFromShape)
        'Set NewElement = ElementServicePersistence.GetElementBasicDataByShapeId(NewShapeId, SysADLTypeFromShape)
        
        If NewElement Is Nothing Then
    
            Set NewElement = New SysAdlElement
            NewElement.Init SysADLTypeFromShape, sysAdlNoStereotype
    
        Else
            'recover element already cached
            IsElementCached = ElementServiceCache.IsElementExists(NewElement)
            
            If IsElementCached Then
            
                Set NewElement = ElementServiceCache.OpenElement(NewElement)
                
            End If
            
            'indicate that this is an recover data operation
            IsOpenShapeCommand = True
            
        End If
        
        Set FireShapeAdded = CreateShapeController(NewShape, NewElement)
        
        If Not IsOpenShapeCommand Then
            CreateNewShape.Init FireShapeAdded
            QueueCommand.AddCommand CreateNewShape
        Else
            OpenShape.Init FireShapeAdded
            QueueCommand.AddCommand OpenShape
        End If
        
        QueueCommand.ExecuteCommandList
        

End Function


Public Sub RemoveShapeController(ByVal ShapeClosed As IVShape, ByVal IsKeepInDatabase As Boolean)
    
    Dim CurrentShapeController As shapeController
    Dim ControllerQtty As Integer
    Dim CurrentIndex As Integer
    Dim CurrentShapeIdClosed As String
    Dim CurrentControllerShapeId As String
    Dim RemoveCommand As CommandRemoveShape
    
    ControllerQtty = ShapeControllerList.Count
    
    CurrentShapeIdClosed = ShapeClosed.UniqueID(visGetGUID)
    
    For CurrentIndex = 1 To ControllerQtty
            Set CurrentShapeController = ShapeControllerList.Item(CurrentIndex)
            
            If Not CurrentShapeController.MarkedAsRemoved Then
            
                CurrentControllerShapeId = CurrentShapeController.GetShapeUniqueId
            
                If CurrentShapeIdClosed = CurrentControllerShapeId Then
                
                    CurrentShapeController.ProcessControllerRemove
                
                    If Not IsKeepInDatabase Then
                        
                        Set RemoveCommand = New CommandRemoveShape
                        Set CurrentShapeController = ShapeControllerList.Item(CurrentIndex)
                        RemoveCommand.Init CurrentShapeController
                        RemoveCommand.ICommand_Execute
                        
                    End If
                    
                End If
            End If
    Next
            
End Sub


Public Function GetShapeControllerByShape(ByVal AShape As IVShape) As shapeController
    
    Dim ShapeId As String
    Dim CurrentShapeController As shapeController
    Dim CurrentShapeControllerID As String
    Dim ControllerQtty As Integer
    Dim CurrentIndex As Integer
    Dim Result As shapeController
    Dim ShapeControllerFound As Boolean
    
    ControllerQtty = ShapeControllerList.Count
    
    ShapeId = AShape.UniqueID(visGetGUID)
    
    For CurrentIndex = 1 To ControllerQtty
            
        Set CurrentShapeController = ShapeControllerList.Item(CurrentIndex)
            
        If Not CurrentShapeController.MarkedAsRemoved Then
            
            CurrentShapeControllerID = CurrentShapeController.GetShapeUniqueId
            
            If ShapeId = CurrentShapeControllerID Then
                
                Set Result = CurrentShapeController
                    
                ShapeControllerFound = True
                    
                Exit For
                
            End If
        End If
    Next
    
    If Not ShapeControllerFound Then
        
        Set Result = Nothing
    
    End If
    
    Set GetShapeControllerByShape = Result
            
End Function

Public Function GetShapeControllerListByElementType(ByVal anElementType As String) As Collection

    Dim Result As New Collection
    Dim CurrentShapeController As shapeController
    
    For Each CurrentShapeController In ShapeControllerList
    
        If Not CurrentShapeController.MarkedAsRemoved Then
    
            If CurrentShapeController.ShapeSysAdlType = anElementType Then
                
                Result.Add CurrentShapeController
            
            End If
        
        End If
    
    Next

    Set GetShapeControllerListByElementType = Result
    
End Function


