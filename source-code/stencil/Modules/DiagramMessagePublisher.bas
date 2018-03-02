Attribute VB_Name = "DiagramMessagePublisher"
Option Explicit

Public Sub ShowIssues(ByVal aVisioDoc As IVDocument, ByVal issueList As UtilSysAdlList)

    Dim PageList As Visio.Pages
    Dim CurrentPage As Visio.page
    Dim ShapeList As Visio.Shapes
    Dim CurrentShape As Visio.shape
    Dim IssuesFound As DiagramShapeIssues
    Dim Message As String

    Set PageList = aVisioDoc.Pages
    
    For Each CurrentPage In PageList
    
        Set ShapeList = CurrentPage.Shapes
        
        For Each CurrentShape In ShapeList
        
            Set IssuesFound = issueList.Item(CurrentShape.UniqueID(visGetGUID))
            
            If Not IssuesFound Is Nothing Then
            
                Message = CreateMessage(IssuesFound)
        
                AddIssue Message, IssuesFound.sysadlType, CurrentPage, CurrentShape
                
            End If
            
        Next
    
    Next

End Sub

Private Sub AddIssue(ByVal Message As String, ByVal sysadlType As String, ByVal page As Visio.page, ByVal shape As Visio.shape)

    Dim PositionX As String
    Dim PositionY As String
    Dim CurrentLineProperty As Integer
    Dim PageSheet As Visio.shape
    Dim IsConnector As Boolean
    
    Set PageSheet = page.PageSheet
    
    IsConnector = False
    
    If (sysadlType = sysAdlTypeSetRepresents _
          Or sysadlType = sysAdlTypeSetChannel _
          Or sysadlType = sysAdlTypeSetComposedBy _
          Or sysadlType = sysAdlTypeSetIsA _
          Or sysadlType = sysAdlTypeSetConnector _
          Or sysadlType = sysAdlTypeSetTransition _
          Or sysadlType = sysAdlTypeSetDependsOn) Then
          
           IsConnector = True
           
    End If
    
    
    If Not IsConnector Then
    
        PositionX = shape.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).ResultStr(visMillimeters)
        PositionY = shape.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).ResultStr(visMillimeters)
     
    Else
    
        PositionX = shape.CellsSRC(visSectionObject, visRowXForm1D, vis1DEndX).ResultStr(visMillimeters)
        PositionY = shape.CellsSRC(visSectionObject, visRowXForm1D, vis1DEndY).ResultStr(visMillimeters)
    
    
    End If
    
    'include the type of the element or relation to help in message visualization
    If IsConnector Then
        Message = "Issues found in this " + sysadlType + " relation:" + vbCrLf + Message
    Else
        Message = "Issues found in this " + sysadlType + " element:" + vbCrLf + Message
    End If
    
   CurrentLineProperty = PageSheet.AddRow(visSectionAnnotation, visRowLast, visTagDefault)
    
    PageSheet.CellsSRC(visSectionAnnotation, CurrentLineProperty, visAnnotationX).FormulaU = """" + PositionX + """"
    PageSheet.CellsSRC(visSectionAnnotation, CurrentLineProperty, visAnnotationY).FormulaU = """" + PositionY + """"
    PageSheet.CellsSRC(visSectionAnnotation, CurrentLineProperty, visAnnotationReviewerID).FormulaU = """" + "SysADL" + """"
    PageSheet.CellsSRC(visSectionAnnotation, CurrentLineProperty, visAnnotationMarkerIndex).FormulaU = CurrentLineProperty + 1
    PageSheet.CellsSRC(visSectionAnnotation, CurrentLineProperty, visAnnotationComment).FormulaU = """" + Message + """"
    
End Sub

Public Sub RemoveIssues(ByVal aVisioDoc As IVDocument)

    Dim PageList As Visio.Pages
    Dim CurrentPage As Visio.page
    Dim ShapeList As Visio.Shapes
    Dim CurrentShape As Visio.shape
    Dim ShapeControllerFound As shapeController

    Set PageList = aVisioDoc.Pages
    
    For Each CurrentPage In PageList
    
        Set ShapeList = CurrentPage.Shapes
        
        For Each CurrentShape In ShapeList
        
            RemoveMessages CurrentPage
            
        Next
    
    Next

End Sub

Private Sub RemoveMessages(ByVal page As Visio.page)

    Dim PageSheet As Visio.shape

    Set PageSheet = page.PageSheet
    
    PageSheet.DeleteSection visSectionAnnotation

    PageSheet.AddSection visSectionAnnotation

End Sub

Private Function CreateMessage(ByVal IssuesFound As DiagramShapeIssues) As String

    Dim Result As String
    Dim Message As Variant
    Dim AlreadyPassed As Boolean
    Dim MessageList As Collection
    
    AlreadyPassed = False
    
    Set MessageList = IssuesFound.Issues
    
    For Each Message In MessageList
    
        If AlreadyPassed = False Then
        
            AlreadyPassed = True
            
        Else
        
            Result = Result + vbCrLf
            
        End If
        
        Result = Result + Message
    
    Next
    
    CreateMessage = Result

End Function

