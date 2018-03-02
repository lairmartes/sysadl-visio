Attribute VB_Name = "DiagramServiceCache"
Public Sub ProcessDiagramSaving(ByVal fileName As String, ByVal analysisResult As DiagramAnalysisResult)

    DiagramServicePersistence.ProcessDiagramSaving fileName, analysisResult

End Sub


Public Sub ProcessDiagramOpening(ByVal fileName As String)

    DiagramServicePersistence.ProcessOpenDiagram fileName

End Sub

Public Function GetDiagramElementByShapeId(ByVal ElementShapeId As String, ByVal elementType As String) As SysAdlElement

    Dim Result As SysAdlElement

    Set Result = DiagramServicePersistence.GetDiagramElementByShapeId(ElementShapeId, elementType)
    
    Set GetDiagramElementByShapeId = Result

End Function
