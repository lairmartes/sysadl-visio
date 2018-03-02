Attribute VB_Name = "DiagramServicePersistence"


Public Sub ProcessDiagramSaving(ByVal fileName As String, ByVal analysisResult As DiagramAnalysisResult)

    DiagramServicePersistenceXMLOut.ProcessDiagramSaving fileName, analysisResult

End Sub

Public Sub ProcessOpenDiagram(ByVal fileName As String)

    DiagramServicePersistenceXMLIn.ProcessOpenDiagram fileName

End Sub

Public Function GetDiagramElementByShapeId(ByVal ElementShapeId As String, ByVal elementType As String) As SysAdlElement

    Dim Result As SysAdlElement

    Set Result = DiagramServicePersistenceXMLIn.GetDiagramElementByShapeId(ElementShapeId, elementType)

    Set GetDiagramElementByShapeId = Result

End Function
