VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiagramElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Dim pElementNamespace As String
    Dim pElementId As String
    Dim pElementType As String
    Dim pElementStereotype As String
    Dim pElementURLInfo As String
    Dim pElementShapeId As String
    
    
Public Sub Init(ByVal elementNamespace As String, _
                 ByVal elementId As String, _
                 ByVal elementType As String, _
                 ByVal ElementStereotype As String, _
                 ByVal ElementURLInfo As String, _
                 ByVal ElementShapeId As String)

    pElementNamespace = elementNamespace
    pElementId = elementId
    pElementType = elementType
    pElementStereotype = ElementStereotype
    pElementURLInfo = ElementURLInfo
    pElementShapeId = ElementShapeId


End Sub
                 
Public Property Get elementNamespace()

    elementNamespace = pElementNamespace

End Property

Public Property Get elementId()

    elementId = pElementId

End Property

Public Property Get ElementURLInfo()

    ElementURLInfo = pElementURLInfo

End Property

Public Property Get ElementStereotype()

    ElementStereotype = pElementStereotype

End Property

Public Property Get elementType()

    elementType = pElementType

End Property



Public Function IsSameShapeId(ByVal elementShapeIdParameter As String) As Boolean

    Dim Result As Boolean
    
    Result = (pElementShapeId = elementShapeIdParameter)
    
    IsSameShapeId = Result
 
End Function

