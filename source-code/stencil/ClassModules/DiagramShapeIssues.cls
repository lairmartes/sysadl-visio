VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiagramShapeIssues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pShapeId As String
    Private pSysAdlType As String
    Private pIssueList As Collection
    
    Public Sub Init(ByVal aShapeId As String, ByVal aSysAdlType)
    
        pShapeId = aShapeId
        pSysAdlType = aSysAdlType
        
        Set pIssueList = New Collection
    
    End Sub

    Public Sub AddIssue(ByVal aMessage As String)
    
        pIssueList.Add aMessage
    
    End Sub

    Public Property Get ShapeId() As String
    
        ShapeId = pShapeId
    
    End Property
    
    Public Property Get sysadlType() As String
    
        sysadlType = pSysAdlType

    End Property

    Public Property Get Issues() As Collection
    
        Set Issues = pIssueList
    
    End Property
