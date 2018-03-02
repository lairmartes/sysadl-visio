VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiagramRelationSysAdl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pRelation As shapeController
    Private pRelationSource As shapeController
    Private pRelationDestiny As shapeController

    Public Sub Init(ByVal Relation As shapeController)

        Set pRelation = Relation
        
        CalculateRelationDirection

    End Sub
    
    Private Sub CalculateRelationDirection()
    
        Dim RelationConnect1 As Visio.Connect
        Dim RelationConnect2 As Visio.Connect
        Dim RelationConnects As Visio.Connects
        
        Dim ShapeFrom As Visio.shape
        Dim ShapeTo As Visio.shape
        
        Set RelationConnects = pRelation.shape.Connects
        
        If RelationConnects.Count = 2 Then
        
            Set RelationConnect1 = RelationConnects.Item(1)
            Set RelationConnect2 = RelationConnects.Item(2)
            
            Set ShapeFrom = RelationConnect1.ToSheet
            Set ShapeTo = RelationConnect2.ToSheet
    
            
            Set pRelationSource = FactoryShapeController.GetShapeControllerByShape(ShapeFrom)
            Set pRelationDestiny = FactoryShapeController.GetShapeControllerByShape(ShapeTo)
        
        End If
        

    End Sub
    
    Public Property Get Id() As String
    
        Dim Result As String
        
        Result = pRelation.GetShapeUniqueId
        
        Id = Result
        
    End Property
    
    Public Property Get RelationSource() As shapeController
    
        Dim Result As shapeController
        
        Set Result = pRelationSource
        
        Set RelationSource = Result
    
    End Property
    
    Public Property Get RelationDestiny() As shapeController
    
        Dim Result As shapeController
        
        Set Result = pRelationDestiny
        
        Set RelationDestiny = Result

    End Property
    
    Public Property Get RelationConnector() As shapeController
    
        Dim Result As shapeController
        
        Set Result = pRelation
        
        Set RelationConnector = pRelation
    
    End Property
    
    Public Property Get Relation() As String
    
        Dim Result As String
        
        Dim relationType As String
        Dim sourceType As String
        Dim destinyType As String
        
        Result = sysAdlStringConstantsEmpty
        
        relationType = pRelation.ShapeSysAdlType
        sourceType = pRelationSource.ShapeSysAdlType
        destinyType = pRelationDestiny.ShapeSysAdlType
        
        If (relationType = sysAdlTypeSetIsA) Then
        
            If (sourceType = destinyType) Then Result = sysAdlTypeSetIsA
            
        ElseIf (relationType = sysAdlTypeSetComposedBy) Then
                
            If (sourceType = sysAdlTypeSetSystem Or _
                sourceType = sysAdlTypeSetObjective Or _
                sourceType = sysAdlTypeSetNode Or _
                sourceType = sysAdlTypeSetRole) Then
                    
                    If (sourceType = destinyType) Then Result = sysAdlTypeSetComposedBy
                    
            End If
                
        ElseIf (relationType = sysAdlTypeSetDependsOn) Then
        
            If (sourceType = sysAdlTypeSetSystem Or _
                sourceType = sysAdlTypeSetObjective Or _
                sourceType = sysAdlTypeSetStakeholder Or _
                sourceType = sysAdlTypeSetQuality) Then
            
                If (sourceType = destinyType) Then Result = sysAdlTypeSetDependsOn
                
            ElseIf (sourceType = sysAdlTypeSetRole And destinyType = sysAdlTypeSetSystem) Then
                
                Result = sysAdlTypeSetDependsOn
                
            End If
            
        ElseIf (relationType = sysAdlTypeSetTransition) Then
        
            If (sourceType = sysAdlTypeSetLayer And destinyType = sysAdlTypeSetLayer) Then Result = sysAdlTypeSetTransition
            
        ElseIf (relationType = sysAdlTypeSetRepresents) Then
            
            If (destinyType = sysAdlTypeSetObjective) Then
            
                If (sourceType = sysAdlTypeSetSender Or _
                    sourceType = sysAdlTypeSetReceiver Or _
                    sourceType = sysAdlTypeSetSystem) Then Result = sysAdlTypeSetRepresents
            End If
        
        End If
        
    Relation = Result

    End Property
