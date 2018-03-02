VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiagramInnerRelationSysADL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Dim ControllerOuter As shapeController
    Dim ControllerInner As shapeController
    Dim RelationId As String
    
    Public Sub Init(ByVal anOuterController As shapeController, ByVal anInnerController As shapeController)
    
        Set ControllerOuter = anOuterController
        Set ControllerInner = anInnerController
        
        RelationId = sysAdlStringConstantsEmpty
    
    End Sub

    Public Property Get Relation() As String
    
        Dim Result As String
        Dim outerShapeViewer As ShapeViewer
        
        Result = sysAdlStringConstantsEmpty
        
        If (ControllerOuter.ShapeSysAdlType = sysAdlTypeSetLayer) Then
        
            If ((ControllerInner.ShapeSysAdlType = sysAdlTypeSetSystem) Or _
                 (ControllerInner.ShapeSysAdlType = sysAdlTypeSetNode) Or _
                 (ControllerInner.ShapeSysAdlType = sysAdlTypeSetLayer) Or _
                 (ControllerInner.ShapeSysAdlType = sysAdlTypeSetStakeholder)) Then
                 
                    Set outerShapeViewer = ControllerOuter.ShapeViewer
                                    
                    If (outerShapeViewer.IsOuterShapeOf(ControllerInner.ShapeViewer)) Then
                
                        Result = sysAdlTypeSetComposedBy
                        
                        RelationId = outerShapeViewer.GetShapeId + ControllerInner.GetShapeUniqueId
                        
                    End If
                    
            End If
        
        ElseIf (ControllerOuter.ShapeSysAdlType = sysAdlTypeSetNode) Or _
               (ControllerOuter.ShapeSysAdlType = sysAdlTypeSetSystem) Then
        
            If (ControllerInner.ShapeSysAdlType = ControllerOuter.ShapeSysAdlType) Then
                 
                    Set outerShapeViewer = ControllerOuter.ShapeViewer
                                    
                    If (outerShapeViewer.IsOuterShapeOf(ControllerInner.ShapeViewer)) Then
                
                        Result = sysAdlTypeSetComposedBy
                        
                        RelationId = outerShapeViewer.GetShapeId + ControllerInner.GetShapeUniqueId
                        
                    End If
                    
            End If
            
        End If
        
    Relation = Result
        
    
    End Property
    
    Public Property Get RelationSource() As shapeController

        Set RelationSource = ControllerOuter

    End Property
    
    Public Property Get RelationDestiny() As shapeController
    
        Set RelationDestiny = ControllerInner
        
    End Property
    
    
    Public Property Get Id() As String
    
        Id = RelationId

    End Property
    
