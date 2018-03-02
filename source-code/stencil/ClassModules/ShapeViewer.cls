VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private VisioShape As IVShape
    Private Const STEREOTYPE_OPENTAG = "<<"
    Private Const STEREOTYPE_CLOSETAG = ">>"
    Private CurrentStereotype As String
    Private Const ZERO_TOLERANCE = 0
    

Public Sub InitProperties(ByVal Properties As CustomPropertySet)


    Dim CurrentProperty As CustomProperty
    Dim CurrentLineProperty As Integer
    Dim CustomPropertyNotAddedYet As Boolean
    
    
        For Each CurrentProperty In Properties.customProperties
        
            CustomPropertyNotAddedYet = Not VisioShape.CellExists("Prop." + CurrentProperty.FieldName, False)
        
            If CustomPropertyNotAddedYet Then
        
                CurrentLineProperty = VisioShape.AddRow(visSectionProp, visRowLast, visTagDefault)
                
                VisioShape.CellsSRC(visSectionProp, CurrentLineProperty, visTagDefault).RowNameU = CurrentProperty.FieldName
                VisioShape.CellsSRC(visSectionProp, CurrentLineProperty, visCustPropsValue).FormulaU = ""
                VisioShape.CellsSRC(visSectionProp, CurrentLineProperty, visCustPropsType).FormulaU = CurrentProperty.FieldType
                VisioShape.CellsSRC(visSectionProp, CurrentLineProperty, visCustPropsLabel).FormulaU = """" + CurrentProperty.Label + """"
                VisioShape.CellsSRC(visSectionProp, CurrentLineProperty, visCustPropsPrompt).FormulaU = """" + CurrentProperty.Prompt + """"
                VisioShape.CellsSRC(visSectionProp, CurrentLineProperty, visCustPropsFormat).FormulaU = """" + CurrentProperty.Format + """"
                VisioShape.CellsSRC(visSectionAction, CurrentLineProperty, visActionAction).FormulaU = 0
                
            
            End If
        Next
End Sub

Public Sub InitVisioShape(ByVal AShape As IVShape)

    Dim ShapeId As String
    
    ShapeId = AShape.UniqueID(visGetOrMakeGUID)  ' cria um Id para o Shape

    Set VisioShape = AShape

End Sub

Public Sub ChangeWithOtherShapeViewer(ByVal aShapeViewer As ShapeViewer)

End Sub

Public Function GetSysAdlTypeOfShape() As String

    Dim Result As String
    
    Result = GUIServices.PrepareMasterNameForElement(VisioShape.Master.Name)

    GetSysAdlTypeOfShape = Result

End Function

Public Sub ClearCustomProperties()

    VisioShape.DeleteSection visSectionProp

End Sub

Public Sub ChangeStereotype(ByVal Stereotype As String)

    If Stereotype = sysAdlNoStereotype Then Stereotype = sysAdlStringConstantsEmpty
    
    CurrentStereotype = Stereotype

    If HasSpecificLabelToStereotype Then
        
        ChangeStereotypeWithSpecificLabel Stereotype
        
    End If

End Sub

Private Function HasSpecificLabelToStereotype() As Boolean

    Dim StereotypeShape As shape
    
    Set StereotypeShape = GetChildShapeByName("Stereotype")
    
    HasSpecificLabelToStereotype = Not StereotypeShape Is Nothing

End Function

Private Sub ChangeStereotypeWithSpecificLabel(ByVal Stereotype As String)

        Dim DisplayText As String
        Dim CurrentDisplay As String
        Dim StereotypeShape As shape

        Set StereotypeShape = GetChildShapeByName("Stereotype")
        
        If Stereotype = "" Then
            DisplayText = ""
        Else
            DisplayText = STEREOTYPE_OPENTAG + Stereotype + STEREOTYPE_CLOSETAG
        End If
        
        CurrentDisplay = StereotypeShape.text
        
        If CurrentDisplay <> DisplayText Then
            StereotypeShape.CellsSRC(visSectionObject, visRowLock, visLockTextEdit).FormulaU = 0 'unlock text edit
            StereotypeShape.text = DisplayText
            StereotypeShape.CellsSRC(visSectionObject, visRowLock, visLockTextEdit).FormulaU = 1 'relock text edit
        End If

End Sub

Public Sub ChangeURLInfo(ByVal NewURL As String)

    Dim RowNumber As Integer

    If Not VisioShape.SectionExists(visSectionHyperlink, False) Then
        VisioShape.AddSection visSectionHyperlink
        RowNumber = VisioShape.AddNamedRow(visSectionHyperlink, sysAdlKeyCustPropRowNameURLInfo, visRow1stHyperlink)
    Else
        If NewURL = sysAdlStringConstantsEmpty Then
            VisioShape.DeleteSection visSectionHyperlink
        End If
    End If
    If NewURL <> sysAdlStringConstantsEmpty Then
        VisioShape.CellsSRC(visSectionHyperlink, visRow1stHyperlink + RowNumber, visHLinkAddress).FormulaU = """" + NewURL + """"
    End If

End Sub

Public Sub RefreshProperty(ByVal PropertyName As String, ByVal PropertyValue As String)

    Dim CurrentValue As String
    
    If PropertyName = sysAdlKeyCustPropRowNameStereotype And PropertyValue = sysAdlNoStereotype Then
    
        PropertyValue = sysAdlStringConstantsEmpty
        
    End If
    
    PropertyValue = GUIServices.PrepareFieldForView(PropertyValue)
    
    If (IsFieldIncluded(PropertyName)) Then
        CurrentValue = VisioShape.CellsU("Prop." + PropertyName).ResultStr(Visio.visNone)
        CurrentValue = GUIServices.PrepareFieldForView(CurrentValue)
        
        If CurrentValue <> PropertyValue Then
            VisioShape.CellsU("Prop." + PropertyName).FormulaU = """" + PropertyValue + """"
        End If
    End If

End Sub

Private Function GetChildShapeByName(ByVal ShapeName As String) As shape

    Dim ChildShapeCollection As Shapes

    Set ChildShapeCollection = VisioShape.Shapes
    
    For Each GetChildShapeByName In ChildShapeCollection
        If GetChildShapeByName.Name = ShapeName Then
            Exit For
        End If
    Next
    
End Function

Public Sub ChangeText(ByVal TextLines As Collection)
    Dim DisplayText As String
    Dim IsFirstLine As Boolean
    Dim TextLine As Variant
    Dim LabelShape As shape
    Dim TextSeparator As String
    
    TextSeparator = vbCrLf
    
    'Earlier versions showed all data in one line for Channels.  This has changed duo calculation line capabilities.
    'If Me.GetSysAdlTypeOfShape = sysAdlTypeSetChannel Then
        
        'TextSeparator = sysAdlStringConstantsBlankSpace
        
    'End If
    
    IsFirstLine = True

    For Each TextLine In TextLines
        If (TextLine <> "") Then
            If (IsFirstLine) Then
                DisplayText = TextLine
                IsFirstLine = False
            Else
                DisplayText = DisplayText + TextSeparator + TextLine
            End If
        End If
    Next

    If CurrentStereotype <> sysAdlStringConstantsEmpty And Not HasSpecificLabelToStereotype Then
        DisplayText = STEREOTYPE_OPENTAG + CurrentStereotype + STEREOTYPE_CLOSETAG + TextSeparator + DisplayText
    End If

    UpdateLabelText (DisplayText)

End Sub

Private Sub UpdateLabelText(ByVal NewText As String)
    
    Dim LabelShape As shape
    
    Set LabelShape = GetLabelShape
    
    LabelShape.CellsSRC(visSectionObject, visRowLock, visLockTextEdit).FormulaU = 0 'unlock text edit
    LabelShape.text = NewText
    LabelShape.CellsSRC(visSectionObject, visRowLock, visLockTextEdit).FormulaU = 1 'relock text edit


End Sub


Public Function GetShapeId() As String

    GetShapeId = VisioShape.UniqueID(visGetGUID)
    
End Function

Private Function IsFieldIncluded(ByVal FieldName As String) As Boolean
    
    Dim Result As Integer

    Result = VisioShape.CellExistsU("Prop." + FieldName, visSectionProp)
    
    If Result = -1 Then
        IsFieldIncluded = True
    End If

End Function


Private Function GetShapeToHighLight() As shape

    Dim Result As shape
    
    Set Result = GetChildShapeByName("Picture")
    
    If Result Is Nothing Then
        Set Result = GetChildShapeByName("Label")
    End If

    If Result Is Nothing Then
        Set Result = VisioShape
    End If

    Set GetShapeToHighLight = Result

End Function

Public Property Get shape() As IVShape

    Set shape = VisioShape

End Property

Private Function GetLabelShape() As shape

    Dim Result As shape
    Set Result = GetChildShapeByName("Label")
        
    If Result Is Nothing Then
        Set Result = VisioShape
    End If

    Set GetLabelShape = Result

End Function

Public Function IsSameShapeViewer(ByVal otherShapeViewer As ShapeViewer) As Boolean

    Dim Result As Boolean
    
    Result = (GetShapeId = otherShapeViewer.GetShapeId)
    
    IsSameShapeViewer = Result
    
End Function

Public Sub ChangeComments(ByVal TextLines As Collection)

    Dim DisplayText As String
    Dim IsFirstLine As Boolean
    Dim TextLine As Variant
        
    IsFirstLine = True

    For Each TextLine In TextLines
        If (TextLine <> "") Then
            If (IsFirstLine) Then
                DisplayText = TextLine
                IsFirstLine = False
            Else
                DisplayText = DisplayText + vbCrLf + TextLine
            End If
        End If
    Next

    DisplayText = GUIServices.PrepareFieldForView(DisplayText)

    VisioShape.CellsSRC(visSectionObject, visRowMisc, visComment).FormulaU = """" + DisplayText + """"
    
End Sub

Public Function IsOuterShapeOf(ByVal otherShapeViewer As ShapeViewer) As Boolean

    Dim Result As Boolean

    Dim shapeElement As shape

    'get shape element of sent ShapeViewer
    Set shapeElement = otherShapeViewer.shape
    
    'A shape can contain another shape.
    'Shape "Me.shape" contains shape "shapeElement" if shape "Me.shape" encloses every region and path of shape "shapeElement".
    Result = (Me.shape.SpatialRelation(shapeElement, ZERO_TOLERANCE, visSpatialIncludeHidden) = VisSpatialRelationCodes.visSpatialContain)
    
    IsOuterShapeOf = Result

End Function
