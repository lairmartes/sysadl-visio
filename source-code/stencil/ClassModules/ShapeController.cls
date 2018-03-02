VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Dim MyShapeViewer As New ShapeViewer
Attribute MyShapeViewer.VB_VarHelpID = -1
    Dim MySysAdlElement As SysAdlElement
    Dim pMarkedAsRemoved As Boolean
    
Public Sub Init(ByVal SomeShape As IVShape, ByVal SomeSysAdlElement As SysAdlElement)
    
    MyShapeViewer.InitVisioShape SomeShape
    
    Set MySysAdlElement = SomeSysAdlElement
        
    AddShapeToElement

End Sub
Public Function GetShapeUniqueId() As String

    Dim Result As String
    
    If pMarkedAsRemoved Then
        Result = sysAdlStringConstantsEmpty
    Else
        Result = MyShapeViewer.GetShapeId
    End If
    
    GetShapeUniqueId = Result
    
End Function
Public Sub ChangeSysAdlElement(ByVal aSysAdlElement As SysAdlElement)
    RemoveShapeFromElement
    Set MySysAdlElement = aSysAdlElement
    AddShapeToElement

End Sub

Public Property Get shape() As shape
    
    Set shape = MyShapeViewer.shape

End Property

Public Property Get SysAdlElement() As SysAdlElement

    Set SysAdlElement = MySysAdlElement

End Property

Public Property Get ShapeSysAdlType() As String

    Dim SysADLTypeFromShape As String
    
    SysADLTypeFromShape = MyShapeViewer.GetSysAdlTypeOfShape()

    ShapeSysAdlType = SysADLTypeFromShape
    
End Property

Public Sub TriggerEvent(ByVal AnEvent As EventSysAdl)

    MySysAdlElement.HandleEvent AnEvent

End Sub
Private Sub AddShapeToElement()
        
    Me.SysAdlElement.AddShapeToList MyShapeViewer

End Sub

Private Sub MarkAsRemoved()

    pMarkedAsRemoved = True

End Sub

Public Sub ProcessControllerRemove()

    RemoveShapeFromElement
    MarkAsRemoved

End Sub

Private Sub RemoveShapeFromElement()

    Me.SysAdlElement.RemoveShapeFromList MyShapeViewer

End Sub

Public Property Get ShapeViewer()

    Set ShapeViewer = MyShapeViewer
    
End Property
Public Property Get MarkedAsRemoved() As Boolean
    
    MarkedAsRemoved = pMarkedAsRemoved

End Property

Public Sub FireCellChanged(ByVal Cell As IVCell)

    On Error GoTo IgnoreErrorWhileCreatingShape:
    
    Dim CommandQueue As New QueueCommandExecution
    
    If Cell.Section = visSectionProp And Cell.column = visCustPropsValue Then
    
        Dim CommandChangeProperties As New CommandChangeCustomProperty
        Dim CommandChangeStereotype As New CommandChangeStereotype
        Dim CommandChangeKey As New CommandChangeKey
        Dim CommandChangeURLInfo As New CommandChangeURLInfo
        Dim FieldChangedName As String
        Dim FieldChangedValue As String
        Dim FieldHasBeenChanged As Boolean
        
        'MsgBox "Trantando evento de troca para " + Cell.RowNameU
        
        FieldChangedName = Cell.RowNameU
        FieldChangedValue = Cell.ResultStr(Visio.visNone)
        
        FieldHasBeenChanged = Me.SysAdlElement.PropertyHasBeenChanged(FieldChangedName, FieldChangedValue)
        
        If FieldHasBeenChanged Then
        
            'check if stereotype has been changed
            If FieldChangedName = sysAdlKeyCustPropRowNameStereotype Then
                ' add command to change stereotype in command queue
                CommandChangeStereotype.Init Me, FieldChangedValue
                CommandQueue.AddCommand CommandChangeStereotype
            End If
            ' add command to respond to key changed
            If FieldChangedName = sysAdlKeyCustPropRowNameId Or FieldChangedName = sysAdlKeyCustPropRowNameNamespace Then
                Dim NewNamespace As String
                Dim NewId As String
                If FieldChangedName = sysAdlKeyCustPropRowNameNamespace Then
                    NewNamespace = FieldChangedValue
                    NewId = Me.SysAdlElement.Id
                Else
                    NewNamespace = Me.SysAdlElement.namespace
                    NewId = FieldChangedValue
                End If
                    
                CommandChangeKey.Init Me, NewNamespace, NewId
                CommandQueue.AddCommand CommandChangeKey
            End If
            
            ' add command to refresh values of further shapes
            CommandChangeProperties.Init Me, FieldChangedName, FieldChangedValue
            CommandQueue.AddCommand CommandChangeProperties

            If FieldChangedName = sysAdlKeyCustPropRowNameURLInfo Then
                CommandChangeURLInfo.Init Me
                CommandQueue.AddCommand CommandChangeURLInfo
            End If

            CommandQueue.ExecuteCommandList
        
        End If

        
    End If
    
    Exit Sub

IgnoreErrorWhileCreatingShape:
' =================================================================================================
' Nota: Essa captura de erros foi incluída por que ocorre um erro quando o shape é criado
'       e tentamos obter o dado da célula.  Esse erro não ocorre depois que o shape foi
'       criado e a função é chamada normalmente.
' =================================================================================================
    If Err.Number <> -2032466967 Then
        MsgBox Err.Description + " (" + Conversion.Str(Err.Number) + ")"
    End If
    Exit Sub
End Sub


