VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandChangeCustomProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICommand
    Private MyShapeController As shapeController
    Private FieldChangedName As String
    Private FieldChangedValue As String
    
Public Sub Init(ByVal aShapeController As shapeController, ByVal aFieldName As String, ByVal AFieldValue As String)

    Set MyShapeController = aShapeController
    FieldChangedName = aFieldName
    FieldChangedValue = AFieldValue
    
End Sub

Public Sub ICommand_Execute()

    Dim ChangedCellEvent As EventSysAdl
    Dim SysAdlElementControlled As SysAdlElement
    
    Set SysAdlElementControlled = MyShapeController.SysAdlElement
    
    SysAdlElementControlled.ChangeFieldValue FieldChangedName, FieldChangedValue
    
    Set ChangedCellEvent = FactoryEvent.CreateEvent(sysAdlEventChangedCellValue)
    
    MyShapeController.TriggerEvent ChangedCellEvent

End Sub

