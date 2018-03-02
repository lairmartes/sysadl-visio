VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VOElementDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pSysAdlType As String
    Private pStereotype As String
    Private pIsDeprecated As Boolean
    
    Private pElementFieldList As Collection
    

Public Sub Init(ByVal aSysAdlType As String, ByVal aStereotype As String, ByVal aIsDeprecated As Boolean)

    pSysAdlType = aSysAdlType
    pStereotype = aStereotype
    pIsDeprecated = aIsDeprecated
    
    Set pElementFieldList = New Collection
    
End Sub

Property Get sysadlType()

    sysadlType = pSysAdlType

End Property

Property Get Stereotype()

    Stereotype = pStereotype

End Property

Property Get IsDeprecated()

    IsDeprecated = pIsDeprecated
    
End Property


Public Sub AddField(ByVal aField As VOElementDefinitionField)

    pElementFieldList.Add aField

End Sub

Public Function GetElementFieldList() As Collection
    
    Set GetElementFieldList = pElementFieldList

End Function

