VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SysAdlElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    
    Private pSysAdlType As String
    Private pStereotype As String
    Private pNamespace As String
    Private pId As String
    Private pURLInfo As String
    Private pShapeList As New Collection
    Private pCustomPropertiesList As New CustomPropertySet
    Private pFields As New UtilSysAdlListString
    Private pDAOKey As String
    Private pIsDirty As Boolean
    Private pIsAlreadyHandlingEvent As Boolean

Public Sub InitKeyDAO(ByVal ADAOKey As Long)
    pDAOKey = Str$(ADAOKey)
End Sub

Public Property Get DAOKey() As String
    DAOKey = pDAOKey
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = pIsDirty
End Property
' start kind of the element
Public Sub Init(ByVal aSysAdlType As String, ByVal aStereotype As String)


    pSysAdlType = aSysAdlType
    pStereotype = aStereotype
    
    
End Sub

Public Sub EventHandlingMark()

    pIsAlreadyHandlingEvent = True

End Sub

Public Sub EventHandlingUnMark()

    pIsAlreadyHandlingEvent = False
    
End Sub


'Namespace identifies an element
Public Property Get namespace() As String
    
    namespace = pNamespace
    
End Property

Public Function Equals(ByVal aSysAdlElement As SysAdlElement) As Boolean
    
    Dim Result As Boolean
    
    Result = False
    
    If Not (aSysAdlElement Is Nothing) Then
    
        If Me.namespace = aSysAdlElement.namespace And Me.Id = aSysAdlElement.Id Then
        
            Result = True
            
        End If
        
    End If
    
    Equals = Result
    
End Function
'calcule key
Public Property Get Key() As String
    
    Key = Me.namespace + " " + Me.Id
    
End Property

'Id identifies an element
Public Property Get Id() As String

    Id = pId

End Property
'SysADLType
Public Property Get sysadlType() As String

    sysadlType = pSysAdlType
    
End Property
'Stereotype qualifies a SysADL Type
Public Property Get Stereotype() As String

    Dim Result As String
    
    If pStereotype = sysAdlStringConstantsEmpty Then
        
        Result = sysAdlNoStereotype
        
    Else
    
        Result = pStereotype
        
    End If

    Stereotype = Result
    
End Property
Public Property Get URLInfo() As String

    URLInfo = pURLInfo
    
End Property

'List of shapes that are bond to this model
Public Property Get ShapeViewerList() As Collection

    Set ShapeViewerList = pShapeList
    
End Property

'Init an elemet with an identifier
Public Sub InitBaseData(ByVal aNamespace As String, ByVal anId As String, ByVal AURLInfo As String)
    pNamespace = aNamespace
    pId = anId
    pURLInfo = AURLInfo
    
End Sub

'add fields values

Public Sub AddField(ByVal FieldName As String, ByVal Fieldvalue As String)

    Dim Field As UtilSysAdlItemString
    
    Set Field = New UtilSysAdlItemString
    
    Field.Init FieldName, Fieldvalue
    
    pFields.Add Field
    
End Sub

Public Property Get Fields() As UtilSysAdlListString
    
    Set Fields = pFields

End Property


'Respond to events.  Must be "overrided" by implementers
Public Sub HandleEvent(ByVal SomeEvent As EventSysAdl)

    Dim doNotBypassEventHandling As Boolean
    
    doNotBypassEventHandling = Not pIsAlreadyHandlingEvent

    If doNotBypassEventHandling Then
    
        EventHandlingMark

        Select Case SomeEvent.EventType
            Case sysAdlEventCreatedSysAdlElement
                ChangeToDirtyState
                CreateCustomFields
                RefreshCustomPropertiesFieldsValues
                CreateNewShape
            Case sysAdlEventChangedSysAdlElement
                ChangeToDirtyState
                RefreshShape
            Case sysAdlEventChangedStereotype
                ChangeToDirtyState
                CreateCustomFields
                RefreshShapeStereotype
                ClearShapesCustomProperties
                RefreshCustomPropertiesFieldsValues
                RefreshCustomPropertiesSet
                RefreshProperties
            Case sysAdlEventChangedCellValue
                ChangeToDirtyState
                RefreshCustomPropertiesFieldsValues
                RefreshProperties
                RefreshComments
            Case sysAdlEventElementRecovered
                RefreshShape
                CreateCustomFields
                RefreshCustomPropertiesSet
                RefreshCustomPropertiesFieldsValues
                RefreshProperties
                RefreshComments
                ChangeToCleanState
            Case sysAdlEventChangedURLInfo
                ChangeToDirtyState
                RefreshURLInfo
            Case sysAdlEventInvalidFieldFound
                RefreshShape
            Case sysAdlEventInvalidFieldCorrected
                RefreshComments
            Case sysAdlEventDocumentOpened
                RefreshShape
                CreateCustomFields
                RefreshCustomPropertiesSet
                RefreshCustomPropertiesFieldsValues
                RefreshProperties
                RefreshComments
                ChangeToCleanState
            Case sysAdlEventElementPersisted
                ChangeToCleanState
            Case sysAdlEventDocumentSavedAs
                ChangeToDirtyState
    
        End Select
        
        EventHandlingUnMark

    End If

End Sub

Private Sub ChangeToDirtyState()
    pIsDirty = True
End Sub

Private Sub ChangeToCleanState()

    pIsDirty = False
    
    pFields.CleanDirtyItems
    
End Sub


'Add a visio element to modifie its view
Public Sub AddShapeToList(ByVal AViewer As ShapeViewer)

    Dim CurrentShape As ShapeViewer
    Dim ElementAlreadyInList As Boolean
    
    For Each CurrentShape In pShapeList
        If CurrentShape.GetShapeId = AViewer.GetShapeId Then
            ElementAlreadyInList = True
            Exit For
        End If
    Next
    
    If Not ElementAlreadyInList Then
        pShapeList.Add AViewer
    End If

End Sub
'Remove a visio element from shape list
Public Sub RemoveShapeFromList(ByVal AViewer As ShapeViewer)

    Dim CurrentShape As ShapeViewer
    Dim CurrentIndex As Integer
    
    CurrentIndex = 0
    
    For Each CurrentShape In pShapeList
        CurrentIndex = CurrentIndex + 1
        If CurrentShape.GetShapeId = AViewer.GetShapeId Then
            pShapeList.Remove CurrentIndex
            Exit For
        End If
    Next
    
End Sub
'iniciar lista de propriedades
Private Sub RefreshCustomPropertiesSet()

    Dim CurrentShapeViewer As ShapeViewer
    
    For Each CurrentShapeViewer In Me.ShapeViewerList
        CurrentShapeViewer.InitProperties pCustomPropertiesList
    Next

End Sub

'Initialize an element with data of other.  Use it to copy element data to this element.
Public Sub InitWithSysAdlElement(ByVal aSysAdlElement As SysAdlElement)

    pSysAdlType = aSysAdlElement.sysadlType
    pNamespace = aSysAdlElement.namespace
    pId = aSysAdlElement.Id
    pStereotype = aSysAdlElement.Stereotype
    pURLInfo = aSysAdlElement.URLInfo
    
    Set pFields = aSysAdlElement.Fields
    
    Set pShapeList = aSysAdlElement.ShapeViewerList

End Sub


'Call this when a shape is created
Private Sub CreateNewShape()

    ClearShapesCustomProperties
    
    RefreshCustomPropertiesSet

    RefreshShapeStereotype
    RefreshText
    RefreshURLInfo

End Sub
'Call this when a shape is created
Private Sub RefreshShape()

    ClearShapesCustomProperties
    
    RefreshProperties

End Sub

'Call this when a shape is open (changed key)
Private Sub RefreshProperties()

    Dim CurrentCustomProperty As CustomProperty
    Dim currentField As UtilSysAdlItemString
    Dim currentFieldName As String
    Dim IsBaseField As String
    Dim currentFieldValue As String

    RefreshShapeViewProperty sysAdlKeyCustPropRowNameId, pId
    RefreshShapeViewProperty sysAdlKeyCustPropRowNameNamespace, pNamespace
    RefreshShapeViewProperty sysAdlKeyCustPropRowNameStereotype, pStereotype
    RefreshShapeViewProperty sysAdlKeyCustPropRowNameURLInfo, pURLInfo
    
    For Each CurrentCustomProperty In pCustomPropertiesList.customProperties
    
        Set currentField = New UtilSysAdlItemString
            
        currentFieldName = CurrentCustomProperty.FieldName
            
        currentFieldValue = pFields.Item(currentFieldName)
            
        IsBaseField = False
            
        If currentFieldName = sysAdlKeyCustPropRowNameId Then IsBaseField = True
        If currentFieldName = sysAdlKeyCustPropRowNameNamespace Then IsBaseField = True
        If currentFieldName = sysAdlKeyCustPropRowNameURLInfo Then IsBaseField = True
        If currentFieldName = sysAdlKeyCustPropRowNameStereotype Then IsBaseField = True
                
        If IsBaseField <> True Then
            
            currentField.Init currentFieldName, currentFieldValue
                
            If pFields.KeyExists(currentFieldName) Then
                
                RefreshShapeViewProperty currentFieldName, currentFieldValue
                     
            End If
    
        End If
    Next
    
    RefreshText
    RefreshShapeStereotype
    RefreshURLInfo
    
End Sub

'Call this to clear custom properties
Private Sub ClearShapesCustomProperties()

    Dim CurrentShapeViewer As ShapeViewer
    
    For Each CurrentShapeViewer In Me.ShapeViewerList
        CurrentShapeViewer.ClearCustomProperties
    Next
End Sub
'Call this when a field has been changed from Shape
Public Sub ChangeFieldValue(ByVal FieldChangedName As String, ByVal FieldChangedNewValue As String)

    If Not pIsAlreadyHandlingEvent Then

        Select Case FieldChangedName
            Case sysAdlKeyCustPropRowNameId
                pId = FieldChangedNewValue
            Case sysAdlKeyCustPropRowNameNamespace
                pNamespace = FieldChangedNewValue
            Case sysAdlKeyCustPropRowNameStereotype
                pStereotype = FieldChangedNewValue
            Case sysAdlKeyCustPropRowNameURLInfo
                pURLInfo = FieldChangedNewValue
            Case Else
                ProcessChangeFieldValue FieldChangedName, FieldChangedNewValue
        End Select

    End If

End Sub

'method that returns if a SysAdlType is the same of other
Public Function IsSameSysAdlType(ByVal aSysAdlElement As SysAdlElement) As Boolean

    Dim Result As Boolean
    
    Result = False
    
    If Not (aSysAdlElement Is Nothing) Then

        If Me.sysadlType = aSysAdlElement.sysadlType Then
            Result = True
        Else
            Result = False
        End If
        
    End If
    
    IsSameSysAdlType = Result

End Function

'validate fields returning issues messages

Public Function ValidateFields() As Collection

    Dim Result As New Collection
    Dim currentMessage As Variant
    Dim CurrentCustomProperty As CustomProperty
    Dim currentFieldValue As String
    Dim currentFieldName As String
    Dim fieldMustBeValidated As Boolean
    Dim fieldValueIsNotBlank As Boolean
    
    For Each CurrentCustomProperty In pCustomPropertiesList.customProperties
    
        currentFieldName = CurrentCustomProperty.FieldName
        
        If currentFieldName = sysAdlKeyCustPropRowNameId Then
            currentFieldValue = Id
        ElseIf currentFieldName = sysAdlKeyCustPropRowNameNamespace Then
            currentFieldValue = namespace
        ElseIf currentFieldName = sysAdlKeyCustPropRowNameURLInfo Then
            currentFieldValue = URLInfo
        ElseIf currentFieldName = sysAdlKeyCustPropRowNameStereotype Then
            currentFieldValue = Stereotype
        Else:
            currentFieldValue = pFields.Item(currentFieldName)
        End If
        
        fieldValueIsNotBlank = (currentFieldValue <> sysAdlStringConstantsEmpty)
        
        If CurrentCustomProperty.IsFieldMandatory Then
        
            fieldMustBeValidated = (pFields.ItemIsDirty(currentFieldName) Or IsBaseField(currentFieldName))
        
        Else
        
            fieldMustBeValidated = fieldValueIsNotBlank
        
        End If
        
        If fieldMustBeValidated Then
        
            currentMessage = CurrentCustomProperty.ValidateValue(currentFieldValue)
            
        Else
        
            currentMessage = sysAdlStringConstantsEmpty
            
        End If
        
        If currentMessage <> sysAdlStringConstantsEmpty Then
        
            Result.Add CurrentCustomProperty.Label + ": " + currentMessage
            
        End If
    
    Next
    
    
    Set ValidateFields = Result

End Function
'Call this to change stereotype of all shapes
Private Sub RefreshShapeStereotype()

    Dim CurrentShapeViewer As ShapeViewer
    
    For Each CurrentShapeViewer In Me.ShapeViewerList
        CurrentShapeViewer.ChangeStereotype pStereotype
    Next

End Sub
'Call to change refresh properties in view
Private Sub RefreshShapeViewProperty(ByVal aProperty As String, ByVal APropertyValue)
        
    Dim CurrentShapeViewer As ShapeViewer
    
    For Each CurrentShapeViewer In Me.ShapeViewerList
        CurrentShapeViewer.RefreshProperty aProperty, APropertyValue
    Next
    
End Sub

'Call to change refresh properties in view
Private Sub ChangeText(ByVal TextLines As Collection)
    
    Dim CurrentShapeViewer As ShapeViewer
    
    For Each CurrentShapeViewer In Me.ShapeViewerList
        CurrentShapeViewer.ChangeText TextLines
    Next

End Sub

'Call to change URL Info
Private Sub RefreshURLInfo()
    
    Dim CurrentShapeViewer As ShapeViewer
    Dim ConvertedURL As String
    Dim BasePub As String
    Dim PathFromNamespace As String
    
    BasePub = FactoryConfigProperty.GetProperty(sysAdlConfigPropertyBaseURLPublish)
    PathFromNamespace = Replace(pNamespace, ".", "/")
        
    ConvertedURL = pURLInfo
    ConvertedURL = Replace(ConvertedURL, sysAdlStringConstantsConvertBasePub, BasePub)
    ConvertedURL = Replace(ConvertedURL, sysAdlStringConstantsConvertNamespace, PathFromNamespace)
    
    For Each CurrentShapeViewer In Me.ShapeViewerList
        CurrentShapeViewer.ChangeURLInfo ConvertedURL
    Next

End Sub
'Call when a property is changed
Private Sub RefreshText()

    Dim DesignFields As Collection
    Dim TextLines As Collection
    Dim CurrentDesignFieldName As Variant
    Dim CustomPropertyDesign As CustomProperty
    Dim ValueToShow As String
    Dim TextToAdd As String
    
    Set TextLines = New Collection
    
    Set DesignFields = pCustomPropertiesList.GetDesignFields

    For Each CurrentDesignFieldName In DesignFields
    
        TextToAdd = ""
    
        Set CustomPropertyDesign = pCustomPropertiesList.GetCustomPropertyByFieldName(CurrentDesignFieldName)
        
        ValueToShow = pFields.Item(CurrentDesignFieldName)
        
        If CustomPropertyDesign.ShowDesignParenthesis And ValueToShow <> sysAdlStringConstantsEmpty Then
        
            TextToAdd = sysAdlStringConstantsParenthesisOpen + ValueToShow + sysAdlStringConstantsParenthesisClose
    
        Else
        
            TextToAdd = ValueToShow
            
        End If
        
        TextLines.Add TextToAdd
        
    Next
    
    ChangeText TextLines
    
End Sub



Public Function PropertyHasBeenChanged(ByVal FieldName As String, ByVal Fieldvalue As String) As Boolean

    Dim StereotypeValue As String
    
    StereotypeValue = sysAdlStringConstantsEmpty
    
    If Me.Stereotype <> sysAdlNoStereotype Then
    
        StereotypeValue = Me.Stereotype
    
    End If

    Select Case FieldName
        Case sysAdlKeyCustPropLabelNamespace
            PropertyHasBeenChanged = (Fieldvalue <> Me.namespace)
        Case sysAdlKeyCustPropRowNameId
            PropertyHasBeenChanged = (Fieldvalue <> Me.Id)
        Case sysAdlKeyCustPropRowNameStereotype
            PropertyHasBeenChanged = (Fieldvalue <> StereotypeValue)
        Case sysAdlKeyCustPropRowNameURLInfo
            PropertyHasBeenChanged = (Fieldvalue <> Me.URLInfo)
        Case Else
            PropertyHasBeenChanged = (Fieldvalue <> GetFieldValue(FieldName))
    End Select

End Function



Private Sub RefreshComments()

    Dim CommentFields As Collection
    Dim TextLines As Collection
    Dim CurrentCommentFieldName As Variant
    Dim CustomPropertyComment As CustomProperty
    Dim ValueToShow As String
    Dim TextToAdd As String
    
    Set TextLines = New Collection
    
    Set CommentFields = pCustomPropertiesList.GetCommentFields

    For Each CurrentCommentFieldName In CommentFields
    
        TextToAdd = ""
    
        Set CustomPropertyComment = pCustomPropertiesList.GetCustomPropertyByFieldName(CurrentCommentFieldName)
        
        ValueToShow = pFields.Item(CurrentCommentFieldName)
        
        If CustomPropertyComment.ShowDesignParenthesis And ValueToShow <> sysAdlStringConstantsEmpty Then
        
            TextToAdd = sysAdlStringConstantsParenthesisOpen + ValueToShow + sysAdlStringConstantsParenthesisClose
    
        Else
        
            TextToAdd = ValueToShow
            
        End If
        
        TextLines.Add TextToAdd
        
    Next

    ChangeComments TextLines

End Sub

Private Sub ChangeComments(ByVal Comments As Collection)

    Dim ShapeViewerCount As Integer
    Dim CurrentShapeViewer As ShapeViewer
    
    ShapeViewerCount = ShapeViewerList.Count

    For Each CurrentShapeViewer In Me.ShapeViewerList
        CurrentShapeViewer.ChangeComments Comments
    Next

End Sub

Private Sub ProcessChangeFieldValue(ByVal FieldName As String, ByVal Fieldvalue As String)

    pFields.ChangeValue FieldName, Fieldvalue
    
End Sub

Private Function GetFieldValue(ByVal FieldName As String) As String

    Dim Result As String
    
    Result = pFields.Item(FieldName)
    
    GetFieldValue = Result
    
End Function

Private Sub RefreshCustomPropertiesFieldsValues()


    Dim CurrentCustomProperty As CustomProperty
    Dim currentField As UtilSysAdlItemString
    Dim currentFieldName As String
    Dim baseFieldFound As String
    Dim currentFieldValue As String

    For Each CurrentCustomProperty In pCustomPropertiesList.customProperties
    
        Set currentField = New UtilSysAdlItemString
        
        currentFieldName = CurrentCustomProperty.FieldName
        
        currentFieldValue = pFields.Item(currentFieldName)
        
        baseFieldFound = IsBaseField(currentFieldName)
                    
        If baseFieldFound <> True Then
        
            currentField.Init currentFieldName, currentFieldValue
            
            If pFields.KeyExists(currentFieldName) Then
            
                pFields.ChangeValue currentFieldName, currentFieldValue
            
            Else
                pFields.Add currentField
                 
            End If

        End If
        

    
    Next
End Sub


Private Sub CreateCustomFields()

    Set pCustomPropertiesList = FactoryCustomProperty.Create(pSysAdlType, pStereotype)

End Sub

Private Function IsBaseField(ByVal FieldName As String) As Boolean

    Dim Result As Boolean
    
    Result = False

    If FieldName = sysAdlKeyCustPropRowNameId Then Result = True
    If FieldName = sysAdlKeyCustPropRowNameNamespace Then Result = True
    If FieldName = sysAdlKeyCustPropRowNameURLInfo Then Result = True
    If FieldName = sysAdlKeyCustPropRowNameStereotype Then Result = True
        
    IsBaseField = Result

End Function

