VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomPropertySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    Private pCustomPropertySet As New Collection
    Private pKey As String
    Private pDesignFields As Collection
    Private pCommentFields As Collection
    Private Const sysAdlVisualFunctionDesign = 0
    Private Const sysAdlVisualFunctionComment = 1
    
Public Sub Init(ByVal aKey As String)
    pKey = aKey
End Sub

Public Property Get Key() As String
    Key = pKey
End Property

Public Property Get customProperties() As Collection
    Set customProperties = pCustomPropertySet
End Property

Public Sub AddCustomProperty(ByVal ACustomProperty As CustomProperty)
    pCustomPropertySet.Add ACustomProperty
End Sub

Public Function Clone() As CustomPropertySet

    Dim Result As New CustomPropertySet
    Dim CustomPropertyCurrent As CustomProperty
    
    Result.Init pKey
    
    For Each CustomPropertyCurrent In pCustomPropertySet
    
        Result.AddCustomProperty CustomPropertyCurrent
        
    Next
    
    Set Clone = Result

End Function

Public Function GetDesignFields() As Collection

    Dim CurrentCustomProperty As CustomProperty
    Dim FieldsToDesign As New Collection
    Dim FieldOrder As Integer
    Dim FieldName As String

    If pDesignFields Is Nothing Then
    
        Set pDesignFields = New Collection
        
        For Each CurrentCustomProperty In pCustomPropertySet
        
            FieldOrder = CurrentCustomProperty.ShowDesignOrder
        
            If (FieldOrder > 0) Then
            
                FieldName = CurrentCustomProperty.FieldName
            
                FieldsToDesign.Add FieldName
                
            End If
        
        Next
        
        Set pDesignFields = SortPropertiesShowing(FieldsToDesign, sysAdlVisualFunctionDesign)
    
    End If
    
    Set GetDesignFields = pDesignFields
    
End Function

Public Function GetCommentFields() As Collection

    Dim CurrentCustomProperty As CustomProperty
    Dim FieldOrder As Integer
    Dim FieldsToComment As New Collection
    Dim FieldName As String

    If pCommentFields Is Nothing Then
    
        Set pCommentFields = New Collection
        
        For Each CurrentCustomProperty In pCustomPropertySet
        
            FieldOrder = CurrentCustomProperty.ShowCommentOrder
        
            If (FieldOrder > 0) Then
            
                FieldName = CurrentCustomProperty.FieldName
            
                FieldsToComment.Add FieldName
                
            End If
        
        Next
        
        Set pCommentFields = SortPropertiesShowing(FieldsToComment, sysAdlVisualFunctionComment)
    
    End If
    
    Set GetCommentFields = pCommentFields
    
End Function

Public Function GetCustomPropertyByFieldName(ByVal aFieldName As String) As CustomProperty

    Dim CurrentCustomProperty As CustomProperty
    Dim currentFieldName As String
    Dim Result As CustomProperty
    
    For Each CurrentCustomProperty In pCustomPropertySet
    
        currentFieldName = CurrentCustomProperty.FieldName
        
        
        If (currentFieldName = aFieldName) Then
        
            Set Result = CurrentCustomProperty
            
            Exit For
    
        End If
        
    Next
    
    Set GetCustomPropertyByFieldName = Result
    
End Function

' sort properties to be shown using buble sort algorithm
Private Function SortPropertiesShowing(ByVal FieldsToDesign As Collection, ByVal VisualFunctionConstant As Integer) As Collection

    Dim DataStructureSortable As New Dictionary
    Dim currentFieldName As String
    Dim NextFieldName As String
    Dim I As Integer
    Dim CountFields As Integer
    Dim CurrentShowOrder As Integer
    Dim CurrentCustomProperty As CustomProperty
    Dim NextCustomProperty As CustomProperty
    Dim CurrentFieldOrder As String
    Dim NextFieldOrder As String
    Dim Result As Collection
    Dim IsOrdering As Boolean
    
    
    Set Result = New Collection
    
    CountFields = FieldsToDesign.Count
    
    For I = 1 To CountFields
    
        currentFieldName = FieldsToDesign(I)
        
        DataStructureSortable.Add I, currentFieldName
        
    Next
    
    IsOrdering = True
    
    While (IsOrdering = True And CountFields > 1)
    
        For I = 1 To CountFields - 1
        
        
            IsOrdering = False
        
            currentFieldName = DataStructureSortable.Item(I)
            NextFieldName = DataStructureSortable.Item(I + 1)
            
            Set CurrentCustomProperty = GetCustomPropertyByFieldName(currentFieldName)
            Set NextCustomProperty = GetCustomPropertyByFieldName(NextFieldName)
            
            If (VisualFunctionConstant = sysAdlVisualFunctionDesign) Then
            
                CurrentFieldOrder = CurrentCustomProperty.ShowDesignOrder
                NextFieldOrder = NextCustomProperty.ShowDesignOrder
                
            ElseIf (VisualFunctionConstant = sysAdlVisualFunctionComment) Then
            
                CurrentFieldOrder = CurrentCustomProperty.ShowCommentOrder
                NextFieldOrder = NextCustomProperty.ShowCommentOrder
                
            End If
            
            
            If (CurrentFieldOrder > NextFieldOrder) Then
            
                DataStructureSortable.Item(I + 1) = currentFieldName
                DataStructureSortable.Item(I) = NextFieldName
                
                IsOrdering = True
            
            End If
            
        Next
        
    Wend
    
    
    For I = 1 To CountFields
        
        currentFieldName = DataStructureSortable.Item(I)
    
        Result.Add currentFieldName
            
    Next

    Set SortPropertiesShowing = Result
    
End Function
