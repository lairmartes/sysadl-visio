VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private pFieldName As String
    Private pLabel As String
    Private pFieldType As Integer
    Private pFormat As String
    Private pPrompt As String
    Private pShowDesignOrder As Integer
    Private pShowDesignParenthesis As Boolean
    Private pShowCommentOrder As Integer
    Private pShowCommentParenthesis As Boolean
    Private pValidator As IValidator
    Private pIsFieldMandatory As Boolean
    
Public Sub Init(ByVal aFieldName As String, ByVal aLabel As String, ByVal aFieldType As Integer, ByVal APrompt As String, Optional APropertyDomainSet As Collection, Optional AFormat As String)

    pFieldName = aFieldName
    pLabel = aLabel
    pFieldType = aFieldType
    pPrompt = APrompt
    pIsFieldMandatory = False

    If APropertyDomainSet Is Nothing Then
    
        pFormat = AFormat
        
    Else
        
        Dim CurrentOption As Variant
        
        pFormat = ""
     
        For Each CurrentOption In APropertyDomainSet
            pFormat = pFormat + ";" + CurrentOption
        Next
        
    
    End If
    
End Sub

Public Sub InitDesignData(ByVal DesignOrder As Integer, _
                          ByVal DesignParenthesis As Boolean, _
                          ByVal CommentOrder As Integer, _
                          ByVal CommentParenthesis As Boolean)

    pShowDesignOrder = DesignOrder
    pShowDesignParenthesis = DesignParenthesis
    pShowCommentOrder = CommentOrder
    pShowCommentParenthesis = CommentParenthesis

End Sub

Public Sub InitValidator(ByVal aValidator As IValidator, ByVal IsFieldMandatory As Boolean)

    Set pValidator = aValidator
    
    pIsFieldMandatory = IsFieldMandatory

End Sub

Public Function ValidateValue(ByVal aValue As String) As String

    Dim Result As String
    
    Result = pValidator.Validate(aValue)
    
    ValidateValue = Result

End Function

Public Property Get Label() As String
    
    Label = pLabel
    
End Property

Public Property Get FieldType() As Integer
    
    FieldType = pFieldType
    
End Property
Public Property Get Prompt() As String
    
    Prompt = pPrompt
    
End Property

Public Property Get Format() As String
    
    Format = pFormat

End Property

Public Property Get FieldName() As String

    FieldName = pFieldName
    
End Property

Public Property Get ShowDesignOrder() As Integer
    
    ShowDesignOrder = pShowDesignOrder
    
End Property

Public Property Get ShowDesignParenthesis() As Boolean

    ShowDesignParenthesis = pShowDesignParenthesis
    
End Property

Public Property Get ShowCommentOrder() As Integer

    ShowCommentOrder = pShowCommentOrder
    
End Property
Public Property Get ShowCommentParenthesis() As Boolean

    ShowCommentParenthesis = pShowCommentParenthesis
    
End Property

Public Property Get IsFieldMandatory() As Boolean

    IsFieldMandatory = pIsFieldMandatory
    
End Property
