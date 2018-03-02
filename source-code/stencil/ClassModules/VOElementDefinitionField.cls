VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VOElementDefinitionField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pFieldName As String
Private pFieldMandatory As Boolean
Private pFieldOrder As Double
Private pShowDesignOrder As Double
Private pShowDesignParenthesis As Boolean
Private pShowCommentsOrder As Double
Private pShowCommentsParenthesis As Boolean

Public Sub Init(ByVal aFieldName As String, _
                 ByVal aFieldMandatory As Boolean, _
                 ByVal aFieldOrder As String)
                 
    pFieldName = aFieldName
    pFieldMandatory = aFieldMandatory
    pFieldOrder = aFieldOrder
    pShowDesignOrder = 0
    pShowDesignParenthesis = False
    pShowCommentsOrder = 0
    pShowCommentsParenthesis = False
                 
End Sub

Property Get FieldName() As String

    FieldName = pFieldName
    
End Property

Property Get FieldMandatory() As Boolean

    FieldMandatory = pFieldMandatory
    
End Property

Property Get FieldOrder() As Double

    FieldOrder = pFieldOrder
    
End Property

Property Get ShowDesignOrder() As Double

    ShowDesignOrder = pShowDesignOrder
    
End Property

Property Get ShowDesignParenthesis() As Double

    ShowDesignParenthesis = pShowDesignParenthesis
    
End Property

Property Get ShowCommentsOrder() As Double

    ShowCommentsOrder = pShowCommentsOrder
    
End Property
    
Property Get ShowCommentsParenthesis() As Double

    ShowCommentsParenthesis = pShowCommentsParenthesis
    
End Property
Public Sub InitDesignPreferences(ByVal aShowDesignOrder As Double, ByVal aShowDesignParenthesis As Boolean)

    pShowDesignOrder = aShowDesignOrder
    pShowDesignParenthesis = aShowDesignParenthesis
    
End Sub

Public Sub InitCommentsPreferences(ByVal aShowCommentsOrder As Double, ByVal aShowCommentsParenthesis As Boolean)

    pShowCommentsOrder = aShowCommentsOrder
    pShowCommentsParenthesis = aShowCommentsParenthesis
    
End Sub
