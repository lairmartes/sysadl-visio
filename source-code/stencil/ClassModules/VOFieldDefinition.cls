VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VOFieldDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String
Private pLabel As String
Private pDescription As String
Private pErrorMessage As String
Private pFieldType As String
Private pStringRegExp As String
Private pDateAllowPast As Boolean
Private pDateAllowPresent As Boolean
Private pDateAllowFuture As Boolean
Private pTimeMinimum As Date
Private pTimeMaximum As Date
Private pValueMinimum As Double
Private pValueMaximum As Double
Private pValueOnlyInteger As Boolean
Private pElementType As String
Private pListDomain As Collection



Public Sub Init(ByVal aName As String, _
                ByVal aLabel As String, _
                ByVal aDescription As String, _
                ByVal anErrorMessage As String, _
                ByVal aFieldType As String)
                
    pName = aName
    pLabel = aLabel
    pDescription = aDescription
    pErrorMessage = anErrorMessage
    pFieldType = aFieldType
    pStringRegExp = ""
    pDateAllowPast = False
    pDateAllowPresent = False
    pDateAllowFuture = False
    pTimeMinimum = CDate("00:00")
    pTimeMaximum = CDate("00:00")
    pValueMinimum = 0
    pValueMaximum = 0
    pValueOnlyInteger = False
    pElementType = ""
    
    Set pListDomain = New Collection

                
End Sub

Public Sub InitTypeString(ByVal aStringRegExp As String)

    pStringRegExp = aStringRegExp


End Sub
                
                
Public Sub InitTypeDate(ByVal aDateAllowPast As Boolean, _
                ByVal aDateAllowPresent As Boolean, _
                ByVal aDateAllowFuture As Boolean)
                
    pDateAllowPast = aDateAllowPast
    pDateAllowPresent = aDateAllowPresent
    pDateAllowFuture = aDateAllowFuture
                
End Sub
                
                
Public Sub InitTypeTime(ByVal aTimeMinimum As Date, _
                ByVal aTimeMaximum As Date)
    pTimeMinimum = aTimeMinimum
    pTimeMaximum = aTimeMaximum
                
End Sub
                
Public Sub InitTypeValue(ByVal aValueMinimum As Double, _
                ByVal aValueMaximum As Double, _
                ByVal aValueOnlyInteger As Boolean)
                
    pValueMinimum = aValueMinimum
    pValueMaximum = aValueMaximum
    pValueOnlyInteger = aValueOnlyInteger
                
End Sub
                
Public Sub InitTypeElement(ByVal anElementType As String)

    pElementType = anElementType

End Sub

Public Sub AddListDomainItem(ByVal aDomainItem As Variant)

    pListDomain.Add aDomainItem

End Sub

Property Get Name()

    Name = pName
    
End Property

Property Get Label()

    Label = pLabel
    
End Property

Property Get Description()

    Description = pDescription
    
End Property

Property Get ErrorMessage()

    ErrorMessage = pErrorMessage
    
End Property

Property Get FieldType()

    FieldType = pFieldType
    
End Property

Property Get StringRegExp()

    StringRegExp = pStringRegExp
    
End Property

Property Get DateAllowPast()

    DateAllowPast = pDateAllowPast
    
End Property

Property Get DateAllowPresent()

    DateAllowPresent = pDateAllowPresent
    
End Property

Property Get DateAllowFuture()

    DateAllowFuture = pDateAllowFuture
    
End Property

Property Get TimeMinimum()

    TimeMinimum = pTimeMinimum
    
End Property

Property Get TimeMaximum()

    TimeMaximum = pTimeMaximum
    
End Property

Property Get ValueMinimum()

    ValueMinimum = pValueMinimum
    
End Property

Property Get ValueMaximum()

    ValueMaximum = pValueMaximum
    
End Property

Property Get ValueOnlyInteger()

    ValueOnlyInteger = pValueOnlyInteger
    
End Property

Property Get elementType()

    elementType = pElementType
    
End Property

Function GetListDomain() As Collection

    Set GetListDomain = pListDomain
    
End Function
