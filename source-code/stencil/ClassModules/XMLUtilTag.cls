VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLUtilTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pName As String
    Private pIsCollapsed As Boolean
    Private pProperties As Collection
    
    Private pOpeningTag As String
    Private pClosingTag As String
    Private pClosedTag As String
    
    Public Sub Init(ByVal Name As String, Optional isCollapsed As Boolean)
    
        Set pProperties = New Collection

        pName = Name
        
        pIsCollapsed = isCollapsed
        
    End Sub

    Public Sub AddProperty(ByVal aProperty As XMLUtilTagValue)
    
        pProperties.Add aProperty
    
    End Sub
    


    Private Function ToStringProperties() As String
    
        Dim Result As String
        Dim CurrentProperty As XMLUtilTagValue

        Result = sysAdlStringConstantsEmpty
        
        For Each CurrentProperty In pProperties
        
            If (pIsCollapsed) Then

                Result = Result + " "

                Result = Result + CurrentProperty.ToStringCollapsed
                
            Else
            
                Result = Result + CurrentProperty.ToStringExpanded + vbCrLf
                
            End If
        Next
        
        ToStringProperties = Result
    
    End Function
    
    Private Function GetProperties() As String
    
        Dim Result As String
        Dim CurrentProperty As XMLUtilTagValue
        
        Result = sysAdlStringConstantsEmpty
        
        For Each CurrentProperty In pProperties
            
            Result = Result + " "
        
            Result = Result + CurrentProperty.ToStringCollapsed
        Next
        
        GetProperties = Result
    
    End Function
    
    Public Property Get OpenTag() As String
    
        Dim Result As String
        
        If pIsCollapsed Then
            
            Result = sysAdlTagOpen + pName + ToStringProperties + sysAdlTagEnd
            
        Else
        
            Result = sysAdlTagOpen + pName + GetProperties + sysAdlTagClose
            
        End If
        
        OpenTag = Result
    
    End Property
    
    Public Property Get CloseTag() As String
    
        Dim Result As String
        
        If pIsCollapsed Then
        
            Result = sysAdlTagClose
            
        Else
        
            Result = sysAdlTagEnding + pName + sysAdlTagClose
        
        End If
            
       CloseTag = Result
        
    End Property
