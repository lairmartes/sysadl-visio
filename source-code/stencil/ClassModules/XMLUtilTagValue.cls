VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLUtilTagValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pProperty As String
    Private pValue As String
    
    Public Sub Init(ByVal aProperty As String, ByVal aValue As String)
    
        pProperty = aProperty
        pValue = aValue
    
    End Sub

    
    Public Function ToStringCollapsed() As String
        
        Dim Result As String
    
        Result = pProperty + "=" + """" + GUIServices.PrepareFieldForXML(pValue) + """"
        
        ToStringCollapsed = Result
    
    End Function
    
    Public Function ToStringExpanded() As String
    
        Dim Result As String
        
        Dim OpeningTag As String
        Dim ClosingTag As String
        
        OpeningTag = sysAdlTagOpen + pProperty + sysAdlTagClose
        ClosingTag = sysAdlTagEnding + pProperty + sysAdlTagClose
        
        Result = OpeningTag + GUIServices.PrepareFieldForXML(pValue) + ClosingTag
        
        ToStringExpanded = Result
    
    End Function
