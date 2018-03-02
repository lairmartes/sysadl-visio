Attribute VB_Name = "FactoryRibbon"
Option Explicit

    Private Const sysadlRibbonCommandElementExplorer = "SysADLElementExplorer"

    Private RibbonItemList As Collection
    
    ' A reference to the sample class that creates and manages the custom UI.
    Private msoCustomRibbon As Ribbon
    

Public Sub CustomUIStart(ByVal vsoTargetDocument As Visio.Document)

' Abstract - This method loads custom UI from an XML file and associates it
' with the document object passed in.
'
' Parameters
' vsoTargetDocument     An open document in a running Visio application

    Dim vsoApplication As Visio.Application
    
    Set vsoApplication = Visio.Application
    Set msoCustomRibbon = New Ribbon

    ' Passing in null rather than targetDocument would make the custom
    ' UI available for all documents.
    vsoApplication.RegisterRibbonX msoCustomRibbon, _
        vsoTargetDocument, _
        Visio.VisRibbonXModes.visRXModeDrawing, _
        "Sys-ADL Tools"
    'NB The FriendlyName text will appear at the bottom of SuperTips
    
    Init
    
End Sub

Public Sub CustomUIStop(ByVal vsoTargetDocument As Visio.Document)

' Abstract - This method removes custom UI from a document.
'
' Parameters
' vsoTargetDocument     An open document in a running Visio application that
' has custom UI associated with it

    Dim vsoApplication As Visio.Application
    On Error Resume Next
    Set vsoApplication = Visio.Application
    vsoApplication.UnregisterRibbonX msoCustomRibbon, _
        vsoTargetDocument
End Sub

Private Sub Init()
    
    If RibbonItemList Is Nothing Then
    
        Dim RibbonElementExplorer As RibbonItem
        
        Set RibbonElementExplorer = CreateElementExplorerButton
            
        Set RibbonItemList = New Collection
        
        RibbonItemList.Add RibbonElementExplorer
        
    End If

End Sub

Public Function GetRibbonItemById(ByVal anId As String) As RibbonItem

    Dim Result As RibbonItem

    For Each Result In RibbonItemList
    
        If Result.Id = anId Then Exit For
    
    Next
    
    Set GetRibbonItemById = Result

End Function

Public Function GetRibbonItems() As Collection
    
    If RibbonItemList Is Nothing Then Init
    
    Set GetRibbonItems = RibbonItemList

End Function


Private Function CreateElementExplorerButton() As RibbonItem

    Dim ShowExplorerCommand As ICommand
    Dim Result As RibbonItem
    
    Set Result = New RibbonItem
    
    Set ShowExplorerCommand = New CommandShowElementExplorer
    
    Result.Init sysadlRibbonCommandElementExplorer, _
                "Element Explorer", _
                "Open window to show element explorer", _
                "OrganizationChartLayoutRightHanging", _
                True, _
                ShowExplorerCommand
                
                
    Set CreateElementExplorerButton = Result

End Function
