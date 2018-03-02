VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ribbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' clsRibbon / Ribbon.cls
' Copyright (c) Microsoft Corporation. All rights reserved.
'
' Summary:
' This class demonstrates how to create custom UI and its associated callbacks.
'Modified by David J Parker, March 2011


Implements IRibbonExtensibility

Private Enum BackstageColumn
    One = 1
    Two = 2
    Three = 3
End Enum

Public Sub Class_Initialize()

' Class_Initialize
'
' Abstract - This constructor is intentionally left blank.

    ' No initialization is required.
End Sub

Public Function IRibbonExtensibility_GetCustomUI(ByVal RibbonID As String) As String

' IRibbonExtensibility_GetCustomUI
'
' Abstract - This method is a callback that is called by Visio to obtain the
' custom UI definition.
'
' Parameters
' RibbonID         A string identifying the Visio UI
'
' Return Value     A string representing the XML defining the custom UI

    IRibbonExtensibility_GetCustomUI = getRibbonXML(True, True, True, True)
End Function

Public Sub OnRibbonLoad(ByVal ribbonUI As IRibbonUI)

' OnRibbonLoad
'
' Abstract - This method is a callback specified in the onLoad attribute of the
' customUI element in the custom UI XML file.  It is called by Visio when the
' custom UI is first loaded.
'
' Parameters
' ribbonUI     A reference to the object representing the custom UI loaded by Visio

    ' Do something with the newly-constructed Ribbon, such as capture
    ' a local reference to it for later use.
End Sub

Public Sub OnAction(ByVal control As IRibbonControl)

' OnAction
'
' Abstract - This method is a callback specified in the custom UI XML file.
' It is called by Visio when the associated button defined in the XML is pressed.
'
' Parameters
' control     The Ribbon UI control that was activated
    
    
    Dim SelectedId As String
    Dim RibbonSelected As RibbonItem
    
    SelectedId = control.Id
    
    Set RibbonSelected = FactoryRibbon.GetRibbonItemById(SelectedId)
    
    RibbonSelected.ExecuteCommand

End Sub

Public Sub CommandOnAction(ByVal control As IRibbonControl, _
    ByVal CancelDefault As Boolean)

' CommandOnAction
'
' Abstract - This method is a callback specified in the custom UI XML file.
' It is called by Visio when the associated repurposed Ribbon control is used.
'
' Parameters
' control           The Ribbon UI control that was activated
' cancelDefault     If true, call the built-in command after the custom code
' is complete

    ' Do custom work when the Copy button is pressed, for example.
    Dim SelectedId As String
    Dim RibbonSelected As RibbonItem
    
    SelectedId = control.Id
    
    Set RibbonSelected = FactoryRibbon.GetRibbonItemById(SelectedId)
    
    RibbonSelected.ExecuteCommand

    CancelDefault = False
    
End Sub

Private Function getRibbonXML(ByVal includeCommands As Boolean, _
    ByVal includeRibbonTab As Boolean, ByVal includeBackstage As Boolean, _
    ByVal includeContextMenus As Boolean) As String

' GetRibbonXML
'
' Abstract - This method is a helper that reads the content of a text file
' resource embedded in the project.
'
' Return Value     A string representing the XML defining the custom UI

    Dim strGetRibbonXML As String
    Dim strGetRibbonXML1 As String
    Dim strGetRibbonXML2 As String
    Dim strGetRibbonXML3 As String
    Dim strGetRibbonXML4 As String
    Dim CurrentRibbon As RibbonItem
    
    strGetRibbonXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
        "<customUI onLoad=""OnRibbonLoad"" " & _
        "xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf

    If includeCommands Then
        'Open the Commands element
        strGetRibbonXML1 = getCommandsBegin
        'Add Command actions as required
        'You can disable commands
        strGetRibbonXML1 = strGetRibbonXML1 & getDisableCommand("Bold")
        'You can re-purpose commands
        strGetRibbonXML1 = strGetRibbonXML1 & getRedirectCommand("Copy")
        'Close the Commands element
        strGetRibbonXML1 = strGetRibbonXML1 & getCommandsEnd
        
        strGetRibbonXML = strGetRibbonXML & strGetRibbonXML1
    End If
    
    If includeRibbonTab Then
        'Open the Ribbon element
        strGetRibbonXML2 = getRibbonBegin
        'Open the Tab element
        strGetRibbonXML2 = strGetRibbonXML2 & getTabBegin(False, "tab1", "Sys-ADL")
        'Open the Group element
        strGetRibbonXML2 = strGetRibbonXML2 & getGroupBegin(False, "group1", "SysADL")
        'Add custom buttons as required
        
        For Each CurrentRibbon In FactoryRibbon.GetRibbonItems
        
            strGetRibbonXML2 = strGetRibbonXML2 & getButton( _
                CurrentRibbon.Id, _
                CurrentRibbon.Label, _
                CurrentRibbon.Tip, _
                CurrentRibbon.ImageMSO, _
                True)
            'Open the split button element
            'strGetRibbonXML2 = strGetRibbonXML2 & getSplitButtonBegin("customSplit1")
            'Open the menu element
            'strGetRibbonXML2 = strGetRibbonXML2 & getMenuBegin("customMenu1")
            'Add custom buttons as required
            'strGetRibbonXML2 = strGetRibbonXML2 & getButton( _
            '    "customMacro4", "Macro 4", "This is macro 4", "VisioDiagramGallery", False)
    
            'Close the menu element
            'strGetRibbonXML2 = strGetRibbonXML2 & getMenuEnd
            'Close the split button element
            'strGetRibbonXML2 = strGetRibbonXML2 & getSplitButtonEnd
            'Close the Group element
            strGetRibbonXML2 = strGetRibbonXML2 & getGroupEnd
            'Close the Tab element
            strGetRibbonXML2 = strGetRibbonXML2 & getTabEnd
            'Close the Ribbon element
            strGetRibbonXML2 = strGetRibbonXML2 & getRibbonEnd
            
            strGetRibbonXML = strGetRibbonXML & strGetRibbonXML2
        Next
            
    End If
    
    If includeBackstage Then
        'http://msdn.microsoft.com/en-us/library/ee815851(office.14).aspx
        'Open the backstage element
        'strGetRibbonXML3 = getBackstageBegin
        'Open the tab element
        'strGetRibbonXML3 = strGetRibbonXML3 & getTabBegin(False, "tab2", "Custom Tab")
        'Open the column element
        'strGetRibbonXML3 = strGetRibbonXML3 & getBackstageColumnBegin(One)
        'strGetRibbonXML3 = strGetRibbonXML3 & vbTab & getGroupBegin(False, "group2", "Custom Group 2")
        'strGetRibbonXML3 = strGetRibbonXML3 & getBackstagePrimaryItemBegin
        'Add custom buttons as required
        'strGetRibbonXML3 = strGetRibbonXML3 & vbTab & vbTab & getButton( _
        '    "customBMacro1", "BS Macro 1", "This is bs macro 1", "GroupSynchronizeWithSite", False)
        'Close the primaryitem element
        'strGetRibbonXML3 = strGetRibbonXML3 & getBackstagePrimaryItemEnd
        'Close the group element
        'strGetRibbonXML3 = strGetRibbonXML3 & vbTab & getGroupEnd
        'Close the column element
        'strGetRibbonXML3 = strGetRibbonXML3 & getBackstageColumnEnd(One)
        
        'Open the column element
        'strGetRibbonXML3 = strGetRibbonXML3 & getBackstageColumnBegin(Two)
        'Open the group element
        'strGetRibbonXML3 = strGetRibbonXML3 & vbTab & getGroupBegin(False, "group3", "Custom Group 3")
        'Open the topitems element
        'strGetRibbonXML3 = strGetRibbonXML3 & getBackstageTopItemsBegin
        'Add custom buttons as required
        'strGetRibbonXML3 = strGetRibbonXML3 & vbTab & vbTab & getButton( _
        '    "customBMacro2", "BS Macro 2", "This is bsmacro 2", "GroupViewsInfoPath", False)
        'strGetRibbonXML3 = strGetRibbonXML3 & vbTab & vbTab & getButton( _
        '    "customBMacro3", "BS Macro 3", "This is bsmacro 3", "PostReply", False)
        'Close the topitems element
        'strGetRibbonXML3 = strGetRibbonXML3 & getBackstageTopItemsEnd
        'Close the group element
        'strGetRibbonXML3 = strGetRibbonXML3 & vbTab & getGroupEnd
        'Close the column element
        'strGetRibbonXML3 = strGetRibbonXML3 & getBackstageColumnEnd(Two)
        'Close the tab element
        'strGetRibbonXML3 = strGetRibbonXML3 & getTabEnd
        'Close the backstage element
        'strGetRibbonXML3 = strGetRibbonXML3 & getBackstageEnd
        
        'strGetRibbonXML = strGetRibbonXML & strGetRibbonXML3
    End If
    
    If includeContextMenus Then
        'Open the context menus group element
        'strGetRibbonXML4 = getContextMenusBegin
        'Open the context menu element
        'See the contextmenus worksheet in C:\Office 2010 Developer Resources\Documents\Office2010ControlIDs\VisioControls.xlsx
        'strGetRibbonXML4 = strGetRibbonXML4 & getContextMenuBegin("ContextMenuShape")
        'Include any menu buttons required
        'strGetRibbonXML4 = strGetRibbonXML4 & getButton( _
         '   "customContextMacro1", "My Button", "This is my context menu macro", _
         '   "MindMapChangeTopic", False)
        'Close the context menu element
        'strGetRibbonXML4 = strGetRibbonXML4 & getContextMenuEnd
        'Close the context menus group element
        'strGetRibbonXML4 = strGetRibbonXML4 & getContextMenusEnd
        
        'strGetRibbonXML = strGetRibbonXML & strGetRibbonXML4
    End If

    strGetRibbonXML = strGetRibbonXML & "</customUI>"

    getRibbonXML = strGetRibbonXML
End Function

Private Function getRibbonBegin() As String
    'This section defines custom Ribbon controls
    Dim txt As String
    txt = vbTab & "<ribbon>" & vbCrLf & vbTab & vbTab & "<tabs>" & vbCrLf
    getRibbonBegin = txt
End Function

Private Function getTabBegin(ByVal isMsoTab As Boolean, ByVal tabId As String, _
        ByVal tabLabel As String) As String
    'This section defines custom Ribbon controls
    Dim txt As String
    txt = vbTab & vbTab & vbTab & "<tab "
    If isMsoTab Then
        txt = txt & "idMso="
    Else
        txt = txt & "id="
    End If
    txt = txt & """" & tabId & """ label=""" & tabLabel & """>"
    getTabBegin = txt & vbCrLf
End Function

Private Function getGroupBegin(ByVal isMsoGroup As Boolean, ByVal groupId As String, _
        ByVal groupLabel As String) As String
    Dim txt As String
    txt = vbTab & vbTab & vbTab & vbTab & "<group "
    If isMsoGroup Then
        txt = txt & "idMso="
    Else
        txt = txt & "id="
    End If
    txt = txt & """" & groupId & """ label=""" & groupLabel & """>"
    getGroupBegin = txt & vbCrLf
End Function

Private Function getGroupEnd() As String
    getGroupEnd = vbTab & vbTab & vbTab & vbTab & "</group>" & vbCrLf
End Function

Private Function getRibbonEnd() As String
    getRibbonEnd = vbTab & vbTab & "</tabs>" & vbCrLf & _
        vbTab & "</ribbon>" & vbCrLf
End Function

Private Function getTabEnd() As String
    getTabEnd = vbTab & vbTab & vbTab & "</tab>" & vbCrLf
End Function

Private Function getButton(ByVal buttonId As String, ByVal buttonLabel As String, _
    ByVal buttonSuperTip As String, ByVal buttonImageMso As String, ByVal IsLarge As Boolean) As String
    Dim txt As String
    txt = vbTab & vbTab & vbTab & vbTab & vbTab & "<button id=""" & buttonId & """ label=""" & buttonLabel & """ "
    txt = txt & "supertip=""" & buttonSuperTip & """ "
    If IsLarge Then
        txt = txt & "size=""large"" "
    End If
    If Len(buttonImageMso) > 0 Then
        txt = txt & "imageMso=""" & buttonImageMso & """ "
    End If
    txt = txt & "onAction=""OnAction""/>" & vbCrLf
    getButton = txt
End Function

Private Function getCommandsBegin() As String
'This section repurposes built-in commands
    getCommandsBegin = vbTab & "<commands>" & vbCrLf
End Function

Private Function getCommandsEnd() As String
    getCommandsEnd = vbTab & "</commands>" & vbCrLf
End Function

Private Function getRedirectCommand(ByVal commandIdMso As String) As String
'Redirects the built-in given command to custom callback
    getRedirectCommand = vbTab & vbTab & "<command idMso=""" & commandIdMso & """ onAction=""CommandOnAction""/>" & vbCrLf
End Function

Private Function getDisableCommand(ByVal commandIdMso As String) As String
'Diables the given command command
    getDisableCommand = vbTab & vbTab & "<command idMso=""" & commandIdMso & """ enabled=""false""/>" & vbCrLf
End Function

Private Function getBackstageBegin() As String
'This section starts the backstage
    getBackstageBegin = vbTab & "<backstage>" & vbCrLf
End Function

Private Function getBackstageEnd() As String
    getBackstageEnd = vbTab & "</backstage>" & vbCrLf
End Function

Private Function getContextMenusBegin() As String
'This section starts the contextMenus
    getContextMenusBegin = vbTab & "<contextMenus>" & vbCrLf
End Function

Private Function getContextMenusEnd() As String

    getContextMenusEnd = vbTab & "</contextMenus>" & vbCrLf
End Function

Private Function getContextMenuBegin(ByVal commandIdMso As String) As String
'This section starts the contextMenu
'Eg ContextMenuShape
    getContextMenuBegin = vbTab & vbTab & "<contextMenu idMso=""" & commandIdMso & """>" & vbCrLf
End Function

Private Function getContextMenuEnd() As String

    getContextMenuEnd = vbTab & vbTab & "</contextMenu>" & vbCrLf
End Function

Private Function getSplitButtonBegin(ByVal buttonId As String) As String
'This section starts a split button
    getSplitButtonBegin = vbTab & vbTab & vbTab & vbTab & vbTab & "<splitButton id=""" & buttonId & """>" & vbCrLf
End Function

Private Function getSplitButtonEnd() As String
    getSplitButtonEnd = vbTab & vbTab & vbTab & vbTab & vbTab & "</splitButton>" & vbCrLf
End Function

Private Function getMenuBegin(ByVal menuId As String) As String
'This section starts a menu
    getMenuBegin = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<menu id=""" & menuId & """>" & vbCrLf
End Function

Private Function getMenuEnd() As String
    getMenuEnd = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "</menu>" & vbCrLf
End Function

Private Function getBackstageColumnBegin(ByVal column As BackstageColumn) As String
Dim strColumn As String
    Select Case column
        Case BackstageColumn.One
            strColumn = "<firstColumn>"
        Case BackstageColumn.Two
            strColumn = "<secondColumn>"
        Case BackstageColumn.Three
            strColumn = "<thirdColumn>"
    End Select
    getBackstageColumnBegin = vbTab & vbTab & vbTab & vbTab & strColumn & vbCrLf
End Function

Private Function getBackstageColumnEnd(ByVal column As BackstageColumn) As String
Dim strColumn As String
    Select Case column
        Case BackstageColumn.One
            strColumn = "</firstColumn>"
        Case BackstageColumn.Two
            strColumn = "</secondColumn>"
        Case BackstageColumn.Three
            strColumn = "</thirdColumn>"
    End Select
    getBackstageColumnEnd = vbTab & vbTab & vbTab & vbTab & strColumn & vbCrLf
End Function

Private Function getBackstagePrimaryItemBegin() As String

    getBackstagePrimaryItemBegin = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<primaryItem>" & vbCrLf
End Function

Private Function getBackstagePrimaryItemEnd() As String

    getBackstagePrimaryItemEnd = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "</primaryItem>" & vbCrLf
End Function

Private Function getBackstageTopItemsBegin() As String
    getBackstageTopItemsBegin = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<topItems>" & vbCrLf
End Function

Private Function getBackstageTopItemsEnd() As String
    getBackstageTopItemsEnd = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "</topItems>" & vbCrLf
End Function
