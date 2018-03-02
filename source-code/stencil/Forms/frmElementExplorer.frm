VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmElementExplorer 
   Caption         =   "Sys-ADL Element Explorer"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7500
   OleObjectBlob   =   "frmElementExplorer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmElementExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private pElementExplorerModel As FrmModelElementExplorer
    Private pDragDropCommand As CommandDragDropElementExplorer


Private Sub UserForm_Initialize()

    Set pElementExplorerModel = New FrmModelElementExplorer
    
    InitializeExplorerTreeNamespaces
    
    InitializeExplorerTreeIds
    
    Set pDragDropCommand = New CommandDragDropElementExplorer
    
    FactoryEventBroadcast.CreateEventBroadcast().AddEventListener pDragDropCommand


End Sub


Private Sub InitializeExplorerTreeNamespaces()

    Dim namespaceList As Collection
    Dim namespacesQttyFound As Integer
    Dim currentNamespace As String
    Dim I As Integer
    
    Set namespaceList = pElementExplorerModel.namespaceList

    namespacesQttyFound = namespaceList.Count
    
    For I = 1 To namespacesQttyFound
    
        currentNamespace = namespaceList(I)
           
        ElementExplorerTree.Nodes.Add , , currentNamespace, currentNamespace

    Next I

End Sub

Private Sub InitializeExplorerTreeIds()

    Dim elementKeyList As Collection
    Dim elementKeyQttyFound As Integer
    Dim currentNamespace As String
    Dim currentId As String
    Dim currentElementKey As String
    Dim I As Integer
    
    Set elementKeyList = pElementExplorerModel.elementKeyList

    elementKeyQttyFound = elementKeyList.Count
    
    For I = 1 To elementKeyQttyFound
    
        currentElementKey = elementKeyList(I)
    
        currentNamespace = GUIServices.GetNamespaceFromString(currentElementKey)
        currentId = GUIServices.GetIdFromString(currentElementKey)
        
           
        ElementExplorerTree.Nodes.Add currentNamespace, tvwChild, currentElementKey, currentId

    Next I

End Sub

Private Sub UserForm_Resize()

    ElementExplorerTree.Width = ElementExplorerTree.Parent.Width - 10
    ElementExplorerTree.Height = ElementExplorerTree.Parent.Height - 10

End Sub


Private Sub ElementExplorerTree_OLECompleteDrag(effect As Long)

    
     pDragDropCommand.ICommand_Execute

End Sub




Private Sub ElementExplorerTree_OLESetData(data As mscomctllib.DataObject, dataformat As Integer)

    
    data.SetData (pDragDropCommand.GetShapeAdded)

End Sub

Private Sub ElementExplorerTree_OLEStartDrag(data As mscomctllib.DataObject, allowedEffects As Long)

    Dim selectedElementKey As String
    
    selectedElementKey = ElementExplorerTree.SelectedItem.Key
    
    pDragDropCommand.Init selectedElementKey
    
    If pDragDropCommand.GetShapeAdded Is Nothing Then
    
        allowedEffects = 0 'vbDropEffectNone
        
    Else
    
        allowedEffects = 1 'vbDropEffectCopy
        
    End If

End Sub


