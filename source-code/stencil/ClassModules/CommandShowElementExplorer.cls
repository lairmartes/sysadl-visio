VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandShowElementExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICommand



Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000


Public Sub ICommand_Execute()

    Dim vsoWindow  As Visio.Window
    Dim frmElementExplorer  As UserForm
    Dim lngFormHandle  As Long
    Dim windowTitle As String
    
    windowTitle = "Sys-ADL Element Explorer" 'Sys-ADL Element Explorer
    
    Set frmElementExplorer = New frmElementExplorer
    
    Set vsoWindow = ActiveWindow.Windows.Add(windowTitle, visWSVisible + visWSDockedBottom, visAnchorBarAddon, , , 500, 210)
    'Set vsoWindow = ActiveWindow.Windows.Add(tituloJanela, visWSVisible, visAnchorBarAddon, , , 300, 210)

    lngFormHandle = FindWindow(vbNullString, windowTitle)
    
    SetWindowLong lngFormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE
    SetParent lngFormHandle, vsoWindow.WindowHandle32

End Sub
