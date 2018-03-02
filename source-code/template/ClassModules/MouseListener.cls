VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MouseListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Dim WithEvents vsoWindow As Visio.Window
Attribute vsoWindow.VB_VarHelpID = -1
 
Private Sub Class_Initialize()
 
 Set vsoWindow = ActiveWindow
 
End Sub
 
Private Sub Class_Terminate()
 
 Set vsoWindow = Nothing
 
End Sub
 
Private Sub vsoWindow_MouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
 
    FactoryEventBroadcast.CreateEventBroadcast().DocumentMouseDown Button, KeyButtonState, x, y, CancelDefault
 
End Sub
 
Private Sub vsoWindow_MouseMove(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
 
    FactoryEventBroadcast.CreateEventBroadcast().DocumentMouseMove Button, KeyButtonState, x, y, CancelDefault
 
End Sub
 
Private Sub vsoWindow_MouseUp(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
 
    FactoryEventBroadcast.CreateEventBroadcast().DocumentMouseUp Button, KeyButtonState, x, y, CancelDefault
 
End Sub
 
