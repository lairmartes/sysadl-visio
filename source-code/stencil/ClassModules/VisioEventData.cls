VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VisioEventData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pMouseX As Double
Private pMouseY As Double
Private pMouseButtonPressed As Long
Private pMouseKeyButtonState As Long
Private pMouseCancelDefault As Boolean


Public Sub InitMouseEvent(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)

     pMouseX = x
     pMouseY = y
     pMouseButtonPressed = Button
     pMouseKeyButtonState = KeyButtonState
     pMouseCancelDefault = CancelDefault

End Sub

Public Property Get MouseX() As Double

    MouseX = pMouseX
    
End Property

Public Property Get MouseY() As Double

    MouseY = pMouseY
    
End Property

Public Property Get MouseButtonPressed() As Double

    MouseButtonPressed = pMouseButtonPressed
    
End Property

Public Property Get MouseKeyButtonState() As Double

    MouseKeyButtonState = pMouseKeyButtonState
    
End Property

Public Property Get MouseCancelDefault() As Double

    MouseCancelDefault = pMouseCancelDefault
    
End Property

