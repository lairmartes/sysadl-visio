VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VisioEventBroadcast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private eventListeners As Collection
Private eventListenersListCreated As Boolean



Public Sub AddEventListener(ByVal eventListener As IVisioEventLister)

    
    
    GetEventListeners.Add eventListener
    

End Sub

Private Function GetEventListeners() As Collection

    If Not eventListenersListCreated Then
    
        Set eventListeners = New Collection
        
        eventListenersListCreated = True
        
    End If
    
    Set GetEventListeners = eventListeners

End Function

Public Sub DocumentMouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
 
    Dim currentListener As IVisioEventLister
    Dim mouseEventData As VisioEventData
    
    Set mouseEventData = CreateMouseEventData(Button, KeyButtonState, x, y, CancelDefault)
    
    For Each currentListener In GetEventListeners
    
        currentListener.MouseDown mouseEventData
         
    Next
    
 
End Sub
 
Public Sub DocumentMouseMove(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
 
    Dim currentListener As IVisioEventLister
    Dim mouseEventData As VisioEventData
    
    Set mouseEventData = CreateMouseEventData(Button, KeyButtonState, x, y, CancelDefault)
    
    For Each currentListener In GetEventListeners
    
        currentListener.MouseMove mouseEventData
         
    Next
 
End Sub
 
Public Sub DocumentMouseUp(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
 
    Dim currentListener As IVisioEventLister
    Dim mouseEventData As VisioEventData
    
    Set mouseEventData = CreateMouseEventData(Button, KeyButtonState, x, y, CancelDefault)
    
    For Each currentListener In GetEventListeners
    
        currentListener.MouseUp mouseEventData
         
    Next
 
 
End Sub

Private Function CreateMouseEventData(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean) As VisioEventData


    Dim Result As VisioEventData
    
    Set Result = New VisioEventData
    
    Result.InitMouseEvent Button, KeyButtonState, x, y, CancelDefault
    
    Set CreateMouseEventData = Result
    
    
End Function
 

