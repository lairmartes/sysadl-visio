Attribute VB_Name = "FactoryEventBroadcast"
Option Explicit

Private pVisioEventBroadCast As VisioEventBroadcast
Private pVisioEventBroadcastCreated As Boolean


Public Function CreateEventBroadcast() As VisioEventBroadcast

    If Not pVisioEventBroadcastCreated Then
    
        Set pVisioEventBroadCast = New VisioEventBroadcast
        
        pVisioEventBroadcastCreated = True
        
    End If
    
    Set CreateEventBroadcast = pVisioEventBroadCast
    

End Function
