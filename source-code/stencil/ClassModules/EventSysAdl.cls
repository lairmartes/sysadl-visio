VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventSysAdl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    Private pEventType As Integer
    
Public Sub Init(ByVal EventFiredType As Integer)
    pEventType = EventFiredType
End Sub

Public Property Get EventType() As Integer
    EventType = pEventType
End Property
