Attribute VB_Name = "FactoryEvent"
Option Explicit
    
    ' Events for Changing
    Public Const sysAdlEventChangedSysAdlElement As Integer = 0
    Public Const sysAdlEventElementRecovered As Integer = 1
    Public Const sysAdlEventChangedCellValue As Integer = 3
    Public Const sysAdlEventChangedStereotype As Integer = 5
    Public Const sysAdlEventChangedURLInfo As Integer = 6
    Public Const sysAdlEventInvalidFieldFound As Integer = 7
    Public Const sysAdlEventInvalidFieldCorrected As Integer = 8
    Public Const sysAdlEventDocumentOpened As Integer = 9
    Public Const sysAdlEventElementPersisted As Integer = 10
    Public Const sysAdlEventFieldsUpdated As Integer = 11
    Public Const sysAdlEventKeyUsedOtherType As Integer = 12
    Public Const sysAdlEventDocumentSavedAs As Integer = 13

    
    'Events for Creation
    Public Const sysAdlEventCreatedSysAdlElement As Integer = 4
    

Public Function CreateEvent(ByVal EventFiredType As Integer) As EventSysAdl

    Dim Result As New EventSysAdl
    
    Result.Init EventFiredType

    Set CreateEvent = Result

End Function
