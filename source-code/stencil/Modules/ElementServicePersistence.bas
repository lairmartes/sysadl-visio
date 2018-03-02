Attribute VB_Name = "ElementServicePersistence"




Public Function OpenPersistedElement(ByVal element As SysAdlElement) As SysAdlElement

    Set Result = ElementServicePersistenceXMLIn.OpenSysADLElement(element)
    Set OpenPersistedElement = Result

End Function


Public Sub ProcessElementDelete(ByVal element As SysAdlElement)

    '' do nothing with files...
    
End Sub


Public Sub ProcessElementSaving(ByVal element As SysAdlElement)

    ElementServicePersistenceXMLOut.ProcessElementSaving element
    
End Sub


Public Function IsElementExists(ByVal element As SysAdlElement) As Boolean

    Dim Result As Boolean
    
    Result = ElementServicePersistenceXMLIn.IsElementExists(element)
        
    IsElementExists = Result

End Function

Public Function IsElementExistsWithType(ByVal element As SysAdlElement) As Boolean

    Dim Result As Boolean
    
    Result = ElementServicePersistenceXMLIn.IsElementExistsWithType(element)
        
    IsElementExistsWithType = Result

End Function




