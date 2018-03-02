Attribute VB_Name = "ElementServicePersistenceXMLOut"

Option Explicit

    'object file for exporting
    Private Const TAG_COLLAPSED = True
    Dim CurrentFileExportingStream As Stream


Public Sub ProcessElementSaving(ByVal element As SysAdlElement)

    Dim FolderName As String
    Dim fileName As String
    
    FolderName = CreateFolderForElement(element.namespace)
    
    fileName = GUIServices.GetFullFileNameForElement(element.namespace, element.Id)
    
    CreateElementDataFile fileName, element
    
    
End Sub

Private Function CreateFolderForElement(ByVal aNamespace As String) As String

    Dim FolderName As String
    Dim Result As String

    FolderName = GUIServices.GetPathFromNamespace(aNamespace)
    
    Result = CreateFolder(FolderName)
    
    CreateFolderForElement = Result

End Function

Private Function CreateFolder(ByVal aFolderName As String) As String

    Dim Result As String
    Dim ParentFolderName As String
    Dim IsFolderExists As Boolean
    Dim IsParentFolderExists As Boolean
    Dim IsBasePathReached As Boolean
    Dim basePath As String
    
    Result = aFolderName
    
    basePath = FactoryConfigProperty.GetProperty(sysAdlConfigPropertyBasePath)

    IsBasePathReached = (aFolderName = basePath)
    
    If Not IsBasePathReached Then
        
        IsFolderExists = Len(Dir(aFolderName, vbDirectory) & "") > 0
        
        If Not IsFolderExists Then
        
            ParentFolderName = CalculateSubfolderName(aFolderName)
            
            IsParentFolderExists = Len(Dir(ParentFolderName, vbDirectory) & "") > 0
            
            If Not IsParentFolderExists Then
            
                CreateFolder (ParentFolderName)
                
                MkDir aFolderName
                
            Else
            
                MkDir aFolderName
                
            End If
            
        End If
        
    End If
    
    CreateFolder = Result

End Function

Private Function CalculateSubfolderName(ByVal aFolderName As String)

    Dim Result As String
    Dim FolderList As Variant
    Dim I As Integer
    
    Result = sysAdlStringConstantsEmpty

    FolderList = Split(aFolderName, sysAdlStringConstantsWindowsPathSeparator)
        
    For I = 0 To UBound(FolderList) - 1
            
        If I > 0 Then
            
            Result = Result + sysAdlStringConstantsWindowsPathSeparator + FolderList(I)
                    
        Else
                
            Result = FolderList(I)
                
        End If
                
    Next I
    
    CalculateSubfolderName = Result

End Function




'publish the result of analysis in XML format (using sysel extension)
'it publishes a file in a path created based on element's namepace and element's id
Private Sub CreateElementDataFile(ByVal fileName As String, ByVal element As SysAdlElement)

    ' declare tag <sysadldiagram>
    Dim TagSectionSysAdlDiagram As New XMLUtilTag
    Dim TagPropertyName As New XMLUtilTagValue

    
    'configure file object
    
    Set CurrentFileExportingStream = New Stream
    CurrentFileExportingStream.Open
    CurrentFileExportingStream.Position = 0
    CurrentFileExportingStream.CharSet = "UTF-8"
    
    
    'publish element data filled in diagram
    PublishElementData element
    
    PublishSkipLine
    
    'save file
    CurrentFileExportingStream.SaveToFile fileName, adSaveCreateOverWrite
    CurrentFileExportingStream.Close
    
    'Close #pFileExportingNumber
    

End Sub

'publish element data

Private Sub PublishElementData(ByVal anElement As SysAdlElement)

    ' declare tag data for element head
    Dim TagSectionSysAdlElement As New XMLUtilTag
    Dim TagPropertySysAdlType As New XMLUtilTagValue
    Dim TagPropertyStereotype As New XMLUtilTagValue
    Dim TagPropertyNamespace As New XMLUtilTagValue
    Dim TagPropertyId As New XMLUtilTagValue
    Dim TagPropertyUrlInfo As New XMLUtilTagValue
    Dim StereotypePublished As String
    
    ' declare tag data for attributes
    Dim TagSectionAttributes As XMLUtilTag
    Dim TagPropertyAttributeName As XMLUtilTagValue
    Dim TagPropertyAttributeValue As XMLUtilTagValue
    
    ' declare tag data for shape id
    Dim TagSectionShapeId As XMLUtilTag
    Dim TagPropertyShapeIdValue As XMLUtilTagValue
    
    'fields from element
    Dim ElementFieldList As Collection
    Dim CurrentElementAttribute As UtilSysAdlItemString
    
    'shapes from element
    Dim ElementShapeViewerList As Collection
    Dim CurrentElementShapeViewer As ShapeViewer
    
    'initiate <sysadl-element> tag
    TagSectionSysAdlElement.Init sysAdlTagSectionSysAdlElement
    
    'initializae stereotype value
    StereotypePublished = anElement.Stereotype
    'change to blank if equals <NO_STEROTYPE>
    If StereotypePublished = sysAdlNoStereotype Then _
        StereotypePublished = sysAdlStringConstantsEmpty
    
    ' add data for type, stereotype, namespace, id and UrlInfo
    TagPropertySysAdlType.Init sysAdlTagPropertyType, anElement.sysadlType
    TagPropertyStereotype.Init sysAdlTagPropertyStereotype, StereotypePublished
    TagPropertyNamespace.Init sysAdlTagPropertyNamespace, anElement.namespace
    TagPropertyId.Init sysAdlTagPropertyId, anElement.Id
    TagPropertyUrlInfo.Init sysAdlTagPropertyUrlInfo, GUIServices.PrepareFieldForXML(anElement.URLInfo)
        
    ' add properties
    TagSectionSysAdlElement.AddProperty TagPropertySysAdlType
    TagSectionSysAdlElement.AddProperty TagPropertyStereotype
    TagSectionSysAdlElement.AddProperty TagPropertyNamespace
    TagSectionSysAdlElement.AddProperty TagPropertyId
    TagSectionSysAdlElement.AddProperty TagPropertyUrlInfo
        
    'public tag <element-[type]> (or only <element>
    Publish TagSectionSysAdlElement.OpenTag
    
    Set ElementFieldList = anElement.Fields.GetUtilItemCollection
    
    For Each CurrentElementAttribute In ElementFieldList
    
        Set TagSectionAttributes = New XMLUtilTag
        Set TagPropertyAttributeName = New XMLUtilTagValue
        Set TagPropertyAttributeValue = New XMLUtilTagValue
        
        TagSectionAttributes.Init sysAdlTagSectionElementAttribute, TAG_COLLAPSED
        TagPropertyAttributeName.Init sysAdlTagPropertyAttributeName, CurrentElementAttribute.Key
        TagPropertyAttributeValue.Init sysAdlTagPropertyAttributeValule, GUIServices.PrepareFieldForXML(CurrentElementAttribute.Item)
        
        TagSectionAttributes.AddProperty TagPropertyAttributeName
        TagSectionAttributes.AddProperty TagPropertyAttributeValue
        
        Publish TagSectionAttributes.OpenTag
    
    Next
        
    
    Publish TagSectionSysAdlElement.CloseTag
    PublishSkipLine


End Sub



'print line in file
Private Sub Publish(ByVal aLine As String)

    CurrentFileExportingStream.WriteText aLine

End Sub

Private Sub PublishSkipLine()

    CurrentFileExportingStream.SkipLine
    
End Sub

