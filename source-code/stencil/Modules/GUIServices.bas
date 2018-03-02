Attribute VB_Name = "GUIServices"
Option Explicit

Private Const sysAdlGuiServicesLogIsEnabled = "True"
Private sysADLVisioStencil As Visio.Document
Private sysADLVisioStencilAlreadyInitiated As Boolean

Public Sub InitSysADLVisioStencil(ByVal aStencilDocument As Visio.Document)

    If Not sysADLVisioStencilAlreadyInitiated Then
    
        Debug.Print "Sys-ADL stencil initiated as " + aStencilDocument.Name
    
        Set sysADLVisioStencil = aStencilDocument
    
    End If


End Sub

Public Function GetSysADLVisioStencil() As Visio.Document
    
    Set GetSysADLVisioStencil = sysADLVisioStencil

End Function




Public Sub ShowWarnMessage(ByVal MessageCode As String)
    
    Dim Message As String
    Dim Title As String
    
    Message = FactoryMessageText.GetMessageText(MessageCode)
    Title = FactoryMessageText.GetMessageTitle(MessageCode)
    
    MsgBox Message, vbExclamation, Title

End Sub

Public Sub LogMessage(ByVal Message As String)

    Dim ConfigLogEnabled As String
    
    ConfigLogEnabled = FactoryConfigProperty.GetProperty(sysAdlConfigPropertyLogEnabled)
    
    If ConfigLogEnabled = sysAdlGuiServicesLogIsEnabled Then
        
        Debug.Print "---------------------------------------------------------------------------------"
        Debug.Print Now
        Debug.Print "---------------------------------------------------------------------------------"
    
        Debug.Print Message
        
        Debug.Print "                                                                              ..."
        
    End If

End Sub

Public Function PrepareMasterNameForElement(ByVal AMasterName As Variant) As String

    Dim PointLocation As Integer
    Dim Result As String
    
    Result = AMasterName
    PointLocation = InStr(AMasterName, ".")
    
    If (PointLocation > 0) Then
        Result = Left(Result, (PointLocation - 1))
    End If
    
    PrepareMasterNameForElement = Result

End Function

Public Function PrepareFieldForSQL(ByVal StringToSQL As Variant) As String

    Dim Result As String
    
    If IsNull(StringToSQL) Then
        StringToSQL = ""
    End If
    
    Result = Replace(StringToSQL, "'", "´")

    PrepareFieldForSQL = Result

End Function

Public Function PrepareFieldForXML(ByVal StringToXML As Variant) As String

    Dim Result As String
    
    If IsNull(StringToXML) Then
        StringToXML = ""
    End If
    
    Result = StringToXML
    
    Result = Replace(Result, "&", "&amp;")
    Result = Replace(Result, "<", "&lt;")
    Result = Replace(Result, ">", "&gt;")
    Result = Replace(Result, """", "&quot;")
    Result = Replace(Result, "'", "&apos;")

    PrepareFieldForXML = Result

End Function

Public Function PredefinedXMLToField(ByVal XMLToString As Variant) As String

    Dim Result As String
    
    If IsNull(XMLToString) Then
        XMLToString = ""
    End If
    
    Result = XMLToString
    
    Result = Replace(Result, "&amp;", "&")
    Result = Replace(Result, "&lt;", "<")
    Result = Replace(Result, "&gt;", ">")
    Result = Replace(Result, "&quot;", """")
    Result = Replace(Result, "&apos;", "'")

    PredefinedXMLToField = Result

End Function

Public Function PrepareFieldForView(ByVal StringToSQL As Variant) As String

    Dim Result As String
    
    If IsNull(StringToSQL) Then
        StringToSQL = ""
    End If
    
    Result = Replace(StringToSQL, Chr(34), sysAdlStringConstantsEmpty)

    PrepareFieldForView = Result

End Function

Public Function PrepareFieldForCustomPropertyKey(ByVal StringToKey As Variant) As String

    Dim Result As String
    
    If IsNull(StringToKey) Then
        StringToKey = ""
    End If
    
    Result = Replace(StringToKey, " ", "_")
    
    PrepareFieldForCustomPropertyKey = Result

End Function

Public Sub ProgressBarViewInit(ByVal GlobalTaskName As String)

    ProgressBarView.Caption = GlobalTaskName
    ProgressBarView.ProgressBar.value = sysAdlIntegerConstantsZero
    ProgressBarView.ProgressBar.Max = 100
    ProgressBarView.ProgressBar.Min = sysAdlIntegerConstantsZero
    ProgressBarView.TaskLabel.Caption = "Initing task..."
    ProgressBarView.Show vbModeless
    
End Sub

Public Sub ProgressBarViewFinish()

    ProgressBarView.Hide

End Sub


Public Sub ProgressBarUpdate(ByVal CurrentTaskName As String, ByVal TotalSteps As Integer, ByVal CurrentStep As Integer)

    Dim CurrentProgress As Integer

    CurrentProgress = (CurrentStep / TotalSteps) * 100
    
    If CurrentProgress > 100 Then
        CurrentProgress = 100
    End If

    ProgressBarView.TaskLabel.Caption = CurrentTaskName
    ProgressBarView.ProgressBar.value = CurrentProgress
    ProgressBarView.Repaint

End Sub

Public Function GetQualifierFromDocumentName(ByVal fileName As String, Optional ByVal fileExtension As String = sysadlstringconstantsExtensionVisio) As String

    Dim basePath As String
    Dim Result As String
    Dim PathLength As Integer
    Dim FileNameLength As Integer
    
    fileName = ChangeDocumentExtension(fileName, fileExtension, sysAdlStringConstantsEmpty)
    
    FileNameLength = Len(fileName)
    basePath = FactoryConfigProperty.GetProperty(sysAdlConfigPropertyBasePath)


    Result = fileName
    Result = Left(Result, FileNameLength)
    Result = Replace(Result, basePath, sysAdlStringConstantsEmpty)
    Result = Replace(Result, sysAdlStringConstantsWindowsPathSeparator, sysAdlStringConstantsNamespacePathSeparator)

    GetQualifierFromDocumentName = Result

End Function

Public Function ChangeDocumentExtension(ByVal fileName As String, _
                                         ByVal NewExtension As String, _
                                         Optional ByVal OldExtension As String = sysadlstringconstantsExtensionVisio) As String
    
    Dim Result As String
    Dim FileNameLength As Integer
    Dim ExtensionLength As Integer
    Dim StartChangePosition As Integer
    Const SECURITY_MARGIN = 2
    
    
    FileNameLength = Len(fileName)
    ExtensionLength = Len(NewExtension)
    
    StartChangePosition = FileNameLength - ExtensionLength - SECURITY_MARGIN
    
    Result = Replace(fileName, OldExtension, NewExtension)

    ChangeDocumentExtension = Result

End Function

Public Function GetDocumentNamespace() As String

    Dim Result As String
    Dim DocumentHasAlreadyBeenSaved As Boolean

    DocumentHasAlreadyBeenSaved = (ActiveDocument.Path <> sysAdlStringConstantsEmpty)

    If DocumentHasAlreadyBeenSaved Then
        Result = GetNamespaceFromFilePath(ActiveDocument.Path)
    Else
        Result = sysAdlStringConstantsEmpty
    End If

    GetDocumentNamespace = Result
    
End Function

Private Function GetNamespaceFromFilePath(ByVal PathName As String) As String

    Dim basePath As String
    Dim Result As String
    Dim PathLength As Integer
    
    PathLength = Len(PathName)
    basePath = FactoryConfigProperty.GetProperty(sysAdlConfigPropertyBasePath)


    Result = PathName
    Result = Left(Result, PathLength - 1)
    Result = Replace(Result, basePath, sysAdlStringConstantsEmpty)
    Result = Replace(Result, sysAdlStringConstantsWindowsPathSeparator, sysAdlStringConstantsNamespacePathSeparator)

    GetNamespaceFromFilePath = Result

End Function

Public Function GetNamespaceFromString(ByVal aKey As String) As String

    Dim Result As String
    Dim keyCollection As Variant
    Dim I As Integer
    Dim Upper As Integer
    Dim Lower As Integer
    
    Dim KeyHasNoNamespace As Boolean
    Dim keyHasOnlyOnePath As Boolean
    Dim keyHasComposedPath As Boolean
    
    keyCollection = Split(aKey, sysAdlStringConstantsNamespacePathSeparator)
    
    Upper = UBound(keyCollection)
    Lower = LBound(keyCollection)
    
    KeyHasNoNamespace = False
    keyHasOnlyOnePath = False
    keyHasComposedPath = False
    
    If Upper <= 0 Then
        
        KeyHasNoNamespace = True
    
    ElseIf Upper = 1 Then
        
        keyHasOnlyOnePath = True
    Else
        
        keyHasComposedPath = True
        
    End If
    
    If KeyHasNoNamespace Then
    
        Result = sysAdlStringConstantsEmpty
        
    ElseIf keyHasOnlyOnePath Then
    
        Result = keyCollection(0)
        
    ElseIf keyHasComposedPath Then
    
        Result = keyCollection(0)
    
        For I = 1 To Upper - 2
        
            Result = Result + sysAdlStringConstantsNamespacePathSeparator + keyCollection(I)
        
        Next
    
        Result = Result + sysAdlStringConstantsNamespacePathSeparator + keyCollection(Upper - 1)
    
    End If
    
    GetNamespaceFromString = Result
    
End Function

Public Function GetIdFromString(ByVal aKey As String) As String

    Dim Result As String
    Dim keyCollection As Variant
    Dim Upper As Integer
    Dim KeyHasNoId As Boolean
    Dim KeyHasNoNamespace As Boolean
    Dim KeyHasNamespace As Boolean
    
    keyCollection = Split(aKey, sysAdlStringConstantsNamespacePathSeparator)
    
    Upper = UBound(keyCollection)
    
    KeyHasNoId = False
    KeyHasNoNamespace = False
    KeyHasNamespace = False
    
    If (Upper < 0) Then
    
        KeyHasNoId = True
        
    ElseIf (Upper = 0) Then
    
        KeyHasNoNamespace = True
        
    Else
    
        KeyHasNamespace = True
        
    End If
    
    If (KeyHasNoId) Then
    
        Result = sysAdlStringConstantsEmpty
        
    ElseIf KeyHasNoNamespace Then
    
        Result = keyCollection(0)
        
    ElseIf KeyHasNamespace Then
    
        Result = keyCollection(Upper)
        
    End If
        
    GetIdFromString = Result
    
End Function


Public Function IsRelationElementType(ByVal elementType As String) As Boolean

    Dim Result As Boolean
    
    Result = False
        
    Select Case elementType
    
        Case sysAdlTypeSetConnector: Result = True
        Case sysAdlTypeSetComposedBy: Result = True
        Case sysAdlTypeSetDependsOn: Result = True
        Case sysAdlTypeSetIsA: Result = True
        Case sysAdlTypeSetRepresents: Result = True
        
    End Select
    
    IsRelationElementType = Result
    
End Function

Public Function GetPathFromNamespace(ByVal aNamespace As String)

    Dim Result As String
    Dim basePath As String
    
    basePath = FactoryConfigProperty.GetProperty(sysAdlConfigPropertyBasePath)

    Result = aNamespace
    Result = Replace(Result, sysAdlStringConstantsNamespacePathSeparator, sysAdlStringConstantsWindowsPathSeparator)
    Result = basePath + Result

    GetPathFromNamespace = Result
    

End Function


Public Function GetFullFileNameForElement(ByVal aNamespace As String, ByVal anId As String)

    Dim Result As String
    Dim FolderName As String
    
    FolderName = GetPathFromNamespace(aNamespace)

    Result = FolderName + _
             sysAdlStringConstantsWindowsPathSeparator + _
             anId + _
             sysadlstringconstantsExtensionSysAdlElement
             
    GetFullFileNameForElement = Result

End Function
