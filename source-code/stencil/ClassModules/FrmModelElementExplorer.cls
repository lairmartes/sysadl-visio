VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FrmModelElementExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pElementKeyList As Collection
    Private pNamespaceList As Collection
    Private beforeBasePathFolderName As String
    Private basePath  As String
    
    
Public Sub Class_Initialize()

    basePath = FactoryConfigProperty.GetProperty(sysAdlConfigPropertyBasePath)
    beforeBasePathFolderName = GetPreviousFolder(basePath)
    BuildFilesAndFolderList

End Sub

Public Property Get elementKeyList() As Collection

    Set elementKeyList = pElementKeyList

End Property

Public Property Get namespaceList() As Collection

    Set namespaceList = pNamespaceList

End Property

Private Sub BuildFilesAndFolderList()

    Set pElementKeyList = New Collection
    Set pNamespaceList = New Collection
    
    populateQualifiers basePath

End Sub

Private Sub populateQualifiers(ByVal dirName As String)

    Dim folderFound As String
    Dim fileName As String
    Dim elementFile As String
    Dim ElementKey As String
    Dim elementNamespace As String
    Dim completePath As String
    Dim isSysAdlElementFile As Boolean
    Dim isDirectoryExists As Boolean
    Dim directoryFoundList As Collection
    Dim directoryFoundQtty As Integer
    Dim isNamespaceAlreadyInList As Boolean
    Dim I As Integer
    
    Set directoryFoundList = New Collection
    
    If dirName = beforeBasePathFolderName Then Exit Sub
    
    isDirectoryExists = (Len(Dir(dirName, vbDirectory)) > 0)
    
    If Not isDirectoryExists Then Exit Sub
    
    fileName = Dir(dirName)
    
    isNamespaceAlreadyInList = False
    
    Do While fileName <> sysAdlStringConstantsEmpty
    
        isSysAdlElementFile = StringEndsWith(fileName, sysadlstringconstantsExtensionSysAdlElement, vbTextCompare)
        
        If (isSysAdlElementFile) Then
    
            elementFile = dirName + fileName + sysAdlStringConstantsWindowsPathSeparator
            ElementKey = GetElementKeyFromElementFileName(elementFile)
            
            Debug.Print ElementKey
            
            pElementKeyList.Add ElementKey
            
            If Not isNamespaceAlreadyInList Then
            
                elementNamespace = GUIServices.GetNamespaceFromString(ElementKey)
            
                pNamespaceList.Add elementNamespace
                
                isNamespaceAlreadyInList = True
                
            End If
            
        End If
        
        fileName = Dir()
        
    Loop
    
    
    'create a list containing the folders and search each one.  This will avoid the error when Dir() is recalled

    folderFound = Dir(dirName, vbDirectory)
    
    Do While folderFound <> sysAdlStringConstantsEmpty
    
        If (folderFound <> "." And folderFound <> ".." And Not StringContainsCharacter(folderFound, ".")) Then
        
            completePath = dirName + folderFound + sysAdlStringConstantsWindowsPathSeparator
    
            directoryFoundList.Add completePath
            
        End If
        
        folderFound = Dir(, vbDirectory)
        
    Loop
    
    
    directoryFoundQtty = directoryFoundList.Count
    
    For I = 1 To directoryFoundQtty
    
        completePath = directoryFoundList(I)
    
        populateQualifiers (completePath)
        
    Next I
    
End Sub

' code from http://www.freevbcode.com/ShowCode.asp?ID=2856
Private Function StringEndsWith(ByVal strValue As String, _
   CheckFor As String, Optional CompareType As VbCompareMethod _
   = vbBinaryCompare) As Boolean
 'Determines if a string ends with the same characters as
 'CheckFor string
 
 'True if end with CheckFor, false otherwise

 'Case sensitive by default.  If you want non-case sensitive, set
 'last parameter to vbTextCompare
 
  'Examples
  'MsgBox StringEndsWith("Test", "ST") 'False
  'MsgBox StringEndsWith("Test", "ST", vbTextCompare) 'True

  Dim sCompare As String
  Dim lLen As Long
   
  lLen = Len(CheckFor)
  If lLen > Len(strValue) Then Exit Function
  sCompare = Right(strValue, lLen)
  StringEndsWith = StrComp(sCompare, CheckFor, CompareType) = 0

End Function


Private Function GetPreviousFolder(ByVal FolderName As String) As String


    Dim Result As String
    Dim FolderPieces() As String
    Dim I As Integer
    Dim qttyOfFolders As Integer
    
    FolderPieces = Split(FolderName, sysAdlStringConstantsWindowsPathSeparator)
    
    qttyOfFolders = UBound(FolderPieces)
    
    Result = ""
    
    For I = 0 To qttyOfFolders - 2
    
        Result = Result + FolderPieces(I) + sysAdlStringConstantsWindowsPathSeparator
    
    Next
    
    GetPreviousFolder = Result


End Function

Private Function StringContainsCharacter(ByVal text As String, ByVal textToSearch As String) As Boolean

    Dim Result As Boolean

    Result = InStr(1, text, textToSearch, vbTextCompare) > 0
    
    StringContainsCharacter = Result

End Function


Private Function GetElementKeyFromElementFileName(ByVal PathName As String) As String

    Dim basePath As String
    Dim Result As String
    Dim PathLength As Integer
    
    PathLength = Len(PathName)
    
    basePath = FactoryConfigProperty.GetProperty(sysAdlConfigPropertyBasePath)

    Result = PathName
    Result = Left(Result, PathLength - 1)
    Result = Replace(Result, basePath, sysAdlStringConstantsEmpty)
    Result = Replace(Result, sysadlstringconstantsExtensionSysAdlElement, sysAdlStringConstantsEmpty)
    Result = Replace(Result, sysAdlStringConstantsWindowsPathSeparator, sysAdlStringConstantsNamespacePathSeparator)

    GetElementKeyFromElementFileName = Result

End Function


