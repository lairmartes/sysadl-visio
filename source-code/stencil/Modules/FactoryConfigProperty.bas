Attribute VB_Name = "FactoryConfigProperty"
Option Explicit
    Dim Properties As New UtilSysAdlListString
    Dim IsPropertyFileLoaded As Boolean
    
    Public Const sysAdlConfigPropertyBasePath = "BasePath"
    Public Const sysAdlConfigPropertyBaseURLPublish = "BaseURLPublish"
    Public Const sysAdlConfigPropertyLogEnabled = "LogEnabled"
    Public Const sysAdlConfigPropertyElementDefinitionPath = "ElementDefinitionPath"
    Public Const sysAdlConfigPropertyElementDefinitionAlwaysRefresh = "ElementDefinitionAlwaysRefresh"
    
    Private Const sysAdlConfigPropertyFileName = "sysadl-config-2.5.properties"
    
Public Function GetProperty(ByVal PropertyKey As String) As String
    
    Dim Result As String
    
    LoadPropertyFile
    
    Result = Properties.Item(PropertyKey)

    GetProperty = Result

End Function

Private Sub LoadPropertyFile()

    On Error GoTo HandleErrorWhileFillingConfigFiles

    If Not IsPropertyFileLoaded Then
    
        Dim HomeFolder As String
        Dim ConfigFilePath As String
        
        Dim ConfigLine As String
        Dim CurrentKey As Variant
        Dim CurrentValue As String
        Dim CurrentEqualSignalPosition As Integer
        Dim CurrentProperty As UtilSysAdlItemString
        
        HomeFolder = VBA.Environ$("APPDATA") + "\sysadl\"
        ConfigFilePath = HomeFolder + sysAdlConfigPropertyFileName
        
        Open ConfigFilePath For Input As #1
        
        While Not EOF(1)
            Line Input #1, ConfigLine
            CurrentEqualSignalPosition = InStr(1, ConfigLine, "=")
            CurrentKey = Left(ConfigLine, CurrentEqualSignalPosition - 1)
            CurrentValue = Mid(ConfigLine, CurrentEqualSignalPosition + 1)
            
            CurrentKey = Trim(CurrentKey)
            CurrentValue = Trim(CurrentValue)
            
            Set CurrentProperty = New UtilSysAdlItemString
            
            CurrentProperty.Init CurrentKey, CurrentValue
            
            Properties.Add CurrentProperty
        Wend
        
        Close #1
        
        IsPropertyFileLoaded = True
        
    End If
    
HandleErrorWhileFillingConfigFiles:
    Dim ErrorMessage As String
    
    If Err.Number = 53 Then
        
         ErrorMessage = "Configuration file " + ConfigFilePath + " has not been found and system will not work properly."
         ErrorMessage = ErrorMessage & vbCrLf + "Please check if file " + sysAdlConfigPropertyFileName + " exists in folder " + HomeFolder + "."
    
        MsgBox ErrorMessage, vbExclamation, "Error while reading config file"
    
    End If
    

End Sub
