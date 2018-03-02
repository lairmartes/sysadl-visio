Attribute VB_Name = "FactoryMessageText"
Option Explicit
    Dim MessageTexts As New UtilSysAdlListString
    Dim IsMessageFileLoaded As Boolean

    Private Const sysAdlMessageTextSufix = "_Message"
    Private Const sysAdlMessageTitleSufix = "_Title"
    
    Public Const sysAdlMessageProblemsSaving = "6000"
    Public Const sysAdlMessageKeyAlreadyInUse = "6001"
    Public Const sysAdlMessageInvalidDiagram = "6002"
    Public Const sysAdlMessageWithoutObjective = "6003"
    
    Private Const sysAdlMessageFileName = "sysadl-messages-1.0.properties"

Public Function GetMessageText(ByVal MessageKey As String) As String
    
    Dim Result As String
    
    LoadMessageFile
    
    Result = MessageTexts.Item(MessageKey + sysAdlMessageTextSufix)
    
    GetMessageText = Result

End Function

Public Function GetMessageTitle(ByVal MessageKey As String) As String
    
    Dim Result As String
    
    LoadMessageFile
    
    Result = MessageTexts.Item(MessageKey + sysAdlMessageTitleSufix)
    
    GetMessageTitle = Result

End Function
Private Sub LoadMessageFile()

    On Error GoTo HandleErrorWhileFillingConfigFiles

    If Not IsMessageFileLoaded Then
    
        Dim HomeFolder As String
        Dim MessageFilePath As String
        Dim MessageFileNumber As Integer
        
        Dim MessageLine As String
        Dim CurrentKey As Variant
        Dim CurrentValue As String
        Dim CurrentEqualSignalPosition As Integer
        Dim currentMessage As UtilSysAdlItemString
        
        HomeFolder = VBA.Environ$("APPDATA") + "\sysadl\"
        MessageFilePath = HomeFolder + sysAdlMessageFileName
        MessageFileNumber = FreeFile()
        
        Open MessageFilePath For Input As #MessageFileNumber
        
        While Not EOF(1)
            Line Input #MessageFileNumber, MessageLine
            CurrentEqualSignalPosition = InStr(1, MessageLine, "=")
            CurrentKey = Left(MessageLine, CurrentEqualSignalPosition - 1)
            CurrentValue = Mid(MessageLine, CurrentEqualSignalPosition + 1)
            
            CurrentKey = Trim(CurrentKey)
            CurrentValue = Trim(CurrentValue)
            
            Set currentMessage = New UtilSysAdlItemString
            
            currentMessage.Init CurrentKey, CurrentValue
            
            MessageTexts.Add currentMessage
        Wend
        
        Close #MessageFileNumber
        
        IsMessageFileLoaded = True
        
    End If
    
HandleErrorWhileFillingConfigFiles:
    Dim ErrorMessage As String
    
    If Err.Number = 53 Then
        
         ErrorMessage = "Message file " + MessageFilePath + " has not been found and system will not work properly."
         ErrorMessage = ErrorMessage & vbCrLf + "Please check if file " + sysAdlMessageFileName + " exists in folder " + HomeFolder + "."
    
        MsgBox ErrorMessage, vbExclamation, "Error while reading config file"
    
    End If
    

End Sub

