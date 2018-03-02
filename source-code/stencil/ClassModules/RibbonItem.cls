VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private pId As String
    Private pLabel As String
    Private pTip As String
    Private pImageMSO As String
    Private pIsLarge As Boolean
    Private pCommand As ICommand
    

Public Sub Init(ByVal anId As String, _
                ByVal aLabel As String, _
                ByVal aTip As String, _
                ByVal anImageMSO As String, _
                ByVal IsLarge As Boolean, _
                ByVal aCommand As ICommand)
                              
                              
    pId = anId
    pLabel = aLabel
    pTip = aTip
    pImageMSO = anImageMSO
    pIsLarge = IsLarge
    Set pCommand = aCommand
                              
End Sub


Public Property Get Id() As String

    Id = pId
    
End Property

Public Property Get Label() As String

    Label = pLabel

End Property

Public Property Get Tip() As String

    Tip = pTip
    
End Property
                              
    
Public Property Get ImageMSO() As String

    ImageMSO = pImageMSO

End Property


Public Property Get IsLarge() As Boolean

    IsLarge = pIsLarge
    
End Property

Public Sub ExecuteCommand()

    pCommand.Execute
    
End Sub


