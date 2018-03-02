VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueueCommandExecution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private CommandList As New Collection

Public Sub AddCommand(ByVal aCommand As ICommand)

    CommandList.Add aCommand

End Sub

Public Sub ExecuteCommandList()

    Dim CurrentCommand As ICommand

    For Each CurrentCommand In CommandList
        CurrentCommand.Execute
    Next

End Sub

