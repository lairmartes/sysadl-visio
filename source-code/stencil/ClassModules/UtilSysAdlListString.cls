VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilSysAdlListString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private MyList As New Collection
Private Const UTIL_ITEM_CLASS_NAME = "UtilSysAdlItem"

Public Sub Add(ByVal AnItem As UtilSysAdlItemString)

    MyList.Add AnItem

End Sub


Public Function Item(ByVal aKey As String) As String
    Dim CurrentItem As UtilSysAdlItemString
    Dim Result As String
    Dim ItemFound As Boolean
    
    For Each CurrentItem In MyList
        If CurrentItem.Key = aKey Then
            Result = CurrentItem.Item
            ItemFound = True
            Exit For
        End If
    Next
    
    If Not ItemFound Then
        Item = sysAdlStringConstantsEmpty
    Else
        Item = Result
    End If

End Function
Public Function ItemIsDirty(ByVal aKey As String) As Boolean
    Dim CurrentItem As UtilSysAdlItemString
    Dim Result As Boolean
    Dim ItemFound As Boolean
        
    Result = False
    
    For Each CurrentItem In MyList
        If CurrentItem.Key = aKey Then
            Result = CurrentItem.IsDirty
            Exit For
        End If
    Next
    
    ItemIsDirty = Result
    
End Function

Public Function KeyExists(ByVal aKey As String) As Boolean
    Dim CurrentItem As UtilSysAdlItemString
    Dim Result As Boolean
    
    Result = False
    
    For Each CurrentItem In MyList
        If CurrentItem.Key = aKey Then
            Result = True
            Exit For
        End If
    Next

    KeyExists = Result

End Function

Public Sub Remove(ByVal aKey As String)
    
    Dim CurrentItem As UtilSysAdlItemString
    Dim CurrentIndex As Integer
    
    For CurrentIndex = 1 To MyList.Count
        Set CurrentItem = MyList.Item(CurrentIndex)
        If CurrentItem.Key = aKey Then
            MyList.Remove CurrentIndex
            Exit For
        End If
    Next
    
End Sub

Public Function Count() As Integer

    Count = MyList.Count

End Function

Public Function ItemByIndex(ByVal AnIndex As Integer) As String

    Dim ItemUtil As UtilSysAdlItem
    Dim Result As String
    
    Set ItemUtil = MyList.Item(AnIndex)
    Result = ItemUtil.Item
    
    ItemByIndex = Result
    
End Function

Public Function Clone() As UtilSysAdlListString

    Dim Result As New UtilSysAdlListString
    Dim CurrentItem As UtilSysAdlItemString
    
    For Each CurrentItem In MyList
        Result.Add CurrentItem
    Next

    Set Clone = Result
    
End Function

Public Sub ChangeValue(ByVal aKey As String, ByVal newValue As String)

    Dim CurrentItem As UtilSysAdlItemString
    Dim CurrentIndex As Integer
    
    For CurrentIndex = 1 To MyList.Count
        Set CurrentItem = MyList.Item(CurrentIndex)
        If CurrentItem.Key = aKey Then
            CurrentItem.Item = newValue
            Exit For
        End If
    Next
    
End Sub

Public Sub CleanDirtyItems()

    Dim CurrentItem As UtilSysAdlItemString
    Dim CurrentIndex As Integer
    
    For CurrentIndex = 1 To MyList.Count
        Set CurrentItem = MyList.Item(CurrentIndex)
        CurrentItem.Clean
    Next
    
End Sub

Public Function GetUtilItemCollection() As Collection

    Dim Result As Collection
    
    Set Result = MyList
    
    Set GetUtilItemCollection = Result
    
End Function
