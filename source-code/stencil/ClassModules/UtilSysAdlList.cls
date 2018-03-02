VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilSysAdlList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private MyList As New Collection
Private Const UTIL_ITEM_CLASS_NAME = "UtilSysAdlItem"

Public Sub Add(ByVal AnItem As UtilSysAdlItem)

    MyList.Add AnItem

End Sub


Public Function Item(ByVal aKey As String) As Variant
    Dim CurrentItem As Variant
    Dim Result As Variant
    Dim ItemFound As Boolean
    
    For Each CurrentItem In MyList
        If CurrentItem.Key = aKey Then
            Set Result = CurrentItem.Item
            ItemFound = True
            Exit For
        End If
    Next
    
    If Not ItemFound Then
        Set Item = Nothing
    Else
        Set Item = Result
    End If

End Function

Public Sub Remove(ByVal aKey As String)
    
    Dim CurrentItem As UtilSysAdlItem
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

Public Function ItemByIndex(ByVal AnIndex As Integer) As Variant

    Dim ItemUtil As UtilSysAdlItem
    Dim Result As Variant
    
    Set ItemUtil = MyList.Item(AnIndex)
    Set Result = ItemUtil.Item
    
    Set ItemByIndex = Result
    
End Function

Public Function Clone() As UtilSysAdlList

    Dim Result As New UtilSysAdlList
    Dim CurrentItem As UtilSysAdlItem
    
    For Each CurrentItem In MyList
        Result.Add CurrentItem
    Next

    Set Clone = Result
    
End Function
