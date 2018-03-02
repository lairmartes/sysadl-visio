VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilSysAdlItemString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private pKey As String
Private pItem As String
Private pDirty As Boolean

Public Sub Init(ByVal aKey As String, ByVal AnItem As Variant)

    pKey = aKey
    pItem = AnItem
    
    pDirty = True
    
End Sub

Public Property Get Key() As String

    Key = pKey
    
End Property

Public Property Get Item() As String

    Item = pItem

End Property

Public Property Let Item(newItem As String)

    pItem = newItem
    
    pDirty = True

End Property

Public Sub Clean()

    pDirty = False

End Sub

Public Property Get IsDirty() As Boolean

    IsDirty = pDirty
    
End Property
