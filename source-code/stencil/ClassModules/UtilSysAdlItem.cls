VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilSysAdlItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private pKey As String
Private pItem As Variant

Public Sub Init(ByVal aKey As String, ByVal AnItem As Variant)

    pKey = aKey
    Set pItem = AnItem
    
End Sub

Public Property Get Key() As String

    Key = pKey
    
End Property

Public Property Get Item() As Variant

    Set Item = pItem

End Property

