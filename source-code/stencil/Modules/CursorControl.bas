Attribute VB_Name = "CursorControl"
Option Explicit

'Declare Windows API Constants for Windows System cursors.
Public Const IDC_APPSTARTING = 32650&    'Standard arrow and small hourglass.
Public Const IDC_ARROW = 32512&          'Standard arrow.
Public Const IDC_CROSS = 32515           'Crosshair.
Public Const IDC_HAND = 32649            'Hand.
Public Const IDC_HELP = 32651            'Arrow and question mark.
Public Const IDC_IBEAM = 32513&          'Text I-beam.
Public Const IDC_ICON = 32641&           'Windows NT only: Empty icon.
Public Const IDC_NO = 32648&             'Slashed circle.
Public Const IDC_SIZE = 32640&           'Windows NT only: Four-pointed arrow.
Public Const IDC_SIZEALL = 32646&        'Four-pointed arrow pointing north, south, east, and west.
Public Const IDC_SIZENESW = 32643&       'Double-pointed arrow pointing northeast and southwest.
Public Const IDC_SIZENS = 32645&         'Double-pointed arrow pointing north and south.
Public Const IDC_SIZENWSE = 32642&       'Double-pointed arrow pointing northwest and southeast.
Public Const IDC_SIZEWE = 32644&         'Double-pointed arrow pointing west and east.
Public Const IDC_UPARROW = 32516&        'Vertical arrow.
Public Const IDC_WAIT = 32514&           'Hourglass.

'Declarations for API Functions.
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long

'Declare handles for cursor.
Private hOldCursor As Long
Private hNewCursor As Long

'The UseCursor function will load and set a system cursor or a cursor from file to a
'controls event property.
Public Function UseCursor(ByVal NewCursor As Variant)

    'Load new cursor.
    Select Case TypeName(NewCursor)
        Case "String" 'Custom cursor from file.
            hNewCursor = LoadCursorFromFile(NewCursor)
        Case "Long", "Integer" 'System cursor.
            hNewCursor = LoadCursor(ByVal 0&, NewCursor)
        Case Else 'Do nothing
    End Select
    'If successful set new cursor.
    If (hNewCursor > 0) Then
        hOldCursor = SetCursor(hNewCursor)
    End If
    'Clean up.
    hOldCursor = DestroyCursor(hNewCursor)
    hNewCursor = DestroyCursor(hOldCursor)

End Function
