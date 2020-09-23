Attribute VB_Name = "Module1"
Declare Function GetCursorPos& Lib "user32" (lpPoint As PointAPI)
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Type PointAPI  '  8 Bytes
     x As Long
     y As Long
End Type

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Sub FormOnTop(frm As Form, condition As Boolean)
'toggles whether a form is on top of all other windows
'if condition = true then form is on top
'if condition = false then form is not top

Dim SetOnTop As Long

    Select Case condition
    
    Case True:
        SetOnTop = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Case False:
        SetOnTop = SetWindowPos(frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)

    End Select
    
End Sub
