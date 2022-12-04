Attribute VB_Name = "mKeepTop"
Option Explicit

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub MakeNormal(Handle As Long)
SetWindowPos Handle, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MakeTopMost(Handle As Long)
SetWindowPos Handle, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub




