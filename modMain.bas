Attribute VB_Name = "modMain"

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Const EM_SETPASSWORDCHAR = &HCC
Public Const EM_GETPASSWORDCHAR = &HD2
Public Const EM_SETMODIFY = &HB9
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function SendMessage Lib "User" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
