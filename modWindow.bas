Attribute VB_Name = "Window"
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_SYSMENU = &H80000
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = (-4)

Function HasMax(hwnd As Long) As Boolean
    Dim Style As Long
    Style = GetWindowLong(hwnd, GWL_STYLE)
    If ((Style And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX) = True Then
        HasMax = True
    Else
        HasMax = False
    End If
End Function

Function HasMin(hwnd As Long) As Boolean
    Dim Style As Long
    Style = GetWindowLong(hwnd, GWL_STYLE)
    If ((Style And WS_MINIMIZEBOX) = WS_MINIMIZEBOX) = True Then
        HasMin = True
    Else
        HasMin = False
    End If
End Function

Function HasTitleBar(hwnd As Long) As Boolean
    Dim Style As Long
    Style = GetWindowLong(hwnd, GWL_STYLE)
    If ((Style And WS_DLGFRAME) = WS_DLGFRAME = True) Then
        HasTitleBar = True
    Else
        HasTitleBar = False
    End If
End Function
