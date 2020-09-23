Attribute VB_Name = "Tray"
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As enm_NIM_Shell, pnid As NOTIFYICONDATA) As Boolean
Public nidProgramData As NOTIFYICONDATA
Public Const WM_MOUSEISMOVING = &H200  ' Mouse is moving
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_SETHOTKEY = &H32
Public Type NOTIFYICONDATA              'Icon Type
         cbSize As Long
         hwnd As Long
         uId As Long
         uFlags As Long
         uCallbackMessage As Long
         hIcon As Long
         szTip As String * 64
End Type
Public Enum enm_NIM_Shell               'Send Type
         NIM_ADD = &H0
         NIM_MODIFY = &H1
         NIM_DELETE = &H2
         NIF_MESSAGE = &H1
         NIF_ICON = &H2
         NIF_TIP = &H4
         WM_MOUSEMOVE = &H200
End Enum

