Attribute VB_Name = "Killer"
Option Explicit

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Public Const HWND_BOTTOM = 1
Public Const HWND_BROADCAST = &HFFFF&
Public Const HWND_DESKTOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Global Const WM_CLOSE As Long = &H10

'// Declarations for Stopper
Global Blacklist() As String
Global PopCount As Integer
Global PlaySound As Boolean
Global UseKeyword As Boolean
Global KillMethod As Integer
Global SoundPath As String
Global StereotypeWindows As Integer
Global NoMaxButton As Integer
Global NoMinButton As Integer
Global NoTitleBar As Integer

Function LoadBlackList()
    Dim NowFile
    Dim ILLCount As Integer
    ILLCount = 0
    NowFile = FreeFile
    MainForm.lstBlack.Clear
    Open App.Path & "\blacklist.dat" For Input As NowFile
        While Not EOF(NowFile)
            ILLCount = ILLCount + 1
            ResBlacklist (ILLCount)
            Input #NowFile, Blacklist(ILLCount)
            MainForm.lstBlack.AddItem Blacklist(ILLCount)
        Wend
    Close NowFile
End Function

Function ClearBlacklist()
    Dim ClearLoop As Integer
    For ClearLoop = 0 To UBound(Blacklist)
        Blacklist(ClearLoop) = ("")
    Next ClearLoop
End Function

Function GetRealText(hwnd As Long) As String
    On Error Resume Next
    Dim Buffer As String
    Dim TextLength As Integer
    Dim WindowText As String
    Dim Delim As String
    Buffer = String(100, Chr(0))
    TextLength = GetWindowTextLength(hwnd)
    WindowText = GetWindowText(hwnd, Buffer, TextLength + 1)
    Buffer = Mid$(Buffer, 1, Len(Buffer) - 1)
    WindowText = Buffer
    Delim = InStrRev(WindowText, "-")
    WindowText = Left(WindowText, Delim - 2)
    GetRealText = WindowText
End Function

Function ResBlacklist(newSize As Integer)
    ReDim Preserve Blacklist(newSize)
End Function

Function AddBlackItem(BlackItem As String)
    On Error GoTo Resover
    ResBlacklist (UBound(Blacklist) + 1)
GoBack:
    Blacklist(UBound(Blacklist)) = BlackItem
    Call WriteBlackList
    Exit Function
Resover:
    ResBlacklist (1)
    GoTo GoBack
End Function

Function WriteBlackList()
    Dim NowFile
    Dim Writeloop As Integer
    NowFile = FreeFile
    Open App.Path & "/blacklist.dat" For Output As NowFile
        For Writeloop = 1 To UBound(Blacklist)
            Print #NowFile, Blacklist(Writeloop)
        Next Writeloop
    Close NowFile
End Function

Function CheckBlackList(BlackItem As String) As Boolean
    Dim Checkloop As Integer
    For Checkloop = 1 To UBound(Blacklist)
        If Blacklist(Checkloop) = BlackItem Then
            CheckBlackList = True
            Exit Function
        End If
    Next Checkloop
    CheckBlackList = False
End Function

Function GetBlackItemPos(BlackItem As String) As Integer
    Dim CurrLoop As Integer
    For CurrLoop = 1 To UBound(Blacklist)
        If Blacklist(CurrLoop) = BlackItem Then
            GetBlackItemPos = CurrLoop
            Exit Function
        End If
    Next CurrLoop
    GetBlackItemPos = -1
End Function

Function DeleteBlackItem(BlackItem As String)
    Dim DelPos As Integer
    Dim KiLLLoop As Integer
    DelPos = GetBlackItemPos(BlackItem)
    For KiLLLoop = DelPos To UBound(Blacklist) - 1
        Blacklist(KiLLLoop) = Blacklist(KiLLLoop + 1)
    Next KiLLLoop
    ResBlacklist (UBound(Blacklist) - 1)
    Call WriteBlackList
    Call LoadBlackList
End Function

Function DisplayIEWindows()
    Dim CurrHWND As Long
    MainForm.lstOpen.Clear
    CurrHWND = GetNextWindow(MainForm.hwnd, GW_HWNDFIRST)
    If CheckClass(CurrHWND) = ("IEFrame") And GetRealText(CurrHWND) <> ("Cannot find server") Then MainForm.lstOpen.AddItem GetRealText(CurrHWND)
    While CurrHWND <> 0
        CurrHWND = GetNextWindow(CurrHWND, GW_HWNDNEXT)
        If CheckClass(CurrHWND) = ("IEFrame") And GetRealText(CurrHWND) <> ("Cannot find server") Then MainForm.lstOpen.AddItem GetRealText(CurrHWND)
    Wend
End Function

Function InBlacklist(BlackItem As String) As Boolean
    On Error GoTo KillMe
    Dim Checkloop As Integer
    For Checkloop = 1 To UBound(Blacklist)
        If KillMethod = 1 Then
            If Blacklist(Checkloop) = BlackItem Then
                InBlacklist = True
                Exit Function
            End If
        Else
            If InStr(BlackItem, Blacklist(Checkloop)) Then
                InBlacklist = True
                Exit Function
            End If
        End If
    Next Checkloop
    InBlacklist = False
    Exit Function
KillMe:
    Exit Function
End Function

Function CheckClass(BlackHWND As Long) As String
    Dim Buffer
    Dim TextLength As Integer
    Dim ClassName As String
    ClassName = String(100, Chr(0))
    Buffer = GetClassName(BlackHWND, ClassName, 100)
    ClassName = Left(ClassName, Buffer)
    CheckClass = ClassName
End Function

Function KillWindows()
    Dim CurrHWND As Long
    CurrHWND = GetNextWindow(MainForm.hwnd, GW_HWNDFIRST)
    If CheckClass(CurrHWND) = ("IEFrame") And InBlacklist(GetRealText(CurrHWND)) = True Then
        PopCount = PopCount + 1
        Call SaveSetting(App.EXEName, "Settings", "Pops", PopCount)
        Call PostMessage(CurrHWND, WM_CLOSE, 0&, 0&)
        If PlaySound = True Then WAVPlay (SoundPath)
        MainForm.Frame1.Caption = (".Black List. -- You've Eliminated " & PopCount)
    ElseIf CheckClass(CurrHWND) = ("IEFrame") And StereotypeWindows = True Then
        If NoMaxButton = 1 Then
            If HasMax(CurrHWND) = False Then
                PopCount = PopCount + 1
                Call SaveSetting(App.EXEName, "Settings", "Pops", PopCount)
                Call PostMessage(CurrHWND, WM_CLOSE, 0&, 0&)
                If PlaySound = True Then WAVPlay (SoundPath)
                MainForm.Frame1.Caption = (".Black List. -- You've Eliminated " & PopCount)
            End If
        ElseIf NoMinButton = 1 Then
            If HasMin(CurrHWND) = False Then
                PopCount = PopCount + 1
                Call SaveSetting(App.EXEName, "Settings", "Pops", PopCount)
                Call PostMessage(CurrHWND, WM_CLOSE, 0&, 0&)
                If PlaySound = True Then WAVPlay (SoundPath)
                MainForm.Frame1.Caption = (".Black List. -- You've Eliminated " & PopCount)
            End If
        ElseIf NoTitleBar = 1 Then
            If HasTitleBar(CurrHWND) = False Then
                PopCount = PopCount + 1
                Call SaveSetting(App.EXEName, "Settings", "Pops", PopCount)
                Call PostMessage(CurrHWND, WM_CLOSE, 0&, 0&)
                If PlaySound = True Then WAVPlay (SoundPath)
                MainForm.Frame1.Caption = (".Black List. -- You've Eliminated " & PopCount)
            End If
        End If
    End If
                
    While CurrHWND <> 0
        CurrHWND = GetNextWindow(CurrHWND, GW_HWNDNEXT)
        If CheckClass(CurrHWND) = ("IEFrame") And InBlacklist(GetRealText(CurrHWND)) = True Then
            Call PostMessage(CurrHWND, WM_CLOSE, 0&, 0&)
            If PlaySound = True Then WAVPlay (SoundPath)
            PopCount = PopCount + 1
            Call SaveSetting(App.EXEName, "Settings", "Pops", PopCount)
            MainForm.Frame1.Caption = (".Black List. -- You've Eliminated " & PopCount)
        ElseIf CheckClass(CurrHWND) = ("IEFrame") And StereotypeWindows = True Then
            If NoMaxButton = 1 Then
                If HasMax(CurrHWND) = False Then
                    PopCount = PopCount + 1
                    Call SaveSetting(App.EXEName, "Settings", "Pops", PopCount)
                    Call PostMessage(CurrHWND, WM_CLOSE, 0&, 0&)
                    If PlaySound = True Then WAVPlay (SoundPath)
                    MainForm.Frame1.Caption = (".Black List. -- You've Eliminated " & PopCount)
                End If
            ElseIf NoMinButton = 1 Then
                If HasMin(CurrHWND) = False Then
                    PopCount = PopCount + 1
                    Call SaveSetting(App.EXEName, "Settings", "Pops", PopCount)
                    Call PostMessage(CurrHWND, WM_CLOSE, 0&, 0&)
                    If PlaySound = True Then WAVPlay (SoundPath)
                    MainForm.Frame1.Caption = (".Black List. -- You've Eliminated " & PopCount)
            End If
        ElseIf NoTitleBar = 1 Then
            If HasTitleBar(CurrHWND) = False Then
                    PopCount = PopCount + 1
                    Call SaveSetting(App.EXEName, "Settings", "Pops", PopCount)
                    Call PostMessage(CurrHWND, WM_CLOSE, 0&, 0&)
                    If PlaySound = True Then WAVPlay (SoundPath)
                    MainForm.Frame1.Caption = (".Black List. -- You've Eliminated " & PopCount)
                End If
        End If
    End If
    Wend
End Function


