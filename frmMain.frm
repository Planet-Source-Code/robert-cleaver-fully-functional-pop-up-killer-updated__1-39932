VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Popup Eliminator"
   ClientHeight    =   3840
   ClientLeft      =   5775
   ClientTop       =   3270
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":030A
   ScaleHeight     =   3840
   ScaleWidth      =   4800
   Begin VB.Frame Frame2 
      Caption         =   ".Open IE Windows."
      Height          =   1680
      Left            =   75
      TabIndex        =   2
      Top             =   1635
      Width           =   4665
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   285
         Left            =   75
         TabIndex        =   9
         Top             =   1320
         Width           =   4485
      End
      Begin VB.ListBox lstOpen 
         Height          =   1035
         Left            =   75
         TabIndex        =   3
         Top             =   225
         Width           =   4500
      End
   End
   Begin VB.Timer TimeKILL 
      Interval        =   500
      Left            =   210
      Top             =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   ".Black List."
      Height          =   1545
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4665
      Begin VB.ListBox lstBlack 
         Height          =   1230
         Left            =   75
         TabIndex        =   1
         Top             =   225
         Width           =   4500
      End
   End
   Begin VB.Frame Frame3 
      Height          =   540
      Left            =   75
      TabIndex        =   4
      Top             =   3285
      Width           =   4665
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide"
         Height          =   285
         Left            =   3705
         TabIndex        =   5
         Top             =   165
         Width           =   840
      End
      Begin VB.CommandButton cmdAddCustom 
         Caption         =   "Add Custom"
         Height          =   285
         Left            =   2610
         TabIndex        =   10
         Top             =   165
         Width           =   1110
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   285
         Left            =   1755
         TabIndex        =   6
         Top             =   165
         Width           =   870
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   285
         Left            =   930
         TabIndex        =   7
         Top             =   165
         Width           =   840
      End
      Begin VB.CommandButton cmdOps 
         Caption         =   "Options"
         Height          =   285
         Left            =   105
         TabIndex        =   8
         Top             =   165
         Width           =   840
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About Author"
      End
      Begin VB.Menu mnuFileSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuBlacklist 
      Caption         =   "Blacklist"
      Begin VB.Menu mnuAddBlack 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuAddCustom 
         Caption         =   "Add Custom"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuAddDelPop 
         Caption         =   "Add/Delete Popups"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsTray 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuSpacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutTray 
         Caption         =   "About Author"
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseTray 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Call LoadBlackList
    Call DisplayIEWindows
End Sub

Private Sub cmdAdd_Click()
    If lstOpen.ListCount > 0 And lstOpen.Text <> ("") Then
        AddBlackItem (lstOpen.Text)
        lstOpen.RemoveItem lstOpen.ListIndex
        LoadBlackList
    End If
End Sub

Private Sub cmdAddCustom_Click()
    Dim AddString As String
    AddString = InputBox("Enter the Title of the Website" & vbNewLine & "ex:Click Here!", "Add Custom")
    If AddString = ("") Then Exit Sub
    AddBlackItem (AddString)
    LoadBlackList
End Sub

Private Sub cmdDelete_Click()
    If lstBlack.ListIndex >= 0 Then DeleteBlackItem (lstBlack.Text)
End Sub

Private Sub cmdHide_Click()
    Me.Visible = False
End Sub

Private Sub cmdOps_Click()
    MainOptions.Show
End Sub

Private Sub cmdRefresh_Click()
    Call DisplayIEWindows
End Sub

Private Sub Form_Load()
    StereotypeWindows = Val(GetSetting(App.EXEName, "Settings", "StereotypeWindows"))
    NoMaxButton = Val(GetSetting(App.EXEName, "Settings", "NoMaxButton"))
    NoMinButton = Val(GetSetting(App.EXEName, "Settings", "NoMinButton"))
    NoTitleBar = Val(GetSetting(App.EXEName, "Settings", "NoTitleBar"))
    If Val(GetSetting(App.EXEName, "Settings", "PlaySound")) = 1 Then
        PlaySound = True
    Else
        PlaySound = False
    End If
    If Val(GetSetting(App.EXEName, "Settings", "Method")) = 1 Then
        KillMethod = 1
    Else
        KillMethod = 0
    End If
    SoundPath = GetSetting(App.EXEName, "Settings", "SoundPath")
    If Len(SoundPath) < 4 Then PlaySound = False
    LoadBlackList
    Call DisplayIEWindows
    With nidProgramData
             .cbSize = Len(nidProgramData)                  'Lenght of Data To Be Added
             .hwnd = Me.hwnd                                'The hWnd of the Form In Control Of the Menu
             .uId = vbNull                                  'ID is NULL
             .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE   'Flags of Data To Be Added
             .uCallbackMessage = WM_MOUSEMOVE               'Message For The MEnu Activation
             .hIcon = Me.Icon                               'Icon to be Shown
             .szTip = "The Mouse Over Text :)" & vbNullChar 'Mouse Over Text
    End With
    Shell_NotifyIcon NIM_ADD, nidProgramData           'Add The Icon
    PopCount = Val(GetSetting(App.EXEName, "Settings", "Pops"))
    MainForm.Frame1.Caption = (".Black List. -- You've Eliminated " & PopCount)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Form_MouseMove_err:           'Error Handler
    Dim Result, MSG         As Long
    If Me.ScaleMode = vbPixels Then             'This gets the Message From The Mouse Over Event
        MSG = X
    Else
        MSG = X / Screen.TwipsPerPixelX
    End If
    Select Case MSG
        Case WM_LBUTTONUP                       'This is what it does if they click the left button
            Me.Show                             'on it, and this is what happens on the left button mouse up
            Me.WindowState = 0
        Case WM_LBUTTONDBLCLK
                                                'If ya wan it to do somethin on the left click, then add it here :)
        Case WM_RBUTTONUP
            PopupMenu mnuOptions                    'Code For Right Mouse Button Up
    End Select
Form_MouseMove_err:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nidProgramData 'Remove The Icon From Memory And Taskbar
End Sub

Private Sub lstBlack_DblClick()
If lstBlack.ListCount > 0 And lstBlack.ListIndex >= 0 Then DeleteBlackItem (lstBlack.Text)
End Sub

Private Sub lstOpen_DblClick()
    If lstOpen.ListCount > 0 And InBlacklist(lstOpen.Text) = False And lstOpen.ListIndex >= 0 Then
        AddBlackItem (lstOpen.Text)
        lstOpen.RemoveItem lstOpen.ListIndex
        LoadBlackList
    End If
End Sub

Private Sub mnuAbout_Click()
    MainAbout.Show
End Sub

Private Sub mnuAboutTray_Click()
    MainAbout.Show
End Sub

Private Sub mnuAddDelPop_Click()
    Me.Visible = True
End Sub

Private Sub mnuClose_Click()
    Shell_NotifyIcon NIM_DELETE, nidProgramData 'Remove The Icon From Memory And Taskbar
    End
End Sub

Private Sub mnuCloseTray_Click()
    Shell_NotifyIcon NIM_DELETE, nidProgramData 'Remove The Icon From Memory And Taskbar
    End
End Sub

Private Sub mnuDelete_Click()
    If lstBlack.ListIndex >= 0 Then DeleteBlackItem (lstBlack.Text)
End Sub

Private Sub mnuMinimize_Click()
    Me.WindowState = 1
End Sub

Private Sub mnuOptionsTray_Click()
    MainOptions.Show
End Sub

Private Sub TimeKILL_Timer()
    Call KillWindows
End Sub
