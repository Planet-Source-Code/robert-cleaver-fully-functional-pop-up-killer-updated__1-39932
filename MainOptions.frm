VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MainOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2625
   ClientLeft      =   4035
   ClientTop       =   3960
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Height          =   285
      Left            =   1410
      MaskColor       =   &H00FF00FF&
      Picture         =   "MainOptions.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   270
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   "Stereotype Windows"
      Height          =   555
      Left            =   135
      TabIndex        =   10
      Top             =   1365
      Width           =   4125
      Begin VB.CheckBox chkStereotype 
         Caption         =   "Stereotype"
         Height          =   210
         Left            =   90
         TabIndex        =   13
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton cmdStereoHelp 
         Caption         =   "Help"
         Height          =   285
         Left            =   3375
         TabIndex        =   12
         Top             =   180
         Width           =   630
      End
      Begin VB.CommandButton cmdStereoSettings 
         Caption         =   "Settings"
         Height          =   285
         Left            =   1725
         TabIndex        =   11
         Top             =   180
         Width           =   1665
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   345
      Left            =   1995
      TabIndex        =   1
      Top             =   2100
      Width           =   2160
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Top             =   2100
      Width           =   1770
   End
   Begin VB.Frame Frame2 
      Caption         =   "Method of Window Selection"
      Height          =   615
      Left            =   150
      TabIndex        =   6
      Top             =   690
      Width           =   4125
      Begin MSComDlg.CommonDialog dlgOps 
         Left            =   3630
         Top             =   90
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.OptionButton opKeyWord 
         Caption         =   "Keyword"
         Height          =   240
         Left            =   105
         TabIndex        =   8
         Top             =   270
         Width           =   1155
      End
      Begin VB.OptionButton opExactMatch 
         Caption         =   "Exact Match"
         Height          =   285
         Left            =   1245
         TabIndex        =   7
         Top             =   255
         Width           =   1395
      End
   End
   Begin VB.TextBox txtSound 
      Height          =   255
      Left            =   1785
      TabIndex        =   4
      Top             =   300
      Width           =   1920
   End
   Begin VB.CheckBox chkPlaySound 
      Caption         =   "Play Sound"
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   315
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sound"
      Height          =   555
      Left            =   150
      TabIndex        =   3
      Top             =   90
      Width           =   4110
      Begin VB.CommandButton cmdGetSound 
         Caption         =   "..."
         Height          =   255
         Left            =   3615
         TabIndex        =   5
         Top             =   210
         Width           =   390
      End
   End
   Begin VB.Frame Frame3 
      Height          =   600
      Left            =   120
      TabIndex        =   9
      Top             =   1935
      Width           =   4140
   End
End
Attribute VB_Name = "MainOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPlaySound_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkPlaySound.Value = 0 Then
        txtSound.Enabled = False
        cmdGetSound.Enabled = False
    Else
        txtSound.Enabled = True
        cmdGetSound.Enabled = True
    End If
End Sub

Private Sub chkStereotype_Click()
    If chkStereotype.Value = 0 Then
        cmdStereoSettings.Enabled = False
    Else
        cmdStereoSettings.Enabled = True
    End If
End Sub

Private Sub cmdApply_Click()
    If chkPlaySound.Value = 1 Then
        PlaySound = True
    Else
        PlaySound = False
    End If
    If opKeyWord.Value = True Then
        Call SaveSetting(App.EXEName, "Settings", "Method", 0)
        KillMethod = 0
    Else
        Call SaveSetting(App.EXEName, "Settings", "Method", 1)
        KillMethod = 1
    End If
    If chkStereotype.Value = 1 Then
        StereotypeWindows = 1
    Else
        StereotypeWindows = 0
    End If
    Call SaveSetting(App.EXEName, "Settings", "PlaySound", chkPlaySound.Value)
    Call SaveSetting(App.EXEName, "Settings", "SoundPath", SoundPath)
    Call SaveSetting(App.EXEName, "Settings", "StereotypeWindows", chkStereotype.Value)
    SoundPath = txtSound.Text
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetSound_Click()
    Dim NewSound    As String
    dlgOps.Filter = ("WAV Files | *.wav")
    dlgOps.ShowOpen
    NewSound = dlgOps.FileName
    If NewSound = ("") Then
        Exit Sub
    Else
        txtSound.Text = NewSound
    End If
End Sub

Private Sub cmdStereoHelp_Click()
    Dim HelpMsg As String
    HelpMsg = ("Stereotyping Windows Help" & vbNewLine & _
            "Explanation: " & vbNewLine & _
            "Most Popup-Ads have ways of keeping themselves " & _
            "Open and in your face, and it gets irritating ... a few " & _
            "of these methods would be, no Max/Min-imize Buttons, " & _
            "or No TitleBar. That is why we added this option, you can " & _
            "Stereotype all Open Internet Explorer Windows, and if they " & _
            "Have say, no Max Button, then treat it as a Popup-Ad and " & _
            "close it. Just something to avoid having to add every single " & _
            "Popup-Ad you come across to the blacklist.")
        Call MsgBox(HelpMsg, vbInformation + vbOKOnly, "Help")
End Sub

Private Sub cmdStereoSettings_Click()
    MainStereo.Show vbModal
End Sub

Private Sub Command3_Click()
    Call WAVPlay(txtSound.Text)
End Sub

Private Sub Form_Load()
    If PlaySound = True Then
        chkPlaySound.Value = 1
    Else
        chkPlaySound.Value = 0
    End If
    If KillMethod = 1 Then
        opKeyWord.Value = False
        opExactMatch.Value = True
    Else
        opKeyWord.Value = True
        opExactMatch.Value = False
    End If
    If PlaySound = True Then
        txtSound.Enabled = True
        cmdGetSound.Enabled = True
    Else
        txtSound.Enabled = False
        cmdGetSound.Enabled = False
    End If
    If StereotypeWindows = 1 Then
        cmdStereoSettings.Enabled = True
        chkStereotype.Value = 1
    Else
        cmdStereoSettings.Enabled = False
        chkStereotype.Value = 0
    End If
    txtSound.Text = SoundPath
    If Len(txtSound.Text) < 4 Then txtSound.Text = App.Path & "\Pow.wav"
End Sub
