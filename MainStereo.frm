VERSION 5.00
Begin VB.Form MainStereo 
   Caption         =   "Settings"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1470
   ScaleWidth      =   2760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Close if Any of the Below is True"
      Height          =   870
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2670
      Begin VB.CheckBox chkNoMinButton 
         Caption         =   "No Min Button"
         Height          =   285
         Left            =   1215
         TabIndex        =   3
         Top             =   525
         Width           =   1335
      End
      Begin VB.CheckBox chkNoMaxButton 
         Caption         =   "No Max Button"
         Height          =   285
         Left            =   1215
         TabIndex        =   2
         Top             =   240
         Width           =   1380
      End
      Begin VB.CheckBox chkNoBorder 
         Caption         =   "No TitleBar"
         Height          =   285
         Left            =   105
         TabIndex        =   1
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Height          =   570
      Left            =   60
      TabIndex        =   4
      Top             =   885
      Width           =   2670
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   330
         Left            =   1125
         TabIndex        =   5
         Top             =   165
         Width           =   1470
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   75
         TabIndex        =   6
         Top             =   165
         Width           =   1065
      End
   End
End
Attribute VB_Name = "MainStereo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Call SaveSetting(App.EXEName, "Settings", "NoTitleBar", chkNoBorder.Value)
    Call SaveSetting(App.EXEName, "Settings", "NoMaxButton", chkNoMaxButton.Value)
    Call SaveSetting(App.EXEName, "Settings", "NoMinButton", chkNoMinButton.Value)
    NoTitleBar = chkNoBorder.Value
    NoMaxButton = chkNoMaxButton.Value
    NoMinButton = chkNoMinButton.Value
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If NoMaxButton = 1 Then chkNoMaxButton.Value = 1 Else chkNoMaxButton.Value = 0
    If NoMinButton = 1 Then chkNoMinButton.Value = 1 Else chkNoMinButton.Value = 0
    If NoTitleBar = 1 Then chkNoBorder.Value = 1 Else chkNoBorder.Value = 0
End Sub
