VERSION 5.00
Begin VB.Form MainAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About the Author"
   ClientHeight    =   2520
   ClientLeft      =   1560
   ClientTop       =   1560
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5055
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2220
      Width           =   2430
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   90
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   1770
      ScaleWidth      =   2370
      TabIndex        =   0
      Top             =   135
      Width           =   2400
   End
   Begin VB.Label Label4 
      Caption         =   $"frmAbout.frx":E090
      Height          =   1800
      Left            =   2550
      TabIndex        =   4
      Top             =   120
      Width           =   2400
   End
   Begin VB.Label Label3 
      Caption         =   "Marital Status: Married"
      Height          =   225
      Left            =   2550
      TabIndex        =   3
      Top             =   1995
      Width           =   2400
   End
   Begin VB.Label Label2 
      Caption         =   "Age: 16"
      Height          =   225
      Left            =   75
      TabIndex        =   2
      Top             =   2235
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Name: Robert Cleaver"
      Height          =   225
      Left            =   90
      TabIndex        =   1
      Top             =   1995
      Width           =   2400
   End
End
Attribute VB_Name = "MainAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

