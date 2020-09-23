Attribute VB_Name = "Sound"
Option Explicit

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found

Function WAVStop()
    Call WAVPlay(" ")
End Function

Function WAVLoop(File)
    Dim SoundName As String
    Dim wFlags As Integer
    Dim X
    SoundName$ = File
    wFlags% = SND_ASYNC Or SND_LOOP
    X = sndPlaySound(SoundName$, wFlags%)
End Function

Function WAVPlay(File)
    Dim SoundName As String
    Dim wFlags As Integer
    Dim X
    SoundName$ = File
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X = sndPlaySound(SoundName$, wFlags%)
End Function
