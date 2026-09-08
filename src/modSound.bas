Attribute VB_Name = "modSound"
Option Explicit

' ============================================================
'  Sound: async speech call-outs, Gauntlet-style ("Warrior
'  needs food badly"). Application.Speech is fire-and-forget
'  and may be absent - every call is guarded.
' ============================================================

Public gSoundOn As Boolean

Public Sub SoundInit()
    gSoundOn = True
    On Error Resume Next
    Application.Speech.Speak "", SpeakAsync:=True     ' warm the engine
    On Error GoTo 0
End Sub

Public Sub Say(ByVal text As String)
    If Not gSoundOn Then Exit Sub
    On Error Resume Next
    Application.Speech.Speak text, SpeakAsync:=True
    On Error GoTo 0
End Sub

Public Sub SoundToggle()
    gSoundOn = Not gSoundOn
End Sub
