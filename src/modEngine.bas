Attribute VB_Name = "modEngine"
Option Explicit

' ============================================================
'  Engine: the frame loop.
'
'  VBA has no game loop, so this is it: poll -> update ->
'  render -> DoEvents -> Sleep the remainder of the frame.
'  Arrow keys are trapped for the loop's lifetime so they
'  drive the player instead of the Excel selection; the
'  physical key state is still visible to GetAsyncKeyState.
' ============================================================

Private mRunning As Boolean

Public Sub StartGauntlex()
    If mRunning Then Exit Sub
    mRunning = True

    On Error GoTo Cleanup
    Application.EnableEvents = False
    TrapArrows True

    GameInit
    RenderInit

    Dim tPrev As Long, tNow As Long, dt As Long, spent As Long
    Dim frames As Long, fpsClock As Long, fps As Double
    tPrev = timeGetTime()
    fpsClock = tPrev

    Do While mRunning
        tNow = timeGetTime()
        dt = tNow - tPrev
        If dt < 0 Then dt = 0
        tPrev = tNow

        PollInput
        If gInQuit Then mRunning = False

        GameUpdate dt
        RenderFrame fps

        frames = frames + 1
        If tNow - fpsClock >= 1000 Then
            fps = frames * 1000# / (tNow - fpsClock)
            frames = 0
            fpsClock = tNow
        End If

        DoEvents
        spent = timeGetTime() - tNow
        If spent < FRAME_MS Then Sleep FRAME_MS - spent
    Loop

Cleanup:
    Dim n As Long, d As String
    n = Err.Number: d = Err.Description
    TrapArrows False
    Application.EnableEvents = True
    mRunning = False
    If n <> 0 Then MsgBox "Gauntlex halted." & vbCrLf & "Error " & n & ": " & d, vbExclamation
End Sub

Public Sub StopGauntlex()
    mRunning = False
End Sub

Private Sub TrapArrows(ByVal enable As Boolean)
    Dim k As Variant
    For Each k In Array("{UP}", "{DOWN}", "{LEFT}", "{RIGHT}")
        If enable Then
            Application.OnKey CStr(k), ""      ' swallow
        Else
            Application.OnKey CStr(k)          ' restore default
        End If
    Next k
End Sub
