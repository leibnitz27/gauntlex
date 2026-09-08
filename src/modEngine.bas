Attribute VB_Name = "modEngine"
Option Explicit

' ============================================================
'  Engine: the frame loop.
'
'  VBA has no game loop, so this is it: poll -> update ->
'  render -> DoEvents -> Sleep the remainder of the frame.
'  Keys that would edit / move the cell selection (arrows,
'  space, enter, tab, F2, delete, backspace) are swallowed for
'  the loop's lifetime - otherwise e.g. space opens cell edit
'  mode and the next Range write throws. GetAsyncKeyState still
'  sees the physical keyboard regardless.
' ============================================================

Private mRunning As Boolean

Public Sub StartGauntlex()
    If mRunning Then Exit Sub
    mRunning = True

    On Error GoTo Cleanup
    Application.EnableEvents = False
    TrapKeys True
    On Error Resume Next
    Application.Interactive = False     ' Excel ignores keyboard/mouse; we poll GetAsyncKeyState directly
    On Error GoTo Cleanup
    SetPlayButton False
    SoundInit

    FitViewport         ' size the view to the Excel window (capped at the map)
    RenderInit
    gState = "TITLE"

    Dim tPrev As Long, tNow As Long, dt As Long, spent As Long
    Dim frames As Long, fpsClock As Long, fps As Double
    Dim resultAt As Long, cursor As Long
    Dim firePrev As Boolean, prevPrev As Boolean, nextPrev As Boolean
    Dim fireEdge As Boolean, prevEdge As Boolean, nextEdge As Boolean
    tPrev = timeGetTime()
    fpsClock = tPrev

    Do While mRunning
        tNow = timeGetTime()
        dt = tNow - tPrev
        If dt < 0 Then dt = 0
        tPrev = tNow

        PollInput
        If gInQuit Then mRunning = False
        ' menu nav - up OR left = previous, down OR right = next (edge-triggered)
        Dim prevHeld As Boolean, nextHeld As Boolean
        prevHeld = gInUp Or gInLeft
        nextHeld = gInDown Or gInRight
        fireEdge = gInFire And Not firePrev: firePrev = gInFire
        prevEdge = prevHeld And Not prevPrev: prevPrev = prevHeld
        nextEdge = nextHeld And Not nextPrev: nextPrev = nextHeld

        Select Case gState
            Case "TITLE"
                RenderTitle
                If fireEdge Then gState = "SELECT": cursor = gChar

            Case "SELECT"
                If prevEdge Then cursor = (cursor + CHAR_COUNT - 1) Mod CHAR_COUNT
                If nextEdge Then cursor = (cursor + 1) Mod CHAR_COUNT
                RenderSelect cursor
                If fireEdge Then
                    gChar = cursor
                    GameInit
                    CenterCamera
                    gState = "PLAY"
                End If

            Case "PLAY"
                GameUpdate dt
                RenderFrame fps

            Case "WON"                          ' level cleared -> next, or victory
                RenderFrame fps
                If resultAt = 0 Then resultAt = tNow + 2000
                If tNow >= resultAt Then
                    resultAt = 0
                    If AdvanceLevel() Then gState = "PLAY" Else gState = "VICTORY"
                End If

            Case "VICTORY"
                RenderVictory
                If resultAt = 0 Then resultAt = tNow + 4500: Say "You have escaped the dungeon"
                If tNow >= resultAt Then gState = "TITLE": resultAt = 0

            Case "OVER"
                RenderFrame fps
                If resultAt = 0 Then resultAt = tNow + 2400
                If tNow >= resultAt Then gState = "TITLE": resultAt = 0
        End Select

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
    On Error Resume Next
    Application.Interactive = True
    TrapKeys False
    SetPlayButton True
    Application.EnableEvents = True
    mRunning = False
    On Error GoTo 0
    If n <> 0 Then MsgBox "Gauntlex halted." & vbCrLf & "Error " & n & ": " & d, vbExclamation
End Sub

' Hidden while the loop runs; shown again on exit so it also means "restart".
Private Sub SetPlayButton(ByVal visible As Boolean)
    On Error Resume Next
    ThisWorkbook.Worksheets(SCREEN_SHEET).Shapes("btnPlay").Visible = visible
    On Error GoTo 0
End Sub

Public Sub StopGauntlex()
    mRunning = False
End Sub

Private Sub TrapKeys(ByVal enable As Boolean)
    Dim k As Variant
    For Each k In Array("{UP}", "{DOWN}", "{LEFT}", "{RIGHT}", _
                       " ", "~", "{ENTER}", "{TAB}", "{F2}", "{DELETE}", "{BS}")
        On Error Resume Next                   ' some tokens vary by Excel build
        If enable Then
            Application.OnKey CStr(k), ""      ' swallow
        Else
            Application.OnKey CStr(k)          ' restore default
        End If
        On Error GoTo 0
    Next k
End Sub
