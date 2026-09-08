Attribute VB_Name = "spikeRender"
Option Explicit

' ============================================================
'  RENDERING SPIKE - throwaway. Run VISIBLY (headless has no
'  repaint, and hidden Excel crashes on picture fills).
'
'  Open spike.xlsm, click the buttons on the Spike sheet.
'  Each runs ~6s and writes "avg N  min N fps" to cell A(VR+4).
'  ESC aborts a run.
'
'    Test1  background only  - value blit + recolour scrolled strip
'    Test2  40 shapes        - reposition every frame, no pic change
'    Test3  40 shapes + anim - + swap .Fill.UserPicture every 6 frames
'    Test4  combined         - background + 40 shapes repositioned
' ============================================================

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

Private Const VC As Long = 36
Private Const VR As Long = 20
Private Const CELL As Double = 22#
Private Const SH As String = "Spike"
Private Const POOL As Long = 40
Private Const MZC As Long = 18          ' maze grid, blocks (Test5)
Private Const MZR As Long = 10
Private Const VK_ESC As Long = &H1B

Private mShp() As Shape
Private mSX() As Double, mSY() As Double, mDX() As Double, mDY() As Double
Private mFrames(0 To 3) As String
Private mMz() As Shape
Private mWallPic As String, mFloorPic As String
Private mVal(1 To VR, 1 To VC) As String
Private mBg(1 To VR, 1 To VC) As Long
Private mBuilt As Boolean

Public Sub Test1_Background():   RunLoop 1: End Sub
Public Sub Test2_Shapes():       RunLoop 2: End Sub
Public Sub Test3_ShapesAnim():   RunLoop 3: End Sub
Public Sub Test4_Combined():     RunLoop 4: End Sub
Public Sub Test5_ShapeMaze():    RunLoop 5: End Sub    ' 180 fixed maze shapes + 40 actors

Private Function Sheet() As Worksheet
    On Error Resume Next
    Set Sheet = ThisWorkbook.Worksheets(SH)
    On Error GoTo 0
    If Sheet Is Nothing Then
        Set Sheet = ThisWorkbook.Worksheets.Add
        Sheet.Name = SH
    End If
End Function

Private Sub EnsureBuilt()
    If mBuilt Then Exit Sub
    Dim ws As Worksheet: Set ws = Sheet()
    ws.Activate
    ws.Cells.Clear
    ws.Cells.Font.Name = "Consolas"
    ws.Cells.Font.Size = 11
    ws.Rows("1:" & VR).RowHeight = CELL
    SquareCols ws

    Dim r As Long, c As Long, wall As Boolean
    For r = 1 To VR
        For c = 1 To VC
            wall = (r = 1 Or r = VR Or c = 1 Or c = VC Or (r Mod 6 = 0) Or (c Mod 9 = 0))
            mVal(r, c) = IIf(wall, ChrW(9608), "")
            mBg(r, c) = IIf(wall, RGB(120, 78, 48), RGB(38, 38, 44))
        Next c
    Next r
    ApplyBackground ws

    Dim base As String: base = ThisWorkbook.Path & "\reference\genesis-tiles\"
    mFrames(0) = base & "r09_c00.png": mFrames(1) = base & "r09_c01.png"
    mFrames(2) = base & "r09_c02.png": mFrames(3) = base & "r09_c03.png"
    mWallPic = base & "r12_c20.png": mFloorPic = base & "r13_c00.png"
    mBuilt = True
End Sub

' 180 block-sized picture shapes, cell-aligned, never repositioned.
Private Sub MakeMaze(ByVal ws As Worksheet)
    Dim s As Shape
    For Each s In ws.Shapes
        If Left$(s.Name, 2) = "mz" Then s.Delete
    Next s
    ReDim mMz(1 To MZC * MZR)
    Dim bp As Double: bp = CELL * 2
    Dim br As Long, bc As Long, i As Long, wall As Boolean
    For br = 1 To MZR
        For bc = 1 To MZC
            i = (br - 1) * MZC + bc
            wall = (br = 1 Or br = MZR Or bc = 1 Or bc = MZC Or (br Mod 4 = 0) Or (bc Mod 6 = 0))
            Set mMz(i) = ws.Shapes.AddShape(msoShapeRectangle, (bc - 1) * bp, (br - 1) * bp, bp, bp)
            mMz(i).Name = "mz" & i
            mMz(i).Line.Visible = msoFalse
            mMz(i).Fill.UserPicture IIf(wall, mWallPic, mFloorPic)
        Next bc
    Next br
End Sub

Private Sub DeleteMaze(ByVal ws As Worksheet)
    Dim s As Shape
    For Each s In ws.Shapes
        If Left$(s.Name, 2) = "mz" Then s.Delete
    Next s
End Sub

' ColumnWidth is in "characters", we want CELL points. Bounded binary search
' (Excel quantises .Width, so never loop on an exact match).
Private Sub SquareCols(ByVal ws As Worksheet)
    Dim lo As Double, hi As Double, mid As Double, i As Long
    lo = 0.25: hi = 60#
    For i = 1 To 30
        mid = (lo + hi) / 2
        ws.Columns(1).ColumnWidth = mid
        If ws.Columns(1).Width > CELL Then hi = mid Else lo = mid
    Next i
    ws.Range(ws.Columns(1), ws.Columns(VC)).ColumnWidth = ws.Columns(1).ColumnWidth
End Sub

Private Sub ApplyBackground(ByVal ws As Worksheet)
    ws.Range(ws.Cells(1, 1), ws.Cells(VR, VC)).Value = mVal
    ws.Range(ws.Cells(1, 1), ws.Cells(VR, VC)).Font.Color = RGB(120, 78, 48)
    Dim r As Long, c As Long
    For r = 1 To VR
        For c = 1 To VC
            ws.Cells(r, c).Interior.Color = mBg(r, c)
        Next c
    Next r
End Sub

Private Sub MakePool(ByVal ws As Worksheet)
    Dim s As Shape
    For Each s In ws.Shapes
        If Left$(s.Name, 3) = "spr" Then s.Delete
    Next s
    ReDim mShp(1 To POOL): ReDim mSX(1 To POOL): ReDim mSY(1 To POOL)
    ReDim mDX(1 To POOL): ReDim mDY(1 To POOL)
    Dim i As Long
    For i = 1 To POOL
        Set mShp(i) = ws.Shapes.AddShape(msoShapeRectangle, _
            (i * 13) Mod CLng((VC - 2) * CELL), (i * 7) Mod CLng((VR - 2) * CELL), CELL, CELL)
        mShp(i).Name = "spr" & i
        mShp(i).Line.Visible = msoFalse
        mShp(i).Fill.UserPicture mFrames(i Mod 4)
        mSX(i) = mShp(i).Left: mSY(i) = mShp(i).Top
        mDX(i) = (((i Mod 5) - 2) + 0.5) * 2.3
        mDY(i) = ((((i \ 3) Mod 5) - 2) + 0.5) * 2.3
    Next i
End Sub

Private Sub MovePool()
    Dim i As Long
    For i = 1 To POOL
        mSX(i) = mSX(i) + mDX(i): mSY(i) = mSY(i) + mDY(i)
        If mSX(i) < 0 Or mSX(i) > (VC - 1) * CELL Then mDX(i) = -mDX(i): mSX(i) = mSX(i) + mDX(i)
        If mSY(i) < 0 Or mSY(i) > (VR - 1) * CELL Then mDY(i) = -mDY(i): mSY(i) = mSY(i) + mDY(i)
        mShp(i).Left = mSX(i): mShp(i).Top = mSY(i)
    Next i
End Sub

Private Sub RunLoop(ByVal mode As Long)
    EnsureBuilt
    Dim ws As Worksheet: Set ws = Sheet()
    Dim useShapes As Boolean: useShapes = (mode >= 2)
    Dim doBg As Boolean:      doBg = (mode = 1 Or mode = 4)
    Dim doAnim As Boolean:    doAnim = (mode = 3)
    Dim doMaze As Boolean:    doMaze = (mode = 5)

    If useShapes Then MakePool ws Else DeletePool ws
    If doMaze Then MakeMaze ws Else DeleteMaze ws
    ws.Cells(VR + 4, 1).Value = "running test " & mode & " ... (ESC aborts)"

    Dim t0 As Long, tNow As Long, tPrev As Long, spent As Long
    Dim fr As Long, bucketFrames As Long, bucketStart As Long
    Dim minFps As Double: minFps = 1E+09
    t0 = timeGetTime(): tPrev = t0: bucketStart = t0
    Dim scrollCol As Long: scrollCol = 1

    Do While timeGetTime() - t0 < 6000
        tNow = timeGetTime()

        If doBg Then
            ws.Range(ws.Cells(1, 1), ws.Cells(VR, VC)).Value = mVal
            scrollCol = (scrollCol Mod VC) + 1
            Dim r As Long
            For r = 1 To VR
                ws.Cells(r, scrollCol).Interior.Color = mBg(r, scrollCol)
            Next r
        End If
        If useShapes Then MovePool
        If doAnim And (fr Mod 6 = 0) Then
            Dim i As Long
            For i = 1 To POOL: mShp(i).Fill.UserPicture mFrames((fr \ 6 + i) Mod 4): Next i
        End If
        If doMaze And (fr Mod 4 = 0) Then
            ' simulate a camera scroll step: re-texture one block column (~10 shapes)
            Dim ec As Long, br As Long, mi As Long
            ec = ((fr \ 4) Mod MZC) + 1
            For br = 1 To MZR
                mi = (br - 1) * MZC + ec
                mMz(mi).Fill.UserPicture IIf(((fr \ 4) Mod 2) = 0, mWallPic, mFloorPic)
            Next br
        End If

        fr = fr + 1: bucketFrames = bucketFrames + 1
        If tNow - bucketStart >= 250 Then
            Dim bf As Double: bf = bucketFrames * 1000# / (tNow - bucketStart)
            If bf < minFps Then minFps = bf
            bucketFrames = 0: bucketStart = tNow
        End If

        DoEvents
        If (GetAsyncKeyState(VK_ESC) And &H8000) <> 0 Then Exit Do
        spent = timeGetTime() - tNow
        If spent < 20 Then Sleep 20 - spent      ' cap ~50fps so we measure headroom
    Loop

    Dim secs As Double: secs = (timeGetTime() - t0) / 1000#
    ws.Cells(VR + 4, 1).Value = "test " & mode & ":  avg " & Format$(fr / secs, "0.0") & _
        "   min " & Format$(minFps, "0.0") & " fps   (" & fr & " frames / " & Format$(secs, "0.0") & "s)"
End Sub

Private Sub DeletePool(ByVal ws As Worksheet)
    Dim s As Shape
    For Each s In ws.Shapes
        If Left$(s.Name, 3) = "spr" Then s.Delete
    Next s
End Sub
