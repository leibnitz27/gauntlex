Attribute VB_Name = "modAssets"
Option Explicit

' ============================================================
'  Embedded tile assets.
'
'  Shape.Fill.UserPicture takes a file PATH, not bytes, so the
'  renderer can't paint straight from the workbook.  Build-
'  Workbook.ps1 stamps every reference\genesis-tiles\*.png onto
'  the very-hidden "Tiles" sheet as base64 (col A = filename,
'  col B = data; A1/B1 = "__ver__" / a version stamp).
'
'  On the first run we decode them once into
'      %TEMP%\gauntlex\tiles\
'  and drop a  _ver_<stamp>  sentinel.  A later build with a new
'  stamp re-extracts; an unchanged build reuses the folder.
'
'  This is what lets gauntlex.xlsm ship as a single file - the
'  reference\ folder is build-time only.
' ============================================================

Private Const TILES_SHEET As String = "Tiles"
Private mDir As String

' Folder holding the extracted tiles for the current build, with a
' trailing "\".  Extracts on the first call, cached thereafter.
Public Function TilesDir() As String
    If Len(mDir) > 0 Then TilesDir = mDir: Exit Function

    Dim root As String
    root = Environ$("TEMP")
    If Len(root) = 0 Then root = Environ$("TMP")
    If Len(root) = 0 Then root = ThisWorkbook.Path
    Dim base As String: base = root & "\gauntlex\tiles\"

    Dim ver As String: ver = TileVersion()
    If Dir$(base & "_ver_" & ver) = "" Then ExtractTiles base, ver

    mDir = base
    TilesDir = mDir
End Function

Private Function TileVersion() As String
    TileVersion = CStr(TilesSheet().Range("B1").Value)
    If Len(TileVersion) = 0 Then TileVersion = "0"
End Function

Private Function TilesSheet() As Worksheet
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If s.Name = TILES_SHEET Then Set TilesSheet = s: Exit Function
    Next s
    Err.Raise vbObjectError + 513, "modAssets", _
        "workbook has no '" & TILES_SHEET & "' sheet - rebuild with build\Build-Workbook.ps1"
End Function

Private Sub ExtractTiles(ByVal base As String, ByVal ver As String)
    Dim ws As Worksheet: Set ws = TilesSheet()

    Dim prevSU As Boolean: prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.StatusBar = "Gauntlex: unpacking tiles (first run)..."

    EnsureEmptyFolder base

    Dim last As Long: last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Err.Raise vbObjectError + 514, "modAssets", _
        "'" & TILES_SHEET & "' sheet holds no tiles - rebuild"
    Dim data As Variant: data = ws.Range("A2:B" & last).Value   ' 1-based (row, col)

    Dim node As Object, stm As Object
    Set node = CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
    node.DataType = "bin.base64"
    Set stm = CreateObject("ADODB.Stream")

    Dim i As Long, nm As String, b64 As String
    For i = 1 To UBound(data, 1)
        nm = Trim$(CStr(data(i, 1)))
        b64 = CStr(data(i, 2))
        If Len(nm) > 0 And Len(b64) > 0 Then
            node.Text = b64
            stm.Type = 1                        ' adTypeBinary
            stm.Open
            stm.Write node.nodeTypedValue
            stm.SaveToFile base & nm, 2         ' adSaveCreateOverWrite
            stm.Close
        End If
    Next i

    WriteMarker base & "_ver_" & ver

    Application.StatusBar = False
    Application.ScreenUpdating = prevSU
End Sub

' Create <root>\gauntlex\tiles\ (its grandparent %TEMP% is assumed to
' exist) and clear any tiles left by an earlier build.
Private Sub EnsureEmptyFolder(ByVal folder As String)
    Dim d As String: d = Left$(folder, Len(folder) - 1)        ' drop trailing "\"
    Dim parent As String: parent = Left$(d, InStrRev(d, "\") - 1)
    If Dir$(parent, vbDirectory) = "" Then MkDir parent
    If Dir$(d, vbDirectory) = "" Then
        MkDir d
    Else
        On Error Resume Next
        Kill d & "\*"
        On Error GoTo 0
    End If
End Sub

Private Sub WriteMarker(ByVal path As String)
    Dim f As Integer: f = FreeFile
    Open path For Output As #f
    Print #f, "gauntlex tiles unpacked " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #f
End Sub
