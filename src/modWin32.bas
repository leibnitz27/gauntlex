Attribute VB_Name = "modWin32"
Option Explicit

' ============================================================
'  Win32 glue: keyboard polling + millisecond timing + sleep
'  (Windows desktop Excel only.)
' ============================================================

#If VBA7 Then
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

Public Const VK_LEFT   As Long = &H25
Public Const VK_UP     As Long = &H26
Public Const VK_RIGHT  As Long = &H27
Public Const VK_DOWN   As Long = &H28
Public Const VK_ESCAPE As Long = &H1B
Public Const VK_SPACE  As Long = &H20
Public Const VK_C      As Long = &H43
Public Const VK_G      As Long = &H47      ' toggle glyph / picture-Shape renderer

' True while the physical key is down, regardless of which window has focus.
Public Function KeyDown(ByVal vKey As Long) As Boolean
    KeyDown = (GetAsyncKeyState(vKey) And &H8000) <> 0
End Function
