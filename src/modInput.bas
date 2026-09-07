Attribute VB_Name = "modInput"
Option Explicit

' ============================================================
'  Input: poll once per frame into intent flags.
'  The rest of the engine reads these, never the keyboard.
' ============================================================

Public gInUp    As Boolean
Public gInDown  As Boolean
Public gInLeft  As Boolean
Public gInRight As Boolean
Public gInQuit  As Boolean

Public Sub PollInput()
    gInUp = KeyDown(VK_UP)
    gInDown = KeyDown(VK_DOWN)
    gInLeft = KeyDown(VK_LEFT)
    gInRight = KeyDown(VK_RIGHT)
    gInQuit = KeyDown(VK_ESCAPE)
End Sub
