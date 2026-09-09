Attribute VB_Name = "modInput"
Option Explicit

' ============================================================
'  Input: poll once per frame into intent flags.
'  The rest of the engine reads these, never the keyboard.
' ============================================================

Public gInUp     As Boolean
Public gInDown   As Boolean
Public gInLeft   As Boolean
Public gInRight  As Boolean
Public gInFire   As Boolean
Public gInPotion As Boolean
Public gInRender As Boolean
Public gInQuit   As Boolean

Public Sub PollInput()
    gInUp = KeyDown(VK_UP)
    gInDown = KeyDown(VK_DOWN)
    gInLeft = KeyDown(VK_LEFT)
    gInRight = KeyDown(VK_RIGHT)
    gInFire = KeyDown(VK_SPACE)
    gInPotion = KeyDown(VK_C)
    gInRender = KeyDown(VK_G)
    gInQuit = KeyDown(VK_ESCAPE)
End Sub
