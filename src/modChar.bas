Attribute VB_Name = "modChar"
Option Explicit

' ============================================================
'  The four heroes. Trade-offs, Gauntlet-style:
'    Warrior  - hard melee, slow, average armour, poor magic
'    Valkyrie - tanky (best armour), balanced
'    Wizard   - huge potion blast, fast shots, very frail
'    Elf      - fastest move + fire, weak everything else
' ============================================================

Public Const CHAR_WARRIOR  As Long = 0
Public Const CHAR_VALKYRIE As Long = 1
Public Const CHAR_WIZARD   As Long = 2
Public Const CHAR_ELF      As Long = 3
Public Const CHAR_COUNT    As Long = 4

Public Function CharName(ByVal c As Long) As String
    Select Case c
        Case CHAR_WARRIOR:  CharName = "WARRIOR"
        Case CHAR_VALKYRIE: CharName = "VALKYRIE"
        Case CHAR_WIZARD:   CharName = "WIZARD"
        Case CHAR_ELF:      CharName = "ELF"
    End Select
End Function

' one half-cell step this often while a key is held (lower = faster)
Public Function CharMoveMs(ByVal c As Long) As Long
    Select Case c
        Case CHAR_WARRIOR:  CharMoveMs = 72
        Case CHAR_VALKYRIE: CharMoveMs = 60
        Case CHAR_WIZARD:   CharMoveMs = 62
        Case CHAR_ELF:      CharMoveMs = 44
    End Select
End Function

' fire cooldown, ms
Public Function CharShotMs(ByVal c As Long) As Long
    Select Case c
        Case CHAR_WARRIOR:  CharShotMs = 260
        Case CHAR_VALKYRIE: CharShotMs = 200
        Case CHAR_WIZARD:   CharShotMs = 170
        Case CHAR_ELF:      CharShotMs = 130
    End Select
End Function

' incoming combat damage is scaled by this percent (lower = tougher)
Public Function CharArmourPct(ByVal c As Long) As Long
    Select Case c
        Case CHAR_WARRIOR:  CharArmourPct = 90
        Case CHAR_VALKYRIE: CharArmourPct = 60
        Case CHAR_WIZARD:   CharArmourPct = 145
        Case CHAR_ELF:      CharArmourPct = 110
    End Select
End Function

' hits dealt to a generator (or to Death) per melee bump / shot
Public Function CharHitPower(ByVal c As Long) As Long
    Select Case c
        Case CHAR_WARRIOR:  CharHitPower = 2
        Case CHAR_VALKYRIE: CharHitPower = 1
        Case CHAR_WIZARD:   CharHitPower = 1
        Case CHAR_ELF:      CharHitPower = 1
    End Select
End Function

' potion blast reach, in half-cells beyond the viewport edge
Public Function CharPotionMargin(ByVal c As Long) As Long
    Select Case c
        Case CHAR_WIZARD: CharPotionMargin = 10
        Case Else:        CharPotionMargin = 2
    End Select
End Function

' short blurb for the select screen
Public Function CharBlurb(ByVal c As Long) As String
    Select Case c
        Case CHAR_WARRIOR:  CharBlurb = "hard hitter, slow"
        Case CHAR_VALKYRIE: CharBlurb = "best armour, balanced"
        Case CHAR_WIZARD:   CharBlurb = "huge magic, very frail"
        Case CHAR_ELF:      CharBlurb = "fastest, fragile"
    End Select
End Function
