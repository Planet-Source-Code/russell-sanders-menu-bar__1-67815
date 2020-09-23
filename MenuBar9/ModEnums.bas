Attribute VB_Name = "ModEnums"
Option Explicit

'Public tmpImage As PictureBox 'used to draw grayed image on
'Private Type Picts 'store the check and dropdown icons
'    enCheck As picture 'enabled check icon
'    dsCheck As picture 'disabled check icon
'    enDrop As picture 'enabled drop down icon
'    dsDrop As picture 'disabled drop down icon
'End Type
'Public MnuPicts As Picts
'used with the subclass--------------------
Public Const TME_HOVER = &H1&
Public Const WM_MOUSEHOVER = &H2A1&
Public Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long '-1 is default
End Type
Public ET As TRACKMOUSEEVENTTYPE
Public Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
'-------------------------------------------
Public DefaultFont As Font 'not used
Public MB As MainMenu
Public Popups() As frmPopup 'array of popups loaded
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long 'speed things up and reduce flicker
Public activeMenu As Boolean 'if you click a main item the popups will
    'auto load until you make a final selection or select an area outside the menu

'Public Const WM_NCACTIVATE = &H86 'trying to make the parent form retain a FOCUSED LOOK
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long

Public Function TrsLteKey(KeyCode As Integer) As String
'converts a keycode to its' corasponding keyboard key
    Select Case KeyCode
        Case vbKeyReturn: TrsLteKey = "return"
        Case vbKeyLeft: TrsLteKey = "left": Case vbKeyUp: TrsLteKey = "up": Case vbKeyRight: TrsLteKey = "right": Case vbKeyDown: TrsLteKey = "down"
        Case vbKeyA: TrsLteKey = "a": Case vbKeyB: TrsLteKey = "b": Case vbKeyC: TrsLteKey = "c": Case vbKeyD: TrsLteKey = "d"
        Case vbKeyE: TrsLteKey = "e": Case vbKeyF: TrsLteKey = "f": Case vbKeyG: TrsLteKey = "g": Case vbKeyH: TrsLteKey = "h"
        Case vbKeyI: TrsLteKey = "i": Case vbKeyJ: TrsLteKey = "j": Case vbKeyK: TrsLteKey = "k": Case vbKeyL: TrsLteKey = "l"
        Case vbKeyM: TrsLteKey = "m": Case vbKeyN: TrsLteKey = "n": Case vbKeyO: TrsLteKey = "o": Case vbKeyP: TrsLteKey = "p"
        Case vbKeyQ: TrsLteKey = "q": Case vbKeyR: TrsLteKey = "r": Case vbKeyS: TrsLteKey = "s": Case vbKeyT: TrsLteKey = "t"
        Case vbKeyU: TrsLteKey = "u": Case vbKeyV: TrsLteKey = "v": Case vbKeyW: TrsLteKey = "w": Case vbKeyX: TrsLteKey = "x"
        Case vbKeyY: TrsLteKey = "y": Case vbKeyZ: TrsLteKey = "z": Case vbKey0: TrsLteKey = "0": Case vbKey1: TrsLteKey = "1"
        Case vbKey2: TrsLteKey = "2": Case vbKey3: TrsLteKey = "3": Case vbKey4: TrsLteKey = "4": Case vbKey5: TrsLteKey = "5"
        Case vbKey6: TrsLteKey = "6": Case vbKey7: TrsLteKey = "7": Case vbKey8: TrsLteKey = "8": Case vbKey9: TrsLteKey = "9"
        Case vbKeyF1: TrsLteKey = "f1": Case vbKeyF2: TrsLteKey = "f2": Case vbKeyF3: TrsLteKey = "f3": Case vbKeyF4: TrsLteKey = "f4"
        Case vbKeyF5: TrsLteKey = "f5": Case vbKeyF6: TrsLteKey = "f6": Case vbKeyF7: TrsLteKey = "f7": Case vbKeyF8: TrsLteKey = "f8"
        Case vbKeyF9: TrsLteKey = "f9": Case vbKeyF10: TrsLteKey = "f10": Case vbKeyF11: TrsLteKey = "f11": Case vbKeyF12: TrsLteKey = "f12"
        Case vbKeyF13: TrsLteKey = "f13": Case vbKeyF14: TrsLteKey = "f14": Case vbKeyF15: TrsLteKey = "f15": Case vbKeyF16: TrsLteKey = "f16"
        'Case vbKeyLButton: TrsLteKey = "left": Case vbKeyRButton: TrsLteKey = "right": Case vbKeyCancel: TrsLteKey = "Cancle": Case vbKeyMButton: TrsLteKey = "MButton"
        'Case vbKeyBack: TrsLteKey = "back": Case vbKeyTab: TrsLteKey = "tab": Case vbKeyClear: TrsLteKey = "clear"
        'Case vbKeyShift: TrsLteKey = "shift":Case vbKeyControl: TrsLteKey = "control"
        'Case vbKeyMenu: TrsLteKey = "menu":Case vbKeyPause: TrsLteKey = "pause":Case vbKeyCapital: TrsLteKey = "capital"
        'Case vbKeyEscape: TrsLteKey = "escape":Case vbKeySpace: TrsLteKey = "space":Case vbKeyPageUp: TrsLteKey = "pageup"
        'Case vbKeyPageDown: TrsLteKey = "pagedown":Case vbKeyEnd: TrsLteKey = "end":Case vbKeyHome: TrsLteKey = "home"
        'Case vbKeySelect: TrsLteKey = "select":Case vbKeyPrint: TrsLteKey = "print":Case vbKeyExecute: TrsLteKey = "execute"
        'Case vbKeySnapshot: TrsLteKey = "screenshot":Case vbKeyInsert: TrsLteKey = "insert":Case vbKeyDelete: TrsLteKey = "delete"
        'Case vbKeyHelp: TrsLteKey = "help":Case vbKeyNumlock: TrsLteKey = "numlock":Case vbKeyScrollLock: TrsLteKey = "scrolllock"
    End Select
End Function
