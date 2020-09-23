Attribute VB_Name = "MouseHook"
Option Explicit
'This Hooking code is from MSDN.
Public Type POINTAPI 'used to get the current mouse position
    X As Long
    Y As Long
End Type

Public Type RECT 'used to get the placement of the parent form to the screen
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public cord As RECT
Public thisPoint As POINTAPI
'Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Const WH_MOUSE                           As Long = 7

Public Const HC_ACTION                          As Long = 0 'Im getting 3
Public Const WM_LBUTTONUP                       As Long = &H202 '514
'Public Const WM_LBUTTONDOWN As Long = &H201
Public hHook                                    As Long
Private A                                       As Long

Public Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    MouseProc = CallNextHookEx(hHook, nCode, wParam, lParam) 'Send the event along in case we clicked on a menu item
        If nCode >= 0 And activeMenu = True Then
            If nCode = HC_ACTION And wParam = WM_LBUTTONUP Then
                    For A = 0 To UBound(Popups)
                        If Not Popups(A) Is Nothing Then Unload Popups(A)
                        Set Popups(A) = Nothing
                    Next A
                ReDim Popups(0)
                activeMenu = False
                Call UnhookWindowsHookEx(hHook) 'we dont need the hook any longer
                hHook = 0
            End If
        End If
End Function

