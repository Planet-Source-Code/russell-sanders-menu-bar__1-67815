VERSION 5.00
Begin VB.UserControl MenuItem 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1320
   EditAtDesignTime=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   300
   ScaleWidth      =   1320
   Begin VB.PictureBox tmpImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   690
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   390
      Width           =   225
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   1185
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   90
      Picture         =   "MenuItem.ctx":0000
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   330
      Picture         =   "MenuItem.ctx":014A
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageDrop 
      Height          =   240
      Left            =   390
      Picture         =   "MenuItem.ctx":0314
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageCheck 
      Height          =   240
      Left            =   150
      Picture         =   "MenuItem.ctx":04DE
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'this control could be 3 times faster if the visual properties were
'all drawn in a picture that resides in memory
'CreateCompatibleBitmap/DC or something like that and then load the result to the
'pic object on this control. that would stop so much redrawing of the picturebox.
'that goes the same for the seperator but it's ok.
'Private beenhere As Boolean
Private mForeColor As OLE_COLOR
Private LastBack As Long
Private LastFont As Long
Private curIcon As picture 'the actual icon is in the picture property this picture holds the stated icon (grayed, enabled)
'Property Variables:
Dim m_Picture As picture
Dim m_Caption As String
Dim m_Border As Boolean
Dim m_Pushed As Boolean
Dim m_Shortcut As String
Dim m_Key As String
Dim m_Gradiant As OLE_COLOR
Dim m_Dropdown As Boolean
Dim m_Checked As Boolean
Dim m_IsPopup As Boolean
'Event Declarations:
Event Change()
Event Click(Caption As String, Key As String, pop As Boolean)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Resize()
Private mAxcel As String
Private mButtonStyle As Integer
Private transImage As picture
Implements WinSubHook.iSubclass   'Subclasser interface
Private sc As cSubclass           'Subclasser

Public Property Get BStyle() As Integer
    BStyle = mButtonStyle
End Property

Public Property Let BStyle(NewValue As Integer)
    mButtonStyle = NewValue
    UpdateFavCaption
End Property

Public Property Get Axcel() As String
Attribute Axcel.VB_MemberFlags = "40"
    Axcel = mAxcel
End Property

Public Property Let Axcel(NewValue As String)
    mAxcel = NewValue
End Property

'-------------------------------------------------------------------------------------------------------------------------------
'The next two subs catch the message from the subclass ------------------------------------------------------------------------
Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)
'force a popup if there is one.
    If Not Popups(0) Is Nothing Then
        Popups(UBound(Popups)).hoverMenu = True
        RaiseEvent MouseMove(1, 0, 100, 100)
    End If
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
'
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------

Private Sub UserControl_Initialize()
    'initialize subclassing
'    UserControl.picture = LoadPicture()
'    UserControl.Cls
    Set curIcon = Nothing
    Set sc = New cSubclass
    Call sc.Subclass(Me.hWnd, Me)
    Call sc.AddMsg(WinSubHook.WM_MOUSEHOVER, MSG_AFTER)  ' WM_MOUSEHOVER
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 300
    Pic.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Pic.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Pic.BackColor() = New_BackColor
    UpdateFavCaption
End Property

Private Sub Pic_Change()
    RaiseEvent Change
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    If New_Enabled = False Then
        If Not picture Is Nothing Then
            Pic.ForeColor = &H80000011
            Call PaintGrayScale(tmpImage.hdc, picture, 0, 0, -1, -1)
            tmpImage.picture = tmpImage.Image
            Set curIcon = tmpImage.picture
            tmpImage.Cls
        End If
    Else
        Set curIcon = picture
        Pic.ForeColor = mForeColor
    End If
    UpdateFavCaption
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Pic.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Pic.Font = New_Font
    UpdateFavCaption
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Pic.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    mForeColor = New_ForeColor
    Pic.ForeColor() = New_ForeColor
    UpdateFavCaption
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

Private Sub Usercontrol_Click()
    RaiseEvent Click(Caption, Key, IsPopup)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'Private Sub Pic_KeyDown(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyDown(KeyCode, Shift)
'End Sub
'Private Sub Pic_KeyPress(KeyAscii As Integer)
'    RaiseEvent KeyPress(KeyAscii)
'End Sub
'Private Sub Pic_KeyUp(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyUp(KeyCode, Shift)
'End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    Pushed = True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    'Track Mouse For the Hover Message
    ET.cbSize = Len(ET)
    ET.hwndTrack = Me.hWnd
    ET.dwFlags = TME_HOVER
    ET.dwHoverTime = 100
    TrackMouseEvent ET
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get picture() As picture
Attribute picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set picture = m_Picture
End Property

Public Property Set picture(ByVal New_Picture As picture)
    Set m_Picture = New_Picture
    Set curIcon = New_Picture
    UpdateFavCaption
End Property

Private Sub Pic_Resize()
    RaiseEvent Resize
    UpdateFavCaption
End Sub

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = Me.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    'Pic.ToolTipText() = New_ToolTipText
    Me.ToolTipText() = New_ToolTipText
End Property

Public Property Get Caption() As String
Attribute Caption.VB_MemberFlags = "200"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    If InStr(1, New_Caption, ".") = 0 Then UserControl.Width = 100
    UpdateFavCaption
    RaiseEvent Change
End Property

Public Property Get Border() As Boolean
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    m_Border = New_Border
    UpdateFavCaption
End Property

Private Sub UpdateFavCaption()
'this sub is almost always run twice for each item in the menu and another time if an item needs to be
'wider once placed in the menu
If Caption = "" Or Caption = "&" Then Exit Sub 'dont run unnessary code Each time a visual part of a menu item is changed this function is called
                            'but if the item has no caption set it indicates the item is being loaded and need not be redrawn
                            'until after the caption is passed that saves alot of time
                            'each time one of these properties changes this code is run
                            'checked,dropdown,backcolor,forecolor,font,picture,gradiant,shortcut,caption,border,pushed
                            'This line makes the loading about 10 times faster(still slow though)
Dim TextSize As POINTAPI
Dim lWidth As Long
Dim cString As String
Dim pStr As String
Dim parts() As String
Dim SPoint As Long
Dim fChrPos As Long
    Pic.Cls
    Pic.CurrentY = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight("Notes") / 2) 'set the height and top of the text to be painted
        If BStyle = 0 Then 'Image1.Visible = False Then  'find the starting point for the caption
            If InStr(1, Caption, ".") = 0 Then
                Pic.CurrentX = 30 'if top level item
                lWidth = 30
                fChrPos = 30
            Else
                Pic.CurrentX = 300 'adjust for icon
                lWidth = 300
                fChrPos = 300
            End If
        Else
            Pic.CurrentX = 300 'adjust for icon
            lWidth = 300
            fChrPos = 300
        End If
        
        If Checked = True Then 'adjust for check
            lWidth = lWidth + 300
                'If BStyle = 0 Then
                '    If InStr(1, Caption, ".") = 0 Then
                '        fChrPos = fChrPos + 300
                '    Else
                '        fChrPos = fChrPos + 300
                '    End If
                'Else
                    fChrPos = fChrPos + 300
                'End If
            Pic.CurrentX = fChrPos
        End If
        
        If Dropdown = True Then 'add to width measurement
            lWidth = lWidth + 300 'dropdown Icon
        End If
    cString = Caption
    cString = Replace(cString, "....", "") 'remove char not viewable
    cString = Replace(cString, "&", "")
        If ShortCut <> "" Then
            cString = cString & "    " & ShortCut
        End If
    GetTextExtentPoint32 Pic.hdc, cString, Len(cString), TextSize 'get the length of the caption
    lWidth = lWidth + (TextSize.X * Screen.TwipsPerPixelX)
        If UserControl.Width < lWidth + 90 Then UserControl.Width = lWidth + 90: Exit Sub 'if we must resize then exit after the reize we will be here again
'if we are here we can do our drawing
        If Gradiant <> 0 Then
            DrawGradientFill BackColor, Gradiant, Pic  'call a function to paint a gradiant background
        End If
'Code Bellow looks for The "&" in the caption and draws the following char underlined
    pStr = Replace(Caption, "....", "")
        If InStr(1, pStr, "&") > 0 Then 'if there is a ALT shortcut
            parts = Split(pStr, "&") 'split up the caption
            Pic.Print parts(0) 'print the part before the &
            Pic.CurrentX = fChrPos
            Pic.FontUnderline = True 'set underline font and get the point to start printing the next char
                Pic.CurrentY = (Pic.ScaleHeight / 2) - (Pic.TextHeight("Notes") / 2) 'set the height and top of the text to be painted
                GetTextExtentPoint32 Pic.hdc, parts(0), Len(parts(0)), TextSize 'get the width of what we just printed
                fChrPos = fChrPos + (TextSize.X * Screen.TwipsPerPixelX) 'add that to where we were
                Pic.CurrentX = fChrPos 'set the new position
            Pic.Print Left(parts(1), 1) 'print the text (the underlined char)
            Axcel = Caption & "," & LCase(Left(parts(1), 1)) 'save the key
            Pic.FontUnderline = False 'stop underlining and print the rest of the caption
                Pic.CurrentY = (Pic.ScaleHeight / 2) - (Pic.TextHeight("Notes") / 2) 'set the height and top of the text to be painted
                GetTextExtentPoint32 Pic.hdc, Left(parts(1), 1), 1, TextSize
                fChrPos = fChrPos + (TextSize.X * Screen.TwipsPerPixelX)
                Pic.CurrentX = fChrPos
            Pic.Print Right(parts(1), Len(parts(1)) - 1)
                If ShortCut <> "" Then  'if there is an action key (Ctrl + C for copy)
                    Pic.CurrentY = (Pic.ScaleHeight / 2) - (Pic.TextHeight("Notes") / 2) 'set the height and top of the text to be painted
                    GetTextExtentPoint32 Pic.hdc, ShortCut, 1, TextSize 'get its size
                    If InStr(1, ShortCut, "F") > 0 Then
                        If Dropdown = True Then 'if there is adrop down icon account for its width
                            fChrPos = UserControl.Width - ((TextSize.X * Screen.TwipsPerPixelX) + 450)
                        Else
                            fChrPos = UserControl.Width - ((TextSize.X * Screen.TwipsPerPixelX) + 150)
                        End If
                    Else
                        If Dropdown = True Then 'if there is adrop down icon account for its width
                            fChrPos = UserControl.Width - ((TextSize.X * Screen.TwipsPerPixelX) + 750)
                        Else
                            fChrPos = UserControl.Width - ((TextSize.X * Screen.TwipsPerPixelX) + 450)
                        End If
                    End If
                    Pic.CurrentX = fChrPos 'position the start print position for the shortcut
                    Pic.Print ShortCut
                End If
        Else
            Axcel = vbNullString
            Pic.Print pStr 'print the text onto the background
                If ShortCut <> "" Then  'if there is an action key (Ctrl + C for copy)
                    Pic.CurrentY = (Pic.ScaleHeight / 2) - (Pic.TextHeight("Notes") / 2) 'set the height and top of the text to be painted
                    GetTextExtentPoint32 Pic.hdc, ShortCut, 1, TextSize 'get its size
                        If InStr(1, ShortCut, "F") > 0 Then
                            If Dropdown = True Then 'if there is a drop down icon account for its width
                                fChrPos = UserControl.Width - ((TextSize.X * Screen.TwipsPerPixelX) + 450)
                            Else
                                fChrPos = UserControl.Width - ((TextSize.X * Screen.TwipsPerPixelX) + 150)
                            End If
                        Else
                            If Dropdown = True Then 'if there is a drop down icon account for its width
                                fChrPos = UserControl.Width - ((TextSize.X * Screen.TwipsPerPixelX) + 750)
                            Else
                                fChrPos = UserControl.Width - ((TextSize.X * Screen.TwipsPerPixelX) + 450)
                            End If
                        End If
                    Pic.CurrentX = fChrPos 'position the start of from to the right of the item
                    Pic.Print ShortCut
                End If
        End If
    Erase parts
On Error Resume Next
    If BStyle <> 0 And Not picture Is Nothing Then
        Set tmpImage.picture = curIcon 'Image1.picture
        DrawImageTrans Pic, tmpImage '----this will draw the image with a transparent background but is very slow
        tmpImage.picture = LoadPicture()
        'this bellow is faster but you get a gray background on the icon
        'Pic.PaintPicture Image1.picture, 45, 30, 225, 225, 0, 0, 225, 225
    End If
        'draw a check if checked = true
        If Checked = True Then
            If Enabled = True Then
                If BStyle <> 0 Then
                    Pic.PaintPicture Image2.picture, 300, 45, 225, 225, 0, 0, 225, 225
                Else
                    Pic.PaintPicture Image2.picture, 30, 45, 225, 225, 0, 0, 225, 225
                End If
            Else
                If BStyle <> 0 Then
                    Pic.PaintPicture ImageCheck.picture, 300, 45, 225, 225, 0, 0, 225, 225
                Else
                    Pic.PaintPicture ImageCheck.picture, 30, 45, 225, 225, 0, 0, 225, 225
                End If
            End If
        End If
        'draw a drop icon if item is a popup
        If Dropdown = True Then
            If Enabled = True Then
                Pic.PaintPicture Image3.picture, Pic.Width - 300, 60, 225, 225, 0, 0, 225, 225
            Else
                Pic.PaintPicture ImageDrop.picture, Pic.Width - 300, 60, 225, 225, 0, 0, 225, 225
            End If
        End If
        'draw the border if it has one
        If Border = True Then
            If Pushed = True Then
                Pic.Line (0, 0)-(0, UserControl.Height), &H404040
                Pic.Line (0, 0)-(UserControl.Width, 0), &H404040
                Pic.Line (UserControl.Width - 15, 0)-(UserControl.Width - 15, UserControl.Height), &HFFFFFF
                Pic.Line (0, UserControl.Height - 15)-(UserControl.Width, UserControl.Height - 15), &HFFFFFF
            Else
                Pic.Line (0, 0)-(0, UserControl.Height), &HFFFFFF
                Pic.Line (0, 0)-(UserControl.Width, 0), &HFFFFFF
                Pic.Line (UserControl.Width - 15, 0)-(UserControl.Width - 15, UserControl.Height), &H404040
                Pic.Line (0, UserControl.Height - 15)-(UserControl.Width, UserControl.Height - 15), &H404040
            End If
        End If
    UserControl.picture = Pic.Image
End Sub

Public Property Get Pushed() As Boolean
    Pushed = m_Pushed
End Property

Public Property Let Pushed(ByVal New_Pushed As Boolean)
    m_Pushed = New_Pushed
    UpdateFavCaption
End Property

Public Property Get ShortCut() As String
    ShortCut = m_Shortcut
End Property

Public Property Let ShortCut(ByVal New_Shortcut As String)
    m_Shortcut = New_Shortcut
    UpdateFavCaption
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal New_Key As String)
    m_Key = New_Key
End Property

Public Property Get Gradiant() As OLE_COLOR
    Gradiant = m_Gradiant
End Property

Public Property Let Gradiant(ByVal New_Gradiant As OLE_COLOR)
    m_Gradiant = New_Gradiant
    UpdateFavCaption
End Property

Public Property Get Dropdown() As Boolean
    Dropdown = m_Dropdown
End Property

Public Property Let Dropdown(ByVal New_Dropdown As Boolean)
    m_Dropdown = New_Dropdown
    UpdateFavCaption
End Property

Public Property Get Checked() As Boolean
    Checked = m_Checked
End Property

Public Property Let Checked(ByVal New_Checked As Boolean)
    m_Checked = New_Checked
    UpdateFavCaption
End Property

Public Property Get IsPopup() As Boolean
    IsPopup = m_IsPopup
End Property

Public Property Let IsPopup(ByVal New_IsPopup As Boolean)
    m_IsPopup = New_IsPopup
End Property

Private Sub UserControl_Terminate()
    Set sc = Nothing
End Sub
