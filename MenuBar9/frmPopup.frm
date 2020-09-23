VERSION 5.00
Begin VB.Form frmPopup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   300
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   300
   ScaleWidth      =   390
   ShowInTaskbar   =   0   'False
   Begin MenuBar.Seperator Sep 
      Height          =   90
      Index           =   0
      Left            =   -30
      TabIndex        =   1
      Top             =   330
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   159
   End
   Begin MenuBar.MenuItem MenuItem1 
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   529
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Long 'counter
Private NewHeight As Long 'maintains a constant menu height
Public NewTop As Long 'keeps a running tab on the top for the next menuitem
Dim Spacers As Long 'number of spacers or deviders loaded
Public lastIndex As Integer 'keeps tab on the menuitem we are over so not to run code over and over
Public LastBack As OLE_COLOR 'used to store the original color to be used for restoring after highlight
Public LastFont As OLE_COLOR 'fore color
Public hoverMenu As Boolean 'used to determen if you should display a popup
Public Event Unload(Cancel As Integer)
Public Event Resize()
Public Event Clicked(Caption As String, Tag As String, pop As Boolean)
Public curItem As Integer 'the last item the mouse was over or currently active item
Public Event KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Public Event KeyPress(Index As Integer, KeyAscii As Integer)
Public Event KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Private Sub Form_Load()
    NewHeight = MenuItem1(0).Height 'set the height of all menuitems to the first items height
    NewTop = 0 'set the first menu items top
    Spacers = 1 'one spacer loaded but not shown
    lastIndex = -1 'mouse isn't over a menu
    curItem = -1
End Sub

Public Sub fireEvent(Caption As String, Tag As String, Index As Long)
    RaiseEvent Clicked(Caption, Tag, MenuItem1(Index).IsPopup) 'pass the event
End Sub

'force all items to be the same as the longest item
Private Sub Form_Resize()
    If MenuItem1.Count = 1 Then Me.Width = MenuItem1(0).Width + 90
    For a = 0 To MenuItem1.Count - 2
        MenuItem1(a).Width = MenuItem1(MenuItem1.Count - 1).Width 'make all but the last items width match the last items width
    Next a
    For a = 0 To Sep.Count - 1
        Sep(a).Width = Me.Width 'streach the seperators to fit the form
    Next a
End Sub

'add an item to the popup
Public Function Add(Optional itemType As Integer = 0, Optional Caption As String = "Caption", Optional Ck As Boolean = False, _
Optional Drop As Boolean = False, Optional Popup As Boolean = False, Optional mTag As String = vbNullString, _
Optional bkcol As OLE_COLOR = &H8000000F, Optional ForColor As OLE_COLOR = 0, _
Optional Font As Font = Nothing, Optional Border As Boolean = True, Optional Icon As picture = Nothing, _
Optional Visible As Boolean = True, Optional Enabled As Boolean = True, _
Optional sCut As String = vbNullString, Optional Gradiant As OLE_COLOR = 0, Optional BStyle As Long = 1, _
Optional toolTip As String = vbNullString) As Boolean
'the font option above is only working cause at this point a font is actually being passed
'it's default value as you can see above is set to nothing
    If itemType <> 0 Then
            If Spacers > 1 Then
                Load Sep(Sep.Count)
            End If
        Sep(Sep.Count - 1).Top = NewTop '+ 15
        Sep(Sep.Count - 1).Visible = True
        Sep(Sep.Count - 1).BackColor = bkcol
        Sep(Sep.Count - 1).Gradiant = Gradiant
        'Set Sep(Sep.Count - 1).Font = Font
        Sep(Sep.Count - 1).ForeColor = ForColor
        Sep(Sep.Count - 1).Caption = toolTip
        NewTop = NewTop + Sep(Sep.Count - 1).Height ' + 30
        Spacers = Spacers + 1
    Else
        If curItem <> -1 Then Load MenuItem1(MenuItem1.Count)
            curItem = 0
            With MenuItem1(MenuItem1.Count - 1)
                .Checked = Ck
                .Dropdown = Drop
                .IsPopup = Popup
                .Tag = mTag
                .BStyle = BStyle
                .BackColor = bkcol
                    If MenuItem1.Count - 1 = 0 Then Me.BackColor = bkcol
                .ForeColor = ForColor
                .Border = Border
                .ToolTipText = toolTip
                If Not Font Is Nothing Then Set .Font = Font
                Set .picture = Icon
                .Top = NewTop
                .Height = NewHeight
                .Visible = Visible
                .ShortCut = sCut
                .Enabled = Enabled
                .Gradiant = Gradiant
                .Caption = Caption
                    If Visible = True Then
                        NewTop = NewTop + NewHeight
                        Me.Height = NewTop + 90
                    End If
            'the menu item resizes based on its' contents say if it's wider than the form
            'we need to resize the form.
                If .Width > Me.Width - 90 Then
                    Me.Width = .Width + 90 'if the control is wider than the form resize the form
                Else
                    .Width = Me.Width - 90 'if not resize the control
                End If
            End With
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent Unload(Cancel)
End Sub

Private Sub MenuItem1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(Index, KeyCode, Shift)
End Sub

Private Sub MenuItem1_KeyPress(Index As Integer, KeyAscii As Integer)
    RaiseEvent KeyPress(Index, KeyAscii)
End Sub

Private Sub MenuItem1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(Index, KeyCode, Shift)
End Sub

Public Sub MenuItem1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
    'with the hook installed if we click on a menu item that is a popup the forms would be unloaded
    'so we unload the hook here and the hook will be reset on mouseup if the item here is a popup
        Call UnhookWindowsHookEx(hHook) 'we dont need the hook
        hHook = 0
        RaiseEvent Clicked(MenuItem1(Index).Caption, MenuItem1(Index).Tag, MenuItem1(Index).IsPopup) 'pass the event
    End If
End Sub

Public Sub MenuItem1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
curItem = Index
    If Not lastIndex = Index Then 'Exit Sub 'if we are still over the same menu
       hoverMenu = False
       If Not lastIndex = -1 Then 'set the last highlighted item to it's backcolor/forecolor properties
                MenuItem1(lastIndex).BackColor = LastBack
                MenuItem1(lastIndex).ForeColor = LastFont
            End If
        LastBack = MenuItem1(Index).BackColor 'save the backcolor/forcolor properties we are moving over
        LastFont = MenuItem1(Index).ForeColor
        MenuItem1(Index).BackColor = MenuItem1(Index).ForeColor '&H8000000D 'highlight the item
        MenuItem1(Index).ForeColor = LastBack '&H80000009
        MenuItem1(Index).SetFocus
            For a = UBound(Popups) To 0 Step -1
                'make sure we're in the last popup if not unload the ones after the one we're in
                If Not Me Is Popups(a) Then
                    If Not Popups(a) Is Nothing Then Unload Popups(a)
                    If Not UBound(Popups) = 0 Then ReDim Preserve Popups(UBound(Popups) - 1)
                Else
                    Exit For
                End If
            Next a
        lastIndex = Index
    Else
'I am subclassing the mouse events in each menu item to catch the WM_HOVER message
'where hovermenu will be set to true and this event fired
        If hoverMenu = True Then
            If MenuItem1(Index).IsPopup = True And activeMenu = True Then
            'if current item ispopup then force the popup by raising the clicked event
                RaiseEvent Clicked(MenuItem1(Index).Caption, MenuItem1(Index).Tag, MenuItem1(Index).IsPopup)
            End If
        End If
    End If
End Sub


Public Sub MenuItem1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'we unhook the mouse when we mouse down on a menuitem otherwise the code in the hook proc would unload
    'all the popup windows so if the menu has a popup then we need to rehook it to be ready for mouse
    'events outside the popups.
    If MenuItem1(Index).IsPopup = True Then hHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0&, App.ThreadID)
End Sub

Private Sub MenuItem1_Resize(Index As Integer)
'if the first item in the menu isn't the longest item this isn't needed but I can't get the form to read the width of the item
'right its always short.
    If Me.Width < MenuItem1(0).Width + 60 Then Me.Width = MenuItem1(0).Width + 60
End Sub

