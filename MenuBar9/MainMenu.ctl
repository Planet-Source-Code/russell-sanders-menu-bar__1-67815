VERSION 5.00
Begin VB.UserControl MainMenu 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1110
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   KeyPreview      =   -1  'True
   PropertyPages   =   "MainMenu.ctx":0000
   ScaleHeight     =   390
   ScaleWidth      =   1110
   ToolboxBitmap   =   "MainMenu.ctx":0013
   Begin MenuBar.MenuItem MenuItem1 
      Height          =   300
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   529
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   -500
      Top             =   -500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   15
      X2              =   15
      Y1              =   30
      Y2              =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   1080
      X2              =   1080
      Y1              =   30
      Y2              =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   15
      X2              =   1020
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   15
      X2              =   1050
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   3
      X1              =   0
      X2              =   1080
      Y1              =   345
      Y2              =   360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   1
      X1              =   15
      X2              =   1050
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   2
      X1              =   1080
      X2              =   1080
      Y1              =   30
      Y2              =   315
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   0
      X1              =   30
      X2              =   30
      Y1              =   60
      Y2              =   345
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'quirks

Option Explicit
'Default Property Values:
Private WithEvents parn As Form 'removes the need to pass key events from the parent form
Attribute parn.VB_VarHelpID = -1
Private WithEvents MDIparn As MDIForm
Attribute MDIparn.VB_VarHelpID = -1
Private a As Long 'loop counter
Private Const m_def_ItemCount = 1
Private Const m_def_Border = True
Private Const m_def_TotalItems As Long = 0 'the array size from the property bag
Private Const m_def_BorderStyle As Integer = 0
Private Const m_def_BotBorder As Boolean = True '
Private Const m_def_Children As String = vbNullString
Private Const m_def_TopBorder As Boolean = True
Private Const m_def_LeftBorder As Boolean = True
Private Const m_def_RightBorder As Boolean = True
Private Const m_def_retainApp As Boolean = False
'Property Variables:
Private allMenus() As mnus 'an array of all menu items and there properties
Private tmpitem As mnus 'used when moving an item up or down in the array
Private mretainApp As Boolean 'tells the control to color all items the same or not
Private newleft As Long 'used for placement of the menu items on the menu bar
Private ItemCount As Integer 'used in readproperties to track the number of items loaded, or rather make sure we arn't at the first item
Private m_Border As Boolean
Private selIndex As Integer 'used to track the item the mouse is over in the main menu bar
Private FirstPart As String 'used to store any "." found in front of a menuitem name
Private SearchFor As String 'used to store any "." found in front of a menuitem name
Private mTopBorder As Boolean
Private mLeftBorder As Boolean
Private mGradientDirection As Long
Private mRightBorder As Boolean
Private mBotBorder As Boolean
Private mChildren As String  'not used now read through the code to understand why its here
Private mTotalItems As Long 'used to keep track of the total number of menuitems in the menu array allowing us to reload
Private mBorderStyle As Integer
Private b As Long 'loop counter
Public curItemM As Integer 'the active menu item
Attribute curItemM.VB_VarMemberFlags = "40"
'Event Declarations:
Private WithEvents mnuPopup As frmPopup 'catch the click from the popup form
Attribute mnuPopup.VB_VarHelpID = -1
Event Resize()
Public Event Click(Caption As String, Key As String)
Attribute Click.VB_MemberFlags = "200"

Public Enum gStyle 'picture in menu = aGraphical
    aStandard
    aGraphical
End Enum

Public Enum checkStyle 'used in the style property of the menuitem
    stlNone = -1
    stlChecked         'property isn't atually in the control but only saved in the property bag
    stlOption          'with the others
End Enum

Public Type mnus 'structure to hold menu info
    Path As String 'isn't used
    Children As String 'isn't used
    Caption As String 'holds the caption of the control
    Tag As String     'will hold any string you like
    ispop As Boolean  'tells the code to load a popup menu
    Drop As Boolean   'visual aid to a dropdown menu
    Check As Boolean  'check option on/off
    Icon As picture   'picture for menu item
    Tool As String    'tool tip text
    bkcol As OLE_COLOR 'backcolor
    frCol As OLE_COLOR 'fore color
    Gradiant As OLE_COLOR 'Gradiant 0 = noGradiant
    Border As Boolean  'border on/off
    Enabled As Boolean  'enable/disable
    Visible As Boolean  'show/hide
    parent As String    'used when loading new items from outside the control
    Font As Font       'font
    sCut As String     'short cut key combo
    Group As Integer   'allows you to have groups of items
    Style As checkStyle 'checked or option style the option style only allows one item in the group to be checked
    BStyle As gStyle 'Picture or no picture
End Type

Public Enum BorderStyles 'set the border of the Main Menu bar
    Framed
    Raised
    Sunken
End Enum

Public Enum GradDirect 'set the direction of the gradient
    grdHorz  'gradient from top to bottom
    grdHorCOut 'gradient from top to center and from bottom to center
    grdVert 'gradient from left to right
    graVerCOut ' from left to center and from right to center
End Enum

Implements WinSubHook.iSubclass   'Subclasser interface
Private sc2 As cSubclass           'Subclasser


'added raised and sunken to the framed style 0 = framed, 1 = raised, 2 = sunken
Public Property Get BorderStyle() As BorderStyles
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(NewValue As BorderStyles)
    mBorderStyle = NewValue
    PropertyChanged "BorderStyle"
    Border = Border
End Property

Public Property Get GradientDirection() As GradDirect
    GradientDirection = mGradientDirection
End Property

Public Property Let GradientDirection(NewValue As GradDirect)
    mGradientDirection = NewValue
    GRADIENT_FILL_RECT_DIRECTION = NewValue
    PropertyChanged GradientDirection
    ForcePaint
End Property

Private Sub FireAccel(keys As Integer) 'used with the menu keys alt + some letter
On Error Resume Next
Dim a As Long
Dim prt() As String
Dim MatchString As String
Dim Keys2 As String
Dim found As Boolean
Dim b As Long
Keys2 = TrsLteKey(keys)
found = False
    If Popups(0) Is Nothing Then 'no popups check the main items for key
        For a = 0 To MenuItem1.Count - 1
            If MenuItem1(a).Axcel = MenuItem1(a).Caption & "," & Keys2 Then
                MenuItem1(a).SetFocus
                MenuItem1_Click CInt(a), MenuItem1(a).Caption, MenuItem1(a).Tag, MenuItem1(a).IsPopup
                found = True
                Exit For
            End If
        Next a
    Else ' if there are some popup forms.
        For b = UBound(Popups) To 0 Step -1 'check in the last popup form to see if any items have the key and go back from there
            With Popups(b)
                For a = 0 To .MenuItem1.Count - 1
                    If .MenuItem1(a).Axcel = .MenuItem1(a).Caption & "," & Keys2 Then 'once found
                            If b < UBound(Popups) Then 'unload any popups after the one the key is found in
                                Do
                                    Unload Popups(UBound(Popups))
                                    ReDim Preserve Popups(UBound(Popups) - 1)
                                Loop Until UBound(Popups) <= b
                            End If
                        .MenuItem1_MouseMove CInt(a), 1, 0, 100, 100
                        .fireEvent mnuPopup.MenuItem1(CInt(a)).Caption, mnuPopup.MenuItem1(CInt(a)).Tag, a
                        '.MenuItem1_MouseDown CInt(A), 1, 0, 100, 100 'send the mouse down event for that item
                        '.MenuItem1_MouseUp CInt(A), 1, 0, 100, 100
                         found = True 'pervent further action
                        Exit For 'exit loop
                    End If
                Next a
            End With
            If found Then Exit For 'if we found the item we exit this loop to
        Next b 'if we didn't find the item check the popup before this one
        If found = False Then 'is all popups were checked and we didn't find the item
            For a = 0 To MenuItem1.Count - 1 'check the mainmenu items
                If MenuItem1(a).Axcel = MenuItem1(a).Caption & "," & Keys2 Then
                    found = True
                        If Not Popups(0) Is Nothing Then 'here we can unload all popups
                            For b = 0 To UBound(Popups)
                                Unload Popups(b)
                                Set Popups(b) = Nothing
                            Next b
                            ReDim Popups(0)
                        End If
                    MenuItem1_Click CInt(a), MenuItem1(a).Caption, MenuItem1(a).Tag, MenuItem1(a).IsPopup 'fire click event
                    Exit For
                End If
            Next a
        End If
    End If
End Sub

Public Property Get retainApp() As Boolean
    retainApp = mretainApp
End Property

Public Property Let retainApp(NewValue As Boolean)
    mretainApp = NewValue
    PropertyChanged "retainApp"
End Property

'this was added to allow property page to update the number of menu items
Public Sub DimToCount(Count As Long, Optional aPreserve As Boolean = True)
Attribute DimToCount.VB_MemberFlags = "40"
    If aPreserve = True Then
        ReDim Preserve allMenus(Count)
    Else
        ReDim allMenus(Count)
    End If
End Sub

'This allows the property page to update an items properties
Public Sub SetItemInfo(item As Integer, ByRef ItemInfo As mnus)
Attribute SetItemInfo.VB_MemberFlags = "40"
        If item > UBound(allMenus) Then Exit Sub 'ensure we arn't asking for an item that doesn't exist
    With allMenus(item) 'save structure to array
        .bkcol = ItemInfo.bkcol
        .Border = ItemInfo.Border
        .Children = ItemInfo.Children
        .Caption = ItemInfo.Caption
        .Check = ItemInfo.Check
        .Drop = ItemInfo.Drop
        .frCol = ItemInfo.frCol
        Set .Icon = ItemInfo.Icon
        Set .Font = ItemInfo.Font
        .ispop = ItemInfo.ispop
        .Path = ItemInfo.Path
        .Tag = ItemInfo.Tag
        .Tool = ItemInfo.Tool
        .Visible = ItemInfo.Visible
        .sCut = ItemInfo.sCut
        .Group = ItemInfo.Group
        .Style = ItemInfo.Style
        .Enabled = ItemInfo.Enabled
        .parent = ItemInfo.parent
        .Gradiant = ItemInfo.Gradiant
        .BStyle = ItemInfo.BStyle
    End With
End Sub
Public Function GetItemInfo(item As Integer) As mnus 'this function allows the property page to get an items properties
Attribute GetItemInfo.VB_MemberFlags = "40"
        If item > UBound(allMenus) Then Exit Function 'ensure we arn't asking for an item that doesn't exist
    With allMenus(item) 'load structure from array
        GetItemInfo.bkcol = .bkcol
        GetItemInfo.Border = .Border
        GetItemInfo.Children = .Children
        GetItemInfo.Caption = .Caption
        GetItemInfo.Check = .Check
        GetItemInfo.Drop = .Drop
        GetItemInfo.frCol = .frCol
        Set GetItemInfo.Icon = .Icon
        Set GetItemInfo.Font = .Font
        GetItemInfo.ispop = .ispop
        GetItemInfo.Path = .Path
        GetItemInfo.Tag = .Tag
        GetItemInfo.Tool = .Tool
        GetItemInfo.Visible = .Visible
        GetItemInfo.sCut = .sCut
        GetItemInfo.Group = .Group
        GetItemInfo.Style = .Style
        GetItemInfo.Enabled = .Enabled
        GetItemInfo.parent = .parent
        GetItemInfo.Gradiant = .Gradiant
        GetItemInfo.BStyle = .BStyle
    End With
End Function

'adds a default item to the array with optional settings
Public Sub AddItem(Path As String, Caption As String, _
    Optional Tag As String = vbNullString, Optional ispop As Boolean = False, _
    Optional Drop As Boolean = False, Optional Check As Boolean = False, _
    Optional Icon As picture = Nothing, Optional Tool As String = vbNullString, _
    Optional bkcol As OLE_COLOR = &H8000000F, Optional frCol As OLE_COLOR = 0, _
    Optional Border As Boolean = False, Optional Enabled As Boolean = True, _
    Optional Children As String = vbNullString, Optional Visible As Boolean = True, _
    Optional sCut As String = vbNullString, Optional Font As Font = Nothing, Optional Group As Integer = 0, _
    Optional Style As Integer = -1, Optional Gradiant As OLE_COLOR = 0, Optional BStyle As gStyle = aGraphical)
Attribute AddItem.VB_MemberFlags = "40"
            If allMenus(0).Caption <> "" Then ReDim Preserve allMenus(UBound(allMenus) + 1)
        With allMenus(UBound(allMenus))
            .bkcol = bkcol
            .Border = Border
            .Children = Children
            .Caption = Caption
            .Check = Check
            .Drop = Drop
            .frCol = frCol
            Set .Icon = Icon
                If Not Font Is Nothing Then
                    Set .Font = Font
                Else
                    Set .Font = UserControl.Font
                End If
            .ispop = ispop
            .Path = Path
            .Tag = Tag
            .Tool = Tool
            .Visible = Visible
            .sCut = sCut
            .Group = Group
            .Style = Style
            .Enabled = Enabled
            .parent = ""
            .Gradiant = Gradiant
            .BStyle = BStyle
        End With
End Sub

Public Sub removeItem(ItemNum As Integer) 'removes an item during the creation of the menu through property page
Attribute removeItem.VB_MemberFlags = "40"
'to remove an item at runtime use the "Remove" feature of the main menu
If UBound(allMenus) = 0 Then ReDim allMenus(0): Exit Sub
        For a = ItemNum To UBound(allMenus)
            If a > ItemNum Then 'starting at the item position bellow the one to remove
                allMenus(a - 1) = allMenus(a) 'move each item up one
            End If
        Next a
        If UBound(allMenus) = 0 Then
            ReDim allMenus(0)
        Else
            ReDim Preserve allMenus(UBound(allMenus) - 1) 'srink the array by one
        End If
End Sub

Public Sub UpItem(ItemNum As Integer) 'move item up in the array
Attribute UpItem.VB_MemberFlags = "40"
        If ItemNum = 0 Then Exit Sub 'we checked this before the call but just in case
    tmpitem = allMenus(ItemNum - 1) 'save the properties of the item in the position we are taking
    allMenus(ItemNum - 1) = allMenus(ItemNum) 'move the item into the position
    allMenus(ItemNum) = tmpitem 'load the properties of the item we moved into the item we moved from
End Sub

Public Sub DnItem(ItemNum As Integer) 'move item down in the array
Attribute DnItem.VB_MemberFlags = "40"
        If ItemNum = UBound(allMenus) Then Exit Sub
    tmpitem = allMenus(ItemNum + 1)
    allMenus(ItemNum + 1) = allMenus(ItemNum)
    allMenus(ItemNum) = tmpitem
End Sub

Public Sub InsertItem(ItemNum As Integer) 'after adding item use setiteminfo to load properties
Attribute InsertItem.VB_MemberFlags = "40"
    ReDim Preserve allMenus(UBound(allMenus) + 1) 'add blank item to array
        For a = UBound(allMenus) - 1 To 0 Step -1
            If a >= ItemNum Then
                allMenus(a + 1) = allMenus(a)
                    With allMenus(a)
                        .bkcol = &H8000000F
                        .Border = False
                        .Caption = ""
                        .Check = False
                        .Children = ""
                        .Drop = False
                        .Enabled = True
                        .frCol = 0
                        Set .Icon = Nothing
                        Set .Font = UserControl.Font
                        .ispop = False
                        .Path = ""
                        .Tag = ""
                        .Tool = ""
                        .Visible = True
                        .Group = 0
                        .sCut = ""
                        .Style = stlNone
                        .parent = ""
                        .Gradiant = 0
                        .BStyle = 1
                    End With
            Else
                Exit For
            End If
        Next a
End Sub

Public Property Get Style(item As Long) As Integer 'check style of an item
    Style = allMenus(item).Style
End Property

Public Property Let Style(item As Long, NewValue As Integer)
    allMenus(item).Style = NewValue
End Property

'The five functions bellow are what I will call property relays. There intent is to Request and set the properties
'in the menu array These properties are only valid for the current session and will be reset when the program is restarted
'you can get/check the state of the visible, enabled, and checked properties. You can see the useage in the test form
'also need to setup a back/fore color and font passthrough. you could then set all properties at run time

'Set Caption
Public Sub Caption(item As String, NewValue As String)
    For a = 0 To UBound(allMenus)
        If Replace(allMenus(a).Caption, ".", "") = item Then
            allMenus(a).Caption = Replace(allMenus(a).Caption, item, "") & NewValue
                If InStr(1, allMenus(a).Caption, ".") = 0 Then
                    For b = 0 To MenuItem1.Count - 1
                        If MenuItem1(b).Caption = item Then
                            MenuItem1(b).Caption = NewValue
                            ResetItems
                            Exit For
                        End If
                    Next b
                End If
            Exit For
        End If
    Next a
End Sub
'Get/Set group
Public Function Group(item As String, Optional NewValue As Integer = 0, Optional retVal As Boolean = True) As Integer
    For a = 0 To UBound(allMenus)
        If Replace(allMenus(a).Caption, ".", "") = item Then
                If retVal = True Then
                    Group = allMenus(a).Group
                Else
                    allMenus(a).Group = NewValue
                    Group = NewValue
                End If
            Exit For
        End If
    Next a
End Function
'Get/Set Check
Public Function Check(item As String, Optional NewValue As Boolean = False, Optional retVal As Boolean = True) As Boolean
    For a = 0 To UBound(allMenus)
        If Replace(allMenus(a).Caption, ".", "") = item Then
                If retVal = True Then
                    Check = allMenus(a).Check
                Else
                    allMenus(a).Check = NewValue
                        If InStr(1, allMenus(a).Caption, ".") = 0 Then
                                For b = 0 To MenuItem1.Count - 1
                                    If MenuItem1(b).Caption = item Then
                                        MenuItem1(b).Checked = NewValue
                                        Exit For
                                    End If
                                Next b
                            ResetItems
                        End If
                    Check = NewValue
                End If
            Exit For
        End If
    Next a
End Function
'Get/Set Visible
Public Function isVisible(item As String, Optional NewValue As Boolean = False, Optional retVal As Boolean = True) As Boolean
    For a = 0 To UBound(allMenus)
        If Replace(allMenus(a).Caption, ".", "") = item Then
                If retVal = True Then
                    isVisible = allMenus(a).Visible
                Else
                    allMenus(a).Visible = NewValue
                        If InStr(1, allMenus(a).Caption, ".") = 0 Then
                            Dim b As Long
                                For b = 0 To MenuItem1.Count - 1
                                    If MenuItem1(b).Caption = item Then
                                        MenuItem1(b).Visible = NewValue
                                        Exit For
                                    End If
                                Next b
                            ResetItems
                        End If
                    isVisible = NewValue
                End If
            Exit For
        End If
    Next a
End Function
'Get/Set Enabled
Public Function isEnabled(item As String, Optional NewValue As Boolean = False, Optional retVal As Boolean = True) As Boolean
    For a = 0 To UBound(allMenus)
        If Replace(allMenus(a).Caption, ".", "") = item Then
                If retVal = True Then
                    isEnabled = allMenus(a).Enabled
                Else
                    If InStr(1, allMenus(a).Caption, ".") > 0 Then
                        allMenus(a).Enabled = NewValue
                        isEnabled = NewValue
                    Else
                        For b = 0 To MenuItem1.Count - 1
                            If MenuItem1(b).Caption = item Then
                                MenuItem1(b).Enabled = NewValue
                                allMenus(a).Enabled = NewValue
                                isEnabled = NewValue
                            End If
                        Next b
                    End If
                End If
            Exit For
        End If
    Next a
End Function

Public Property Get TotalItems() As Long 'a count of all menu items
Attribute TotalItems.VB_MemberFlags = "40"
    TotalItems = mTotalItems
End Property

Public Property Let TotalItems(NewValue As Long)
    mTotalItems = NewValue
End Property

Public Property Get Children() As String 'a list of main menu items
Attribute Children.VB_MemberFlags = "40"
    Children = mChildren
End Property

Public Property Let Children(NewValue As String)
    mChildren = NewValue
    PropertyChanged "Children"
End Property

Public Sub ClearGroup(item As Integer) 'used with the style property to uncheck the items in a group
Attribute ClearGroup.VB_MemberFlags = "40"
'and then check the item selected
Dim a As Long
    For a = 0 To UBound(allMenus)
        If allMenus(a).Group = allMenus(item).Group Then
            If a = item Then
                allMenus(a).Check = True
            Else
                allMenus(a).Check = False
            End If
        End If
    Next a
End Sub

Private Sub LoadPopupItems(popForm As frmPopup, Caption As String)
'this code was moved into asub to reduce code
'Function: Searches the menu array for a match to Caption and
'loads any children it finds under that item into popform.
Dim ItemIndex As Integer
Dim htdif As Long
Dim cnt As Long
GradientDirection = mGradientDirection
    With popForm
        Load popForm
        For a = 0 To UBound(allMenus)
            If allMenus(a).Caption = Caption Then
                ItemIndex = CInt(a)
                    If InStr(1, Caption, ".") > 0 Then
                        FirstPart = Left(Caption, InStrRev(Caption, "."))
                    Else
                        FirstPart = ""
                    End If
                    For b = a + 1 To UBound(allMenus)
                        SearchFor = Left(allMenus(b).Caption, InStrRev(allMenus(b).Caption, "."))
                            If SearchFor = FirstPart & "...." Then
                                If Replace(allMenus(b).Caption, ".", "") = "-" Then
                                    If retainApp = True Then
                                        .Add 1, allMenus(b).Caption, allMenus(b).Check, allMenus(b).Drop, allMenus(b).ispop, allMenus(b).Tag, allMenus(0).bkcol, allMenus(0).frCol, allMenus(b).Font, allMenus(b).Border, allMenus(b).Icon, allMenus(b).Visible, allMenus(b).Enabled, allMenus(b).sCut, allMenus(0).Gradiant, allMenus(b).BStyle, allMenus(b).Tool
                                    Else
                                        .Add 1, allMenus(b).Caption, allMenus(b).Check, allMenus(b).Drop, allMenus(b).ispop, allMenus(b).Tag, allMenus(b).bkcol, allMenus(b).frCol, allMenus(b).Font, allMenus(b).Border, allMenus(b).Icon, allMenus(b).Visible, allMenus(b).Enabled, allMenus(b).sCut, allMenus(b).Gradiant, allMenus(b).BStyle, allMenus(b).Tool
                                    End If
                                Else
                                    If retainApp = True Then
                                        .Add , allMenus(b).Caption, allMenus(b).Check, allMenus(b).Drop, allMenus(b).ispop, allMenus(b).Tag, allMenus(0).bkcol, allMenus(0).frCol, allMenus(b).Font, allMenus(b).Border, allMenus(b).Icon, allMenus(b).Visible, allMenus(b).Enabled, allMenus(b).sCut, allMenus(0).Gradiant, allMenus(b).BStyle, allMenus(b).Tool
                                    Else
                                        .Add , allMenus(b).Caption, allMenus(b).Check, allMenus(b).Drop, allMenus(b).ispop, allMenus(b).Tag, allMenus(b).bkcol, allMenus(b).frCol, allMenus(b).Font, allMenus(b).Border, allMenus(b).Icon, allMenus(b).Visible, allMenus(b).Enabled, allMenus(b).sCut, allMenus(b).Gradiant, allMenus(b).BStyle, allMenus(b).Tool
                                    End If
                                End If
                            ElseIf SearchFor = FirstPart Then
                                Exit For
                            End If
                    Next b
                Exit For
            End If
        Next a
    End With
End Sub


Private Sub mnuPopup_Clicked(Caption As String, Tag As String, pop As Boolean)
'event from the popup forms
Dim ItemIndex As Integer
Dim htdif As Long
Dim cnt As Long
        If pop = True Then 'if ispopup then load new form
        RaiseEvent Click(Replace(Caption, ".", ""), Tag) ': DoEvents
        activeMenu = True
            If Not Popups(0) Is Nothing Then ReDim Preserve Popups(UBound(Popups) + 1)
            Set Popups(UBound(Popups)) = New frmPopup
            Set mnuPopup = Popups(UBound(Popups))
            'If hHook = 0 Then hHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0&, App.ThreadID)
            With mnuPopup 'later these will be read from the property bag
                LoadPopupItems mnuPopup, Caption
                    'This only works if the item is docked to the top of the form
                    'ensure the menu popup doesn't go beyond the right or bottom of the screen
                    If UBound(Popups) = 0 Then
                        .Left = Popups(0).Left + Popups(0).Width
                        .Top = Popups(0).ActiveControl.Top + Popups(0).Top
                    Else
                        If Popups(UBound(Popups) - 1).Left + Popups(UBound(Popups) - 1).Width + Popups(UBound(Popups)).Width > Screen.Width Then
                            .Left = Popups(UBound(Popups) - 1).Left - Popups(UBound(Popups)).Width
                        Else
                            .Left = Popups(UBound(Popups) - 1).Left + Popups(UBound(Popups) - 1).Width
                        End If

                        If Popups(UBound(Popups) - 1).ActiveControl.Top + Popups(UBound(Popups) - 1).Top + Popups(UBound(Popups)).Height > (Screen.Height - 900) Then
                            htdif = Popups(UBound(Popups) - 1).ActiveControl.Top + Popups(UBound(Popups) - 1).Top + Popups(UBound(Popups)).Height - (Screen.Height - 800)
                            .Top = Popups(UBound(Popups) - 1).ActiveControl.Top + Popups(UBound(Popups) - 1).Top - htdif
                        Else
                            .Top = Popups(UBound(Popups) - 1).ActiveControl.Top + Popups(UBound(Popups) - 1).Top
                        End If
                    End If
                .Show , UserControl
                HighlightFirstEnabled mnuPopup
            End With
        Else
        'selection made unload all popups
        unloadAll
        cleanmenu
                For a = 0 To UBound(allMenus) 'this only finds the item number for the next step
                    If allMenus(a).Caption = Caption Then
                    ItemIndex = CInt(a)
                    Exit For
                    End If
                Next a
        RaiseEvent Click(Replace(Caption, ".", ""), Tag) ': DoEvents
        End If
            If allMenus(ItemIndex).Style <> stlNone Then
                If allMenus(ItemIndex).Style = 0 Then
                    allMenus(ItemIndex).Check = Not allMenus(ItemIndex).Check
                Else
                    ClearGroup CInt(ItemIndex)
                End If
            End If
End Sub

Private Sub unloadAll()
        For a = 0 To UBound(Popups)   ' To 0 Step -1
            If Not Popups(a) Is Nothing Then Unload Popups(a): Set Popups(a) = Nothing
        Next a
    activeMenu = False
    ReDim Popups(0)
End Sub

Private Sub mnuPopup_Unload(Cancel As Integer)
    If Not UBound(Popups) = 0 Then
        Set mnuPopup = Popups(UBound(Popups) - 1)
    Else
        MenuItem1(curItemM).SetFocus
        ReDim Popups(0)
    End If
End Sub

Private Sub MenuItem1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Ambient.UserMode() Then Exit Sub
    If Button = 1 Then
        MenuItem1(Index).Pushed = MenuItem1(Index).IsPopup
    End If
End Sub

Private Sub MenuItem1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Ambient.UserMode() Then Exit Sub
    If Button = 1 Then
        MenuItem1(Index).Pushed = True
    End If
End Sub

Private Sub MenuItem1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Not Ambient.UserMode() Then Exit Sub
    GRADIENT_FILL_RECT_DIRECTION = mGradientDirection
        If selIndex = Index Then Exit Sub
    cleanmenu
    MenuItem1(Index).Border = True
    MenuItem1(Index).SetFocus
    selIndex = Index
    curItemM = Index
        If activeMenu = False Then
            Exit Sub
        End If
        If MenuItem1(Index).IsPopup = True Then
            MenuItem1(Index).Pushed = True
            MenuItem1_Click Index, MenuItem1(Index).Caption, MenuItem1(Index).Tag, MenuItem1(Index).IsPopup
        Else
            If Not Popups(0) Is Nothing Then
                For a = 0 To UBound(Popups)
                    Unload Popups(a)
                    Set Popups(a) = Nothing
                Next a
            End If
        End If
End Sub

Private Sub MenuItem1_Click(Index As Integer, Caption As String, Key As String, pop As Boolean)
'same as mnuPopup_Clicked above
    If Not Ambient.UserMode() Then Exit Sub
On Error Resume Next
Dim ItemIndex As Integer
Dim Corners As RECT
Dim p As POINTAPI
    RaiseEvent Click(Caption, Key)
    selIndex = Index
    unloadAll
    ReDim Popups(0)
        If MenuItem1(Index).IsPopup = True Then
        If hHook = 0 Then hHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0&, App.ThreadID)
            activeMenu = True
            Set Popups(0) = New frmPopup
            Set mnuPopup = Popups(0)
                With Popups(0)
                    LoadPopupItems Popups(0), Caption
                    GetCursorPos p
                    Call GetWindowRect(MenuItem1(Index).hWnd, cord)
                    .Left = (p.X - (p.X - cord.Left)) * Screen.TwipsPerPixelX
                    .Top = (p.Y - (p.Y - cord.Bottom)) * Screen.TwipsPerPixelY
                    .Show vbModeless, UserControl
                    HighlightFirstEnabled Popups(0)
                End With
        Else
            unloadAll
            cleanmenu
                For a = 0 To UBound(allMenus) 'this only finds the item number for the next step
                    If allMenus(a).Caption = Caption Then
                    ItemIndex = CInt(a)
                    Exit For
                    End If
                Next a
        End If
            If allMenus(ItemIndex).Style <> stlNone Then 'This controls the checked/unchecked properties of the item If set to none nothing is done
                If allMenus(ItemIndex).Style = 0 Then 'if it's set to checked then the item will be auto checked when clicked 'if set to option(Must be used with the group property) all other menu items in the group will
                    allMenus(ItemIndex).Check = Not allMenus(ItemIndex).Check 'be unchecked and the clicled one will be checked. You can still set the checked property yourself
                Else
                    ClearGroup CInt(ItemIndex)
                End If
            End If
End Sub

Public Sub cleanmenu()
Attribute cleanmenu.VB_MemberFlags = "40"
    For a = 0 To MenuItem1.Count - 1
        MenuItem1(a).Border = False
        MenuItem1(a).Pushed = False
    Next a
End Sub

'Private Sub parn_Deactivate()
'    Debug.Print "deactivate"
'End Sub
'
'Private Sub parn_LostFocus()
'    Debug.Print "Lost Focus"
'End Sub

Private Sub UserControl_Initialize()
    ReDim allMenus(0)
    ReDim Popups(0)
    selIndex = -1
    ItemCount = 1
    newleft = 60
End Sub

Private Sub mnuPopup_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Shift
        Case 4
            FireAccel KeyCode
        Case 0
        Select Case KeyCode
            Case 13
                mnuPopup.fireEvent mnuPopup.MenuItem1(Index).Caption, mnuPopup.MenuItem1(Index).Tag, CLng(Index)
            Case 37
                If Not UBound(Popups) = 0 Then
                    Unload Popups(UBound(Popups))
                    Set Popups(UBound(Popups)) = Nothing
                    ReDim Preserve Popups(UBound(Popups) - 1)
                    mnuPopup.MenuItem1(mnuPopup.curItem).SetFocus
                End If
            Case 38
aloop:
                If Not Index = 0 Then
                    If mnuPopup.MenuItem1(Index - 1).Enabled = False Or mnuPopup.MenuItem1(Index - 1).Visible = False Then
                        Index = Index - 1
                        GoTo aloop
                    End If
                    mnuPopup.MenuItem1_MouseMove Index - 1, 1, 0, 100, 100
                Else
                        Unload Popups(UBound(Popups))
                        Set Popups(UBound(Popups)) = Nothing
                        If Not UBound(Popups) = 0 Then ReDim Preserve Popups(UBound(Popups) - 1)
                    If Not Popups(0) Is Nothing Then
                        mnuPopup.MenuItem1(mnuPopup.curItem).SetFocus
                    End If
                End If
            Case 39
                If mnuPopup.MenuItem1(Index).IsPopup = True Then
                    mnuPopup.fireEvent mnuPopup.MenuItem1(Index).Caption, mnuPopup.MenuItem1(Index).Tag, CLng(Index)
                Else
                    ProcessKeys curItemM, 39, Shift
                End If
            Case 40
bloop:
                If Not Index = mnuPopup.MenuItem1.Count - 1 Then
                    If mnuPopup.MenuItem1(Index + 1).Enabled = False Or mnuPopup.MenuItem1(Index + 1).Visible = False Then
                        Index = Index + 1
                        GoTo bloop
                    End If
                    mnuPopup.MenuItem1_MouseMove Index + 1, 1, 0, 100, 100
                End If
        End Select
    End Select
End Sub

Private Sub MenuItem1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then ProcessKeys Index, KeyCode, Shift
End Sub

Public Sub ProcessKeys(Index As Integer, Key As Integer, Shift As Integer)
Select Case Shift
    Case 0
        Select Case Key
            Case 112 To 127 'if any F keys are pushed look for the shortcut
                For a = 0 To UBound(allMenus)
                    If LCase(allMenus(a).sCut) = LCase(TrsLteKey(Key)) Then
                        RaiseEvent Click(Replace(allMenus(a).Caption, "....", ""), allMenus(a).Tag)
                        Exit For
                    End If
                Next a
            Case 18
                MenuItem1_MouseMove curItemM, 1, 0, 100, 100
                MenuItem1(curItemM).SetFocus
            Case 37
aloop: 'ensures you dont stop on a disabled item
                If Not curItemM = 0 Then
                    If MenuItem1(curItemM - 1).Enabled = False Or MenuItem1(curItemM - 1).Visible = False Then
                        curItemM = curItemM - 1
                        GoTo aloop
                    End If
                    MenuItem1_MouseMove curItemM - 1, 1, 0, 100, 100
                Else
                    curItemM = MenuItem1.Count
                    GoTo aloop
                End If
            Case 39
bloop:
                If Not curItemM = MenuItem1.Count - 1 Then
                        If MenuItem1(curItemM + 1).Enabled = False Or MenuItem1(curItemM + 1).Visible = False Then
                            curItemM = curItemM + 1
                            GoTo bloop
                        End If
                    MenuItem1_MouseMove curItemM + 1, 1, 0, 100, 100
                Else
                    curItemM = 0
                    MenuItem1_MouseMove curItemM, 1, 0, 100, 100
                End If
            Case 13
                If Not MenuItem1(curItemM).Enabled = False Then MenuItem1_Click curItemM, MenuItem1(curItemM).Caption, MenuItem1(curItemM).Tag, MenuItem1(curItemM).IsPopup
        End Select
    Case 2 'if ctrl is pushed look for shortcut
        For a = 0 To UBound(allMenus)
            If LCase(allMenus(a).sCut) = LCase("ctrl + " & TrsLteKey(Key)) Then
                RaiseEvent Click(Replace(allMenus(a).Caption, "....", ""), allMenus(a).Tag)
                Exit For
            End If
        Next a
    Case 4 'if alt is pushed
        If Key <> 0 Then FireAccel Key
End Select
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Ambient.UserMode() Then Exit Sub
    GRADIENT_FILL_RECT_DIRECTION = mGradientDirection
    selIndex = -1
 '   UserControl_GotFocus
    cleanmenu
End Sub

Private Sub UserControl_Show()
If Not Ambient.UserMode() Then Exit Sub
'    Set MB = Me
'setup to catch the events of the parent form -------------------------------------
    If TypeOf UserControl.parent Is Form Then
        If TypeOf UserControl.parent Is MDIForm Then
            Set MDIparn = UserControl.parent
        Else
            Set parn = UserControl.parent
            parn.KeyPreview = True
            Set sc2 = New cSubclass
            Call sc2.Subclass(UserControl.parent.hWnd, Me)
            Call sc2.AddMsg(WinSubHook.WM_NCACTIVATE, MSG_BEFORE)

        End If
    End If
''----------------------------------------------------------------------------------
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)
'
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
    Select Case uMsg
        Case 134    'WM_NCACTIVATE
            If TypeOf UserControl.parent Is Form Then
                If Not TypeOf UserControl.parent Is MDIForm Then
                    If UserControl.parent.MDIChild = True Then
                        If wParam = 0 Then
                            UserControl.Extender.Visible = False
                        Else
                            UserControl.Extender.Visible = True
                        End If
                    Else
                        If Not Popups(0) Is Nothing Then wParam = 1
                    End If
                Else
                    If Not Popups(0) Is Nothing Then wParam = 1 'force the window to look active
                End If
            End If
        End Select
End Sub

'this is the parent form if the form isn't an mdiform
Private Sub parn_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Or Shift = 4 Then
        Call ProcessKeys(curItemM, KeyCode, Shift)
    End If
End Sub

Private Sub UserControl_Terminate()
'ensure we arn't still hooked shouldn't be as we unhook after each ues of the menu
    If hHook <> 0 Then Call UnhookWindowsHookEx(hHook) 'unhook first cause it was last to be hooked
    If Not MB Is Nothing Then Set MB = Nothing
    Erase allMenus
    Erase Popups
    If Not mnuPopup Is Nothing Then Set mnuPopup = Nothing
    If Not sc2 Is Nothing Then Set sc2 = Nothing
End Sub

Private Function IsIDE() As Boolean
    IsIDE = (App.LogMode = 0)
End Function

Private Sub UserControl_Resize()
'in the resizeing and placing of the main items allow for a wrapable menu
   UserControl.Height = 405 'force height
'Framed border sizing
    RaiseEvent Resize
    Select Case BorderStyle
        Case 0
            Line1(0).X1 = 0
            Line1(0).X2 = 0
            Line1(0).Y1 = 0
            Line1(0).Y2 = UserControl.Height - 15
            Line2(0).BorderWidth = 2
            Line2(0).X1 = 15
            Line2(0).X2 = 15
            Line2(0).Y1 = 0
            Line2(0).Y2 = UserControl.Height
            Line1(1).X1 = 0
            Line1(1).Y1 = 0
            Line1(1).Y2 = 0
            Line1(1).X2 = UserControl.Width '- 15
            Line2(1).BorderWidth = 1
            Line2(1).X1 = 15
            Line2(1).Y1 = 15
            Line2(1).Y2 = 15
            Line2(1).X2 = UserControl.Width
            Line1(2).X1 = UserControl.Width - 30
            Line1(2).X2 = UserControl.Width - 30
            Line1(2).Y1 = 15
            Line1(2).Y2 = UserControl.Height '- 15
            Line2(2).BorderWidth = 2
            Line2(2).X1 = UserControl.Width - 15
            Line2(2).X2 = UserControl.Width - 15
            Line2(2).Y1 = 15
            Line2(2).Y2 = UserControl.Height
            Line1(3).X1 = 0
            Line1(3).X2 = UserControl.Width - 15
            Line1(3).Y1 = UserControl.Height - 30
            Line1(3).Y2 = UserControl.Height - 30
            Line2(3).BorderWidth = 2
            Line2(3).X1 = 0
            Line2(3).X2 = UserControl.Width  '- 15
            Line2(3).Y1 = UserControl.Height - 15
            Line2(3).Y2 = UserControl.Height - 15
        Case 2 'sunken
            Line1(0).X1 = 0
            Line1(0).X2 = 0
            Line1(0).Y1 = 15
            Line1(0).Y2 = UserControl.Height
            Line1(1).X1 = 0
            Line1(1).X2 = UserControl.Width
            Line1(1).Y1 = 0
            Line1(1).Y2 = 0
            Line2(2).BorderWidth = 1
            Line2(2).X1 = UserControl.Width - 10
            Line2(2).X2 = UserControl.Width - 10
            Line2(2).Y1 = 15
            Line2(2).Y2 = UserControl.Height - 10
            Line2(3).BorderWidth = 1
            Line2(3).X1 = 15
            Line2(3).X2 = UserControl.Width - 10
            Line2(3).Y1 = UserControl.Height - 10
            Line2(3).Y2 = UserControl.Height - 10
        Case 1 'raised
            Line2(0).BorderWidth = 1
            Line2(0).X1 = 0
            Line2(0).X2 = 0
            Line2(0).Y1 = 15
            Line2(0).Y2 = UserControl.Height
            Line2(1).BorderWidth = 1
            Line2(1).X1 = 0
            Line2(1).X2 = UserControl.Width
            Line2(1).Y1 = 0
            Line2(1).Y2 = 0
            Line1(2).X1 = UserControl.Width - 10
            Line1(2).X2 = UserControl.Width - 10
            Line1(2).Y1 = 15
            Line1(2).Y2 = UserControl.Height - 10
            Line1(3).X1 = 15
            Line1(3).X2 = UserControl.Width - 10
            Line1(3).Y1 = UserControl.Height - 10
            Line1(3).Y2 = UserControl.Height - 10
    End Select
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    ForcePaint
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Sub Refresh()
    UserControl.Refresh
End Sub

Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

Public Property Get Border() As Boolean
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    m_Border = New_Border
    PropertyChanged "Border"
    If New_Border = True Then
            If LeftBorder = True Then
                Select Case BorderStyle
                    Case 2 'sunken
                        Line1(0).Visible = New_Border
                        Line2(0).Visible = False
                    Case 1 'raised
                        Line1(0).Visible = False
                        Line2(0).Visible = New_Border
                    Case 0 'fram
                        Line1(0).Visible = New_Border
                        Line2(0).Visible = New_Border
                End Select
            Else
                Line1(0).Visible = False
                Line2(0).Visible = False
            End If
        If TopBorder = True Then
            Select Case BorderStyle
                Case 2
                    Line1(1).Visible = New_Border
                    Line2(1).Visible = False
                Case 1
                    Line1(1).Visible = False
                    Line2(1).Visible = New_Border
                Case 0
                    Line1(1).Visible = New_Border
                    Line2(1).Visible = New_Border
            End Select
        Else
            Line1(1).Visible = False
            Line2(1).Visible = False
        End If
        If RightBorder = True Then
            Select Case BorderStyle
                Case 2
                    Line1(2).Visible = False
                    Line2(2).Visible = New_Border
                Case 1
                    Line1(2).Visible = New_Border
                    Line2(2).Visible = False
                Case 0
                    Line1(2).Visible = New_Border
                    Line2(2).Visible = New_Border
            End Select
        Else
            Line1(2).Visible = False
            Line2(2).Visible = False
        End If
        If BotBorder = True Then
            Select Case BorderStyle
                Case 2
                    Line1(3).Visible = False
                    Line2(3).Visible = New_Border
                Case 1
                    Line1(3).Visible = New_Border
                    Line2(3).Visible = False
                Case 0
                    Line1(3).Visible = New_Border
                    Line2(3).Visible = New_Border
            End Select
        Else
            Line1(3).Visible = False
            Line2(3).Visible = False
        End If
    Else
            Line1(0).Visible = New_Border
            Line2(0).Visible = New_Border
            Line1(1).Visible = New_Border
            Line2(1).Visible = New_Border
            Line1(2).Visible = New_Border
            Line2(2).Visible = New_Border
            Line1(3).Visible = New_Border
            Line2(3).Visible = New_Border
    End If
UserControl_Resize
End Property

Public Property Get TopBorder() As Boolean
    TopBorder = mTopBorder
End Property

Public Property Let TopBorder(NewValue As Boolean)
    mTopBorder = NewValue
    PropertyChanged "TopBorder"
        If Border = True Then
            Select Case BorderStyle
                Case 2
                    Line1(1).Visible = NewValue
                    Line2(1).Visible = False
                Case 1
                    Line1(1).Visible = False
                    Line2(1).Visible = NewValue
                Case 0
                    Line1(1).Visible = NewValue
                    Line2(1).Visible = NewValue
            End Select
        End If
End Property

Public Property Get LeftBorder() As Boolean
    LeftBorder = mLeftBorder
End Property

Public Property Let LeftBorder(NewValue As Boolean)
    mLeftBorder = NewValue
    PropertyChanged "LeftBorder"
        If Border = True Then
            Select Case BorderStyle
                Case 2
                    Line1(0).Visible = NewValue
                    Line2(0).Visible = False
                Case 1
                    Line1(0).Visible = False
                    Line2(0).Visible = NewValue
                Case 0
                    Line1(0).Visible = NewValue
                    Line2(0).Visible = NewValue
            End Select
        End If
End Property

Public Property Get RightBorder() As Boolean
    RightBorder = mRightBorder
End Property

Public Property Let RightBorder(NewValue As Boolean)
    mRightBorder = NewValue
    PropertyChanged "RightBorder"
        If Border = True Then
            Select Case BorderStyle
                Case 2
                    Line1(2).Visible = False
                    Line2(2).Visible = NewValue
                Case 1
                    Line1(2).Visible = NewValue
                    Line2(2).Visible = False
                Case 0
                    Line1(2).Visible = NewValue
                    Line2(2).Visible = NewValue
            End Select
        End If
End Property

Public Property Get BotBorder() As Boolean
    BotBorder = mBotBorder
End Property

Public Property Let BotBorder(NewValue As Boolean)
    mBotBorder = NewValue
    PropertyChanged "BotBorder"
        If Border = True Then
            Select Case BorderStyle
                Case 2 'sunk
                    Line1(3).Visible = False
                    Line2(3).Visible = NewValue
                Case 1
                    Line1(3).Visible = NewValue
                    Line2(3).Visible = False
                Case 0
                    Line1(3).Visible = NewValue
                    Line2(3).Visible = NewValue
            End Select
        End If
End Property

Private Sub UserControl_InitProperties()
    m_Border = m_def_Border
    newleft = 60
End Sub

'Load property values
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim parts() As String
    ItemCount = 0 'set number of "MainMenu MenuItems" to 0 and load them from scrach
    GradientDirection = PropBag.ReadProperty("GradientDirection", 0)
    BackColor = PropBag.ReadProperty("BackColor", &H8000000F) 'load menubar properties
    mChildren = PropBag.ReadProperty("Children", vbNullString)
    BotBorder = PropBag.ReadProperty("BotBorder", True)
    retainApp = PropBag.ReadProperty("retainApp", m_def_retainApp)
    RightBorder = PropBag.ReadProperty("RightBorder", True)
    LeftBorder = PropBag.ReadProperty("LeftBorder", True)
    TopBorder = PropBag.ReadProperty("TopBorder", True)
    Enabled = PropBag.ReadProperty("Enabled", True)
    mTotalItems = PropBag.ReadProperty("TotalItems", 0)
    Border = PropBag.ReadProperty("Border", m_def_Border)
    BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    ReDim allMenus(mTotalItems) 'reload the menu array
        For a = 0 To mTotalItems
            With allMenus(a)
                 .Caption = PropBag.ReadProperty("Caption" & a, "")
                 .bkcol = PropBag.ReadProperty("bkCol" & a, &H8000000F)
                 .Border = PropBag.ReadProperty("Border" & a, False)
                 .Check = PropBag.ReadProperty("Check" & a, False)
                 .Children = PropBag.ReadProperty("Children" & a, vbNullString)
                 .Drop = PropBag.ReadProperty("Drop" & a, False)
                 .frCol = PropBag.ReadProperty("frCol" & a, 0)
                 Set .Icon = PropBag.ReadProperty("Icon" & a, Nothing)
                 .ispop = PropBag.ReadProperty("isPop" & a, False)
                 .parent = PropBag.ReadProperty("Parent" & a, vbNullString)
                 .Path = PropBag.ReadProperty("Path" & a, vbNullString)
                 .Tag = PropBag.ReadProperty("Tag" & a, vbNullString)
                 .Tool = PropBag.ReadProperty("Tool" & a, vbNullString)
                 Set .Font = PropBag.ReadProperty("Font" & a, UserControl.Font)
                 .BStyle = PropBag.ReadProperty("BStyle" & a, 1)
                 .Visible = PropBag.ReadProperty("Visible" & a, True)
                 .Enabled = PropBag.ReadProperty("Enabled" & a, True)
                 .sCut = PropBag.ReadProperty("sCut" & a, vbNullString)
                 .Group = PropBag.ReadProperty("group" & a, 0)
                 .Style = PropBag.ReadProperty("Style" & a, -1)
                 .Gradiant = PropBag.ReadProperty("Gradiant" & a, 0)
                     If Not Left(.Caption, 1) = "." Then 'display the main menu items
                         If ItemCount > 0 Then Load MenuItem1(ItemCount)
                            If retainApp = True Then
                               MenuItem1(ItemCount).BackColor = allMenus(0).bkcol
                               MenuItem1(ItemCount).Gradiant = allMenus(0).Gradiant
                               MenuItem1(ItemCount).ForeColor = allMenus(0).frCol
                            Else
                               MenuItem1(ItemCount).BackColor = .bkcol
                               MenuItem1(ItemCount).Gradiant = .Gradiant
                               MenuItem1(ItemCount).ForeColor = .frCol
                            End If
                         MenuItem1(ItemCount).Border = .Border
                         MenuItem1(ItemCount).Checked = .Check
                         MenuItem1(ItemCount).Dropdown = .Drop
                         Set MenuItem1(ItemCount).picture = .Icon
                         MenuItem1(ItemCount).IsPopup = .ispop
                         MenuItem1(ItemCount).Tag = .Tag
                         Set MenuItem1(ItemCount).Font = .Font
                         MenuItem1(ItemCount).BStyle = .BStyle
                         MenuItem1(ItemCount).Enabled = .Enabled
                         MenuItem1(ItemCount).ToolTipText = .Tool
                         MenuItem1(ItemCount).Width = 100
                         MenuItem1(ItemCount).Visible = .Visible
                         MenuItem1(ItemCount).ShortCut = .sCut
                         MenuItem1(ItemCount).Caption = .Caption
                         MenuItem1(ItemCount).ToolTipText = .Tool
                         MenuItem1(ItemCount).Left = newleft
                            If .Visible = True Then
                               newleft = newleft + MenuItem1(ItemCount).Width + 30
                            End If
                         ItemCount = ItemCount + 1
                     End If
            End With
        Next a
    Close #1
    UserControl_Resize
   If hHook <> 0 Then Call UnhookWindowsHookEx(hHook)
End Sub
Public Sub Clear()
Attribute Clear.VB_MemberFlags = "40"
        If Not MenuItem1.Count = 1 Then
            For a = 1 To MenuItem1.Count - 1
                Unload MenuItem1(a)
                ItemCount = ItemCount - 1
            Next a
        End If
    MenuItem1(0).Width = 100
    MenuItem1(0).Enabled = True 'If we dont have the item enabled and visible we will error
    MenuItem1(0).Visible = True
    MenuItem1(0).SetFocus
End Sub

Public Sub Remove(item As String)
Dim ItemNum As Integer
        For b = 0 To UBound(allMenus) 'find items place in the array
            If Replace(allMenus(b).Caption, ".", "") = item Then
                ItemNum = b
                Exit For
            Else
                'We couldn't Find the item so eather we entered the caption wrong or it doesn't exist
                'eather way get out
                If b = UBound(allMenus) Then Exit Sub
            End If
        Next b
    
        If InStr(1, allMenus(ItemNum).Caption, ".") > 0 Then
            removeItem ItemNum
        Else
        'this item is a top level item so just hide it
                For b = 0 To MenuItem1.Count - 1
                    If MenuItem1(b).Caption = item Then
                        MenuItem1(b).Visible = False
                        allMenus(ItemNum).Visible = False
                        ResetItems
                        Exit For
                    End If
                Next b
        End If
End Sub

Public Sub Add(Optional parent As String = vbNullString, Optional itemType As Integer = 0, Optional Caption As String = "Caption", Optional Ck As Boolean = False, _
Optional Drop As Boolean = False, Optional Popup As Boolean = False, Optional mTag As String = vbNullString, _
Optional bkcol As OLE_COLOR = &H8000000F, Optional ForColor As OLE_COLOR = 0, _
Optional Font As Font = Nothing, Optional Border As Boolean = True, Optional Icon As picture = Nothing, _
Optional Visible As Boolean = True, Optional Enabled As Boolean = True, Optional sCut As String = vbNullString, _
Optional Group As Integer = 0, Optional Style As checkStyle = -1, Optional Gradiant As OLE_COLOR = 0, Optional BStyle As gStyle = aGraphical, Optional Tool As String = vbNullString)
'there isn't a seperator for the menu bar yet seems there are some issues useing the one I have
'now and use in the form without error, if I create more than one on the menu bar it isn't created
'and causes the one that exist to act strange.
'
'I have found that if you use an item with BStyle set to 0 and caption = " "(space) then it will serve as a spacer
'
Dim clipedCaption As String
Dim clipedPart As String
    If parent <> vbNullString Then
        For a = 0 To UBound(allMenus)
        If InStr(1, allMenus(a).Caption, ".") > 0 Then
            clipedCaption = Replace(allMenus(a).Caption, ".", "")
            clipedPart = Left(allMenus(a).Caption, Len(allMenus(a).Caption) - Len(clipedCaption)) & "...."
        Else
            clipedCaption = allMenus(a).Caption
            clipedPart = "...."
        End If
            If LCase(clipedCaption) = LCase(parent) Then
                InsertItem a + 1
                    With allMenus(a + 1)
                        .Caption = clipedPart & Caption
                        .bkcol = bkcol
                        .frCol = ForColor
                        .Gradiant = Gradiant
                        .Check = Ck
                        .Drop = Drop
                        If Not Font Is Nothing Then Set .Font = Font
                        .ispop = Popup  'if it has a popup form if it is true ismain will be false
                        Set .Icon = Icon
                        .Enabled = Enabled
                        .Tool = Tool
                        .Visible = Visible
                        .Border = Border
                        .sCut = sCut
                        .Tag = mTag
                        .Group = Group
                        .Style = Style
                        .BStyle = BStyle
                    End With
                Exit For
            End If
        Next a
    Else
        If itemType <> 0 Then
        Else
                    If ItemCount > 0 Then Load MenuItem1(ItemCount)
                With MenuItem1(ItemCount)
                    .Caption = Caption
                        If retainApp = True Then
                            .BackColor = allMenus(0).bkcol
                            .ForeColor = allMenus(0).frCol
                            .Gradiant = allMenus(0).Gradiant
                        Else
                            .BackColor = bkcol
                            .ForeColor = ForColor
                            .Gradiant = Gradiant
                        End If
                    .Checked = Ck
                    .Dropdown = Drop
                    .Border = Border
                    If Not Font Is Nothing Then Set .Font = Font
                        'not sure how to get the default value for a font yet using the font of usercontrol if none found
                    .IsPopup = Popup 'if it has a popup form if it is true ismain will be false
                    .Width = 100
                        If ItemCount = 0 Then
                            .Left = 60
                        Else
                            .Left = MenuItem1(ItemCount - 1).Left + MenuItem1(ItemCount - 1).Width + 15
                        End If
                    Set .picture = Icon
                    .Enabled = Enabled
                    .Tag = mTag
                    .ToolTipText = Tool
                    .Visible = Visible
                    .ShortCut = sCut
                    .BStyle = BStyle
                End With
            ItemCount = ItemCount + 1
        End If
    End If
End Sub

'Write property values
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("retainApp", mretainApp, m_def_retainApp)
    Call PropBag.WriteProperty("GradientDirection", GradientDirection, 0)
    Call PropBag.WriteProperty("BotBorder", mBotBorder, m_def_BotBorder)
    Call PropBag.WriteProperty("RightBorder", mRightBorder, m_def_RightBorder)
    Call PropBag.WriteProperty("LeftBorder", mLeftBorder, m_def_LeftBorder)
    Call PropBag.WriteProperty("TopBorder", mTopBorder, m_def_TopBorder)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("TotalItems", UBound(allMenus), 0)
    Call PropBag.WriteProperty("Border", m_Border, m_def_Border)
    Call PropBag.WriteProperty("BorderStyle", BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Children", mChildren, vbNullString)
        For a = 0 To UBound(allMenus)
            With allMenus(a)
                    Call PropBag.WriteProperty("Caption" & a, .Caption, "")
                    Call PropBag.WriteProperty("bkCol" & a, .bkcol, &H8000000F)
                    Call PropBag.WriteProperty("Border" & a, .Border, False)
                    Call PropBag.WriteProperty("Check" & a, .Check, False)
                    Call PropBag.WriteProperty("Children" & a, .Children, vbNullString)
                    Call PropBag.WriteProperty("Drop" & a, .Drop, False)
                    Call PropBag.WriteProperty("Enabled" & a, .Enabled, True)
                    Call PropBag.WriteProperty("frCol" & a, .frCol, 0)
                    Call PropBag.WriteProperty("Icon" & a, .Icon, Nothing)
                    Call PropBag.WriteProperty("Font" & a, .Font, MenuItem1(0).Font)
                    Call PropBag.WriteProperty("isPop" & a, .ispop, False)
                    Call PropBag.WriteProperty("Parent" & a, .parent, vbNullString)
                    Call PropBag.WriteProperty("Path" & a, .Path, vbNullString)
                    Call PropBag.WriteProperty("Tag" & a, .Tag, vbNullString)
                    Call PropBag.WriteProperty("Tool" & a, .Tool, vbNullString)
                    Call PropBag.WriteProperty("Visible" & a, .Visible, True)
                    Call PropBag.WriteProperty("sCut" & a, .sCut, vbNullString)
                    Call PropBag.WriteProperty("group" & a, .Group, 0)
                    Call PropBag.WriteProperty("Style" & a, .Style, -1)
                    Call PropBag.WriteProperty("Gradiant" & a, .Gradiant, 0)
                    Call PropBag.WriteProperty("BStyle" & a, .BStyle, 1)
            End With
        Next a
End Sub

Public Sub PopupMenu(Caption As String) 'added suport for a right click menu to be displayed
'It does the same thing as the functions above but puts the popup at the mouse pointer
'no code yet to insure it doesn't go beyond the edges of the screen.you can also popup any
'menu that has child items.
'same as mnuPopup_Clicked above
Dim MouseP As POINTAPI
Call GetCursorPos(MouseP)
GradientDirection = mGradientDirection
    unloadAll
    ReDim Popups(0)
    For a = 0 To UBound(allMenus) 'find the item in question and load it as though it was a main menu item
        If Replace(allMenus(a).Caption, ".", "") = Caption Then
        RaiseEvent Click(Replace(allMenus(a).Caption, ".", ""), allMenus(a).Tag)
            If allMenus(a).ispop = True Then
            If hHook = 0 Then hHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0&, App.ThreadID)
                activeMenu = True
                selIndex = 0
                Set Popups(0) = New frmPopup
                Set mnuPopup = Popups(0)
                With Popups(0)
                        If InStr(1, allMenus(a).Caption, ".") > 0 Then
                            FirstPart = Left(allMenus(a).Caption, InStrRev(allMenus(a).Caption, "."))
                        Else
                            FirstPart = ""
                        End If
                        For b = a + 1 To UBound(allMenus)
                            SearchFor = Left(allMenus(b).Caption, InStrRev(allMenus(b).Caption, "."))
                                If SearchFor = FirstPart & "...." Then
                                    If Replace(allMenus(b).Caption, ".", "") = "-" Then
                                        If retainApp = True Then
                                            .Add 1, allMenus(b).Caption, allMenus(b).Check, allMenus(b).Drop, allMenus(b).ispop, allMenus(b).Tag, allMenus(0).bkcol, allMenus(0).frCol, allMenus(b).Font, allMenus(b).Border, allMenus(b).Icon, allMenus(b).Visible, allMenus(b).Enabled, allMenus(b).sCut, allMenus(0).Gradiant, allMenus(b).BStyle, allMenus(b).Tool
                                        Else
                                            .Add 1, allMenus(b).Caption, allMenus(b).Check, allMenus(b).Drop, allMenus(b).ispop, allMenus(b).Tag, allMenus(b).bkcol, allMenus(b).frCol, allMenus(b).Font, allMenus(b).Border, allMenus(b).Icon, allMenus(b).Visible, allMenus(b).Enabled, allMenus(b).sCut, allMenus(b).Gradiant, allMenus(b).BStyle, allMenus(b).Tool
                                        End If
                                    Else
                                        If retainApp = True Then
                                            .Add , allMenus(b).Caption, allMenus(b).Check, allMenus(b).Drop, allMenus(b).ispop, allMenus(b).Tag, allMenus(0).bkcol, allMenus(0).frCol, allMenus(b).Font, allMenus(b).Border, allMenus(b).Icon, allMenus(b).Visible, allMenus(b).Enabled, allMenus(b).sCut, allMenus(0).Gradiant, allMenus(b).BStyle, allMenus(b).Tool
                                        Else
                                            .Add , allMenus(b).Caption, allMenus(b).Check, allMenus(b).Drop, allMenus(b).ispop, allMenus(b).Tag, allMenus(b).bkcol, allMenus(b).frCol, allMenus(b).Font, allMenus(b).Border, allMenus(b).Icon, allMenus(b).Visible, allMenus(b).Enabled, allMenus(b).sCut, allMenus(b).Gradiant, allMenus(b).BStyle, allMenus(b).Tool
                                        End If
                                    End If
                                ElseIf SearchFor = FirstPart Then
                                    Exit For
                                End If
                        Next b
                    .Left = MouseP.X * Screen.TwipsPerPixelX
                    .Top = MouseP.Y * Screen.TwipsPerPixelY
                    .Show , UserControl
                    HighlightFirstEnabled Popups(0)
                'Call SendMessage(UserControl.parent.hWnd, WM_NCACTIVATE, 1, 0)
                End With
            End If
        Exit For 'to keep from loading dup menus exit here
        End If
    Next a
End Sub

Public Sub PopupMenuGroup(Group As Integer) 'added suport for a right click menu to be displayed
'It does the same thing as the functions above but puts the popup at the mouse pointer
'no code yet to insure it doesn't go beyond the edges of the screen.you can also popup any
'menu that has child items.
'same as mnuPopup_Clicked above
Dim MouseP As POINTAPI
Dim itemsloaded As Integer
GradientDirection = mGradientDirection
itemsloaded = 0
If hHook = 0 Then hHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0&, App.ThreadID)
Call GetCursorPos(MouseP)
    unloadAll
    ReDim Popups(0)
    activeMenu = True
    selIndex = 0
    Set Popups(0) = New frmPopup
    Set mnuPopup = Popups(0)
    Load mnuPopup
        For a = 0 To UBound(allMenus) 'find the item in question and load it as though it was a main menu item
            If allMenus(a).Group = Group Then
                With Popups(0)
                    If Replace(allMenus(a).Caption, ".", "") = "-" Then
                        If retainApp = True Then
                            .Add 1, allMenus(a).Caption, allMenus(a).Check, allMenus(a).Drop, allMenus(a).ispop, allMenus(a).Tag, allMenus(0).bkcol, allMenus(0).frCol, allMenus(a).Font, allMenus(a).Border, allMenus(a).Icon, allMenus(a).Visible, allMenus(a).Enabled, allMenus(a).sCut, allMenus(0).Gradiant, allMenus(a).BStyle, allMenus(a).Tool
                        Else
                            .Add 1, allMenus(a).Caption, allMenus(a).Check, allMenus(a).Drop, allMenus(a).ispop, allMenus(a).Tag, allMenus(a).bkcol, allMenus(a).frCol, allMenus(a).Font, allMenus(a).Border, allMenus(a).Icon, allMenus(a).Visible, allMenus(a).Enabled, allMenus(a).sCut, allMenus(a).Gradiant, allMenus(a).BStyle, allMenus(a).Tool
                        End If
                    Else
                        If retainApp = True Then
                            .Add , allMenus(a).Caption, allMenus(a).Check, allMenus(a).Drop, allMenus(a).ispop, allMenus(a).Tag, allMenus(0).bkcol, allMenus(0).frCol, allMenus(a).Font, allMenus(a).Border, allMenus(a).Icon, allMenus(a).Visible, allMenus(a).Enabled, allMenus(a).sCut, allMenus(0).Gradiant, allMenus(a).BStyle, allMenus(a).Tool
                        Else
                            .Add , allMenus(a).Caption, allMenus(a).Check, allMenus(a).Drop, allMenus(a).ispop, allMenus(a).Tag, allMenus(a).bkcol, allMenus(a).frCol, allMenus(a).Font, allMenus(a).Border, allMenus(a).Icon, allMenus(a).Visible, allMenus(a).Enabled, allMenus(a).sCut, allMenus(a).Gradiant, allMenus(a).BStyle, allMenus(a).Tool
                        End If
                    End If
                End With
            End If
        Next a
    Popups(0).Left = MouseP.X * Screen.TwipsPerPixelX
    Popups(0).Top = MouseP.Y * Screen.TwipsPerPixelY
    Popups(0).Show , UserControl
    HighlightFirstEnabled Popups(0)
End Sub

Private Sub HighlightFirstEnabled(frm As frmPopup) 'selects the first enabled and visible item in the popup if it isn't a popup menu
    For a = 0 To frm.MenuItem1.Count - 1
        If frm.MenuItem1(a).Enabled = True And frm.MenuItem1(a).Visible = True And Not frm.MenuItem1(a).IsPopup = True Then
            frm.LastBack = frm.MenuItem1(a).BackColor
            frm.LastFont = frm.MenuItem1(a).ForeColor
            frm.lastIndex = a
            frm.MenuItem1(a).BackColor = frm.MenuItem1(a).ForeColor '&H8000000D
            frm.MenuItem1(a).ForeColor = frm.LastBack '&H80000009
            frm.MenuItem1(a).SetFocus
            frm.curItem = a
            Exit For
        End If
    Next a
End Sub

Private Sub ResetItems()
newleft = 60
    For a = 0 To MenuItem1.Count - 1
        MenuItem1(a).Left = newleft
            If MenuItem1(a).Visible = True Then
                newleft = newleft + MenuItem1(a).Width + 30
            End If
    Next a
End Sub

Public Sub ForcePaint()
    For a = 0 To MenuItem1.Count - 1
        MenuItem1(a).Caption = MenuItem1(a).Caption
    Next a
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

