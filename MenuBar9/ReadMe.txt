'Welcome to The new and improved Menu Control.
'Created By Russell Sanders 12/2006
'
'  In the past two months I have put close to 250 hours in this control. I thank you all for looking
'	at it and welcome any comment good or bad.
'  This works well for small apps.
'
'   PSC for a place to get help and code on anything. And to all the dedicated users of PSC
'       who post there thoughts there advice and criticisms on code, Thanks.
'   other credits are with the relative code.
'
'  This uses Pauls' safe subclass Code Found At: 
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=42918&lngWId=1
'
'Note: TO MAKE THIS WORK
'
'First register the type library
'	Type Lib\WinSubHook.tlb
'
'to do that,
' from the Project\References dialog in VB 'Browse' to this file
'	"WinSubHook.tlb" it's in App.Path
'
'You may want to first move it to your windows system folder so you will know where
' to find it.
'
'to get a full view of the workings of the subclasser you should get his code and also vote
'
'****************************************************************************************
'I'm creating this menu to be able to have a menu with any container control
'and be able to add icons to the menus
'The way it works at least for now
'I have created a user control named "MenuItem" a menuitem will be the controls
'on the mainmenu and on the popup forms.
'The "MainMenu" is the name of the next control which is as the name sugest
'the main menu bar.
'The Popup Form is what will be displayed when a menuItem has submenu items.
'
'when the menu is loaded all menu item properties will be stored in an array "allMenus"
'as they are loaded any mainmenu menuitem(Item located on the menuBar) will also
'be loaded and displayed. When the user selects an menuitem an event is
'fired, giving the user the time here to set properties of the items in a popup.
'if that menuitem activates a popup/dropdown menu then the popup form is loaded
'and any child items are displayed in the form(s). The forms are also an array
'although they can't catch each others events any popup form(menu) can close another
'popup form(menu). It's this way because I can only catch the events from one popup form
'set to be the last object in the array "popups()".
'lets just say you open a menu popup in it open another popup and from there open
'another. You now have three popups visible if you click the last one the event is
'passed to the mainMenu through mnuPopup popup events; but, if you click the first
'popup window there is nowhere for that event to go and the click would go unnoticed
'Therefore, In each popup form, in the mouse down event of the menuitem, a check is made
'to see if the last popup window is equal to the popup clicked. If it's not it's unloaded
'in the unload event of that popup being unloaded the next popup in the array is set to
'the active popup. the cycle then continues until the active popup is the popup you
'clicked in and the events are fired as normal from there.
'each menuitem on the menubar and on the popup forms is controlled
'by the menuBar(MainMenu). Local Properties are handled by the MenuItem Control
' The Back color will be drawn by the "MenuItem" but the "MainMenu" controls What Color
'
'
'Things you will see
'Array:
'   add/remove item anywhere in the array
'   move items up/down in the array
'
'save and load properties for a control that doesn’t exist
'   save properties code for a control you will create at runtime
'   load those properties to the control once it's created.
'
'   allowing the property page to read/write the properties of the non existent controls.
'
'   and creating runtime access to the properties of the non existent controls.
'
'   drawing items(pictures and text) on a picturebox
'
'   windows hooks: hooking for mouse up events
'
'   subclassing: subclass the menuitems for the mouse hover event
'
'Description:
'    This menu control functions the same as vbs' menu. With a few new features
'    it 's a control that can be placed on any container.
'
'Features:
'    Allows Icons in each menu Item.
'        BStyle is a new property I am testing that allows you to hide the icon
'
'    Backcolor,forecolor, and gradient.
'        there is a retainApp property if set true will cause all menu items fonts'
'        backcolor, and forecolor to be the same as the first item in the list.
'        the gradient background if it's set to 0(Black) = no gradient
'	NOTE: the selected colors are the same as the backcolor and forecolor
'		just switched.
'	      the gradients have left to right, top to bottom and center out settings
'
'    Shortcuts
'        Added support for shortcut keys but only for "ctrl + A - Z" and "F1 - F16"
'        you also have use of menu keys alt + a - z for your menu captions
'
'        NOTE: To catch the keys without hooking the keyboard You'll need to add this line to
'        The Keydown Event of the form its on:
'
'            If Shift = 2 Or Shift = 4 Then MainMenu1.ProcessKeys MainMenu1.curItem, KeyCode, Shift
'            "if the name of your menu control is MainMenu1" else replace MainMenu1
'            with the name of your menu control.
'
'    Accelerator keys "alt" + a key corresponding to an item in the main menu or a visible popup.
'
'    Popups
'        There is also a popup menu that works the same as a standard menu
'        on a right click event, call ObjectName.PopupMenu "Item Caption" and a popup will
'        be displayed at the current mouse cord.
'
'    Popup Group
'        Another Function Added Is A PopupMenuGroup This allows you to show items from different
'        popup menus in a popup together. you can even include top level menu items in the popup.
'
'    seperator bar with caption: The seperator bar and menuitem controls have both been updated
'	and reconstructed. A seperator is denoted with a hyphen in the caption property
'	so I use the tooltip property to store the caption(text displayed on the seperator).
'	the seperator bar has backcolor, forecolor, gradient, and font properties all optional
'	if you don't set them you will just have a gray item with a line through it.
'
'MainMenu Properties:
'    TotalItems: this is a count of the total number of menu items in the main menu including any
'        popup items and there children. Used when loading the menu to set the loop counter.
'
'    Backcolor: sets the backcolor of the menu bar

'    BorderStyle: allows you to have raised, sunken, or framed style borders

'    Border: sets the border on or off
'        The items bellow are subsets of border and will make no change if border is set to false
'        RightBorder: shows or hides the right border. has no effect if border is false
'        BottomBorder: the other items do the same as the first but for there respective properties
'        LeftBorder
'        TopBorder
'
'    retainApp: this property when set to true will force all items in the menu to share
'        the ForeColor,Backcolor, and Gradient with the first item in the menu.
'
'    Children: This item was used to hold the names of any menu items that were to be displayed
'        when the menu was loaded. Its' use has been abandon but the property remains in case I
'        find another use for it.
'
'    GradientDirection: allows you to have horizontal, vertical, horizontal center out, vertical center out
'
'MenuItem Properties:
'
'
'    Path As String 'isn't used It was intended to hold the path to the item example: file\open
'
'    Children As String 'isn't used this was intended to hold the names of the child menu items
'        and possibly speed up the loading of the items in a popup
'
'    Caption As String 'holds the caption of the control
'
'    tag As String     'will hold any string you like
'
'    ispop As Boolean  'tells the code to load a popup menu. This property must be set to true
'        if the menu item has children. If set to true without children a popup
'        will be displayed with no items.
'
'    Drop As Boolean   'visual aid to a dropdown menu. This allows you to indicate to your users
'        that an item is a popup
'
'    Check As Boolean  'check option on/off
'
'    Style As checkStyle 'none, checked or option style the option style only allows one item in the group to be checked
'        The checked style if set will toggle the checked state of an item. You can still set the checked
'        unchecked state through code without regard to the style.
'
'    Icon As Picture   'picture for menu item
'
'    tool As String    'tool tip text: this property allows a tool tip for the items.
'        this doesn’t 't work right now. I'm working on it; but, haven't gotten past it yet. If an item is disabled
'        its ' tooltip works but otherwise no go. If anyone has any Idea let me know. I thought it was highlighting
'        and setting focus to the item that kept the tooltip from showing; but, testing determined this wasn't
'        the case
'	NOTE: The tooltip is used to store the caption of seperator. The caption property is used to indicate
'		the item is a seperator.
'
'    bkcol As OLE_COLOR 'backcolor
'    frCol As OLE_COLOR 'fore color
'    Gradient As OLE_COLOR 'Gradient 0 = noGradient
'
'    Border As Boolean  'border on/off. This property was included for the main menu items. As the mouse is moved
'        over an item its' border is set to true. And if the mouse is moved out its' border is set to false
'        so setting this property for top level items is useless. However, if you use it with the popup items
'        it will function as expected.
'
'    Enabled As Boolean  'enable/disable with support for disabled icons
'
'    Visible As Boolean  'show/hide
'
'    parent As String    'This item isn't used in the menu item. it was intended to allow faster loading of the items
'        by checking the property and matching it with the item clicked but the loop was the same;
'        therefore the idea was abandon. I left the property just in case I(the user) need another string
'        container in the menu item.
'
'    Font As Font       'font 'although the font property is working there are quirks in that the first menu item
'		will somehow obtain the font of another item who’s font property has been changed. This requires
'		a little more study.
'
'    sCut As String     'short cut key combo This property Allows you to have Shortcut Keys
'        Ranging from "ctrl + A" To "ctrl + Z" and all the F keys "F1" To "F16"
'
'    Group As Integer   'allows you to have groups of items: This option will allow you to popup
'        a menu from anywhere in the array. You can have the "save" item from the "file Menu"
'        dropdown in the same popup as undo/redo from the edit dropdown. You can also
'        include Top level items in the popup.
'
'    BStyle As gStyle 'Picture or no picture: This item if set on a top level item will allow you
'        to hide the icon. but does not work for child menu items. A child items icon will
'        be displayed regardless of this setting.
'
'
'Seperator Properties
'    Orientation As Long 'Not used yet meant to set the seperator vertical or horizontal
'    BackColor As OLE_COLOR 'the backcolor forecolor and gradient are mates to the menuitem 
'    Gradient As OLE_COLOR   'and will share properties with the first item if retainApp = true
'    ForeColor As OLE_COLOR
'    Caption As String      'this will allow you to have a caption on your separators
'			Note if the caption is longer than the seperator is wide
'			the caption will not display correctly.
'
'    Sorry; but, I had to resort to Hooking into the message stream to catch mouse
'    left button up.Please be
'    careful. You will crash VB and/or hang your system. I advise you to comment the
'    two places where the hook is installed when testing in the IDE. Although, there
'    is a lot more to know about hooking. I have only included what is needed for
'    function. I am trying to keep things as simple as possible in the code. If you
'    need other info about hooking, this hooking code is from MSDN, search for
'    “hooking controls” and I will yield any help I can. I'll also take any advice
'    about the function of the control, and ideas to increase speed or usability,
'    very seriously.
'
''*************************************************************************************************************
'
'Read The comments in the code.
'
'Enjoy The control and please give feedback so I can improve it, Thanks in advance.
'
' made some speed improvements and cleaned the code a little.
'
'Quirks:
'
'Although this will work in an MDI Form it will not be affected by a child form menu
'meaning any menu on the child form want be displayed in the parent menu.
'
'
' if you open the property page and make a change select apply or ok the menu is loaded
'   in the form and it's active, clicking on an item of the menu in the ide will load the popup
'   or whatever action would take place at run time. This hasn't caused a problem yet but I 
'   would like to know why it does this.
'
'The title bar of the container form looses focus when a popup is displayed
'
'If you have more than one menu on a form and you have a popup visible, without selecting
'  an item mouse over a popup menuitem in the second menu bar on the form the popup sometimes
'  isn't displayed in the right position.
'
' I'm hooking for the left mouse Up; therefore, if you have a popup visible
'   and click off its' area with the right button nothing happens. If I set the hook
'   to catch the right button, there arises problems then displaying the popups.
'
' Saving the Font property for my controls for some reason
'   Fonts aren’t working right.
'
' If you add a menu item at runtime at current state
'   The item is always added as the first item in the popup menu.
'
'trouble with the way I save the icons and all are saved with a gray background
'the work around redraws the icon excluding the grayed background,
'at a cost of about three times the load time.
'
'lots of unnecessary redrawing in the pictureboxs causing flicker
'
