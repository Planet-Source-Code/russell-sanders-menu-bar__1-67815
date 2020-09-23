VERSION 5.00
Object = "*\AMenuBar.vbp"
Begin VB.Form frmTest 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin MenuBar.MainMenu MainMenu3 
      Height          =   405
      Left            =   240
      TabIndex        =   5
      Top             =   3330
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   714
      retainApp       =   -1  'True
      GradientDirection=   1
      BackColor       =   16711935
      TotalItems      =   3
      BorderStyle     =   1
      Children        =   "Test Menu Function\Test2"
      Caption0        =   "Test Menu Function"
      bkCol0          =   255
      Icon0           =   "frmTest.frx":0000
      isPop0          =   -1  'True
      Gradiant0       =   65535
      Caption1        =   "....Testing"
      Icon1           =   "frmTest.frx":0352
      Tool1           =   "Testing Tool Tips"
      sCut1           =   "Ctrl + T"
      Caption2        =   "....This is another"
      Check2          =   -1  'True
      Icon2           =   "frmTest.frx":06A4
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BStyle2         =   0
      Caption3        =   "Test2"
      Icon3           =   "frmTest.frx":09F6
      BeginProperty Font3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Just to show the abillity to have many menus"
      Height          =   1035
      Left            =   3570
      TabIndex        =   3
      Top             =   2790
      Width           =   4545
      Begin MenuBar.MainMenu MainMenu2 
         Height          =   405
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   714
         retainApp       =   -1  'True
         GradientDirection=   1
         BackColor       =   16777088
         TotalItems      =   4
         Children        =   "File\Address Book"
         Caption0        =   "File"
         bkCol0          =   8388608
         frCol0          =   255
         Icon0           =   "frmTest.frx":0D48
         isPop0          =   -1  'True
         Gradiant0       =   16777088
         BStyle0         =   0
         Caption1        =   "....Open"
         frCol1          =   8454143
         Icon1           =   "frmTest.frx":118A
         Tool1           =   "Open a file"
         Caption2        =   "....-"
         Icon2           =   "frmTest.frx":14DC
         BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tool2           =   "Testing"
         Caption3        =   "....New"
         Icon3           =   "frmTest.frx":182E
         BeginProperty Font3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption4        =   "Address Book"
         Check4          =   -1  'True
         Drop4           =   -1  'True
         Icon4           =   "frmTest.frx":1B80
         BeginProperty Font4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   4260
      TabIndex        =   4
      Text            =   "Right Click on me for a custom popup "
      Top             =   390
      Width           =   4005
   End
   Begin MenuBar.MainMenu MainMenu1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   714
      retainApp       =   -1  'True
      GradientDirection=   1
      BackColor       =   128
      TotalItems      =   28
      BorderStyle     =   1
      Children        =   "&File\Pictures\ \&Address Book\Email\Project\Disabled Item"
      Caption0        =   "&File"
      bkCol0          =   255
      frCol0          =   8388608
      Icon0           =   "frmTest.frx":1ED2
      isPop0          =   -1  'True
      Tool0           =   "Gives you a list of options you can use with assoiated files."
      Gradiant0       =   8454143
      BStyle0         =   0
      Caption1        =   "....&New Item For Testing"
      bkCol1          =   255
      Check1          =   -1  'True
      Drop1           =   -1  'True
      Enabled1        =   0   'False
      frCol1          =   16777215
      Icon1           =   "frmTest.frx":2224
      Tool1           =   "This is Just A new Item I'm testing"
      sCut1           =   "Ctrl + D"
      Gradiant1       =   16777088
      Caption2        =   "....-"
      Icon2           =   "frmTest.frx":2576
      Tool2           =   "This Test"
      BStyle2         =   0
      Caption3        =   "....Open"
      bkCol3          =   10485760
      frCol3          =   16777215
      Icon3           =   "frmTest.frx":28C8
      Tool3           =   "Open an assoiated document"
      sCut3           =   "Ctrl + O"
      Gradiant3       =   16777088
      BStyle3         =   0
      Caption4        =   "....&Save As"
      bkCol4          =   255
      Drop4           =   -1  'True
      frCol4          =   16777215
      Icon4           =   "frmTest.frx":2C1A
      isPop4          =   -1  'True
      Tool4           =   "This is a list of Save options"
      Gradiant4       =   16777088
      Caption5        =   "........Save Doc"
      bkCol5          =   255
      frCol5          =   16777215
      Icon5           =   "frmTest.frx":2F6C
      Tool5           =   "Save the document to it's default location and name."
      sCut5           =   "Ctrl + D"
      Gradiant5       =   16777088
      Caption6        =   "........Save Doc As"
      bkCol6          =   255
      Check6          =   -1  'True
      frCol6          =   16777215
      Icon6           =   "frmTest.frx":32BE
      Tool6           =   "Allows you to choose what name and file type."
      sCut6           =   "Ctrl + A"
      Gradiant6       =   16777088
      Caption7        =   "........-"
      Icon7           =   "frmTest.frx":3610
      BeginProperty Font7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tool7           =   "This is a test"
      Caption8        =   "........Save As Html"
      bkCol8          =   255
      frCol8          =   16777215
      Icon8           =   "frmTest.frx":3962
      Tool8           =   "This saves the document as a web page that can be viewed in a browser."
      sCut8           =   "Ctrl + H"
      Gradiant8       =   16777088
      Caption9        =   "........Test55"
      bkCol9          =   255
      Drop9           =   -1  'True
      frCol9          =   16777215
      Icon9           =   "frmTest.frx":3CB4
      BeginProperty Font9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      isPop9          =   -1  'True
      Gradiant9       =   16777088
      Caption10       =   "............Test44"
      bkCol10         =   255
      frCol10         =   16777215
      Icon10          =   "frmTest.frx":4006
      BeginProperty Font10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Gradiant10      =   16777088
      Caption11       =   "............Test33"
      bkCol11         =   255
      Drop11          =   -1  'True
      frCol11         =   16777215
      Icon11          =   "frmTest.frx":4358
      BeginProperty Font11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      isPop11         =   -1  'True
      Gradiant11      =   16777088
      Caption12       =   "................Test22"
      bkCol12         =   255
      frCol12         =   16777215
      Icon12          =   "frmTest.frx":46AA
      BeginProperty Font12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Gradiant12      =   16777088
      Caption13       =   "................Test111"
      bkCol13         =   255
      frCol13         =   16777215
      Icon13          =   "frmTest.frx":49FC
      BeginProperty Font13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Gradiant13      =   16777088
      Caption14       =   "Pictures"
      bkCol14         =   16711680
      frCol14         =   16777215
      Icon14          =   "frmTest.frx":4D4E
      BeginProperty Font14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      isPop14         =   -1  'True
      group14         =   1
      Gradiant14      =   16777088
      BStyle14        =   0
      Caption15       =   "....Import Pictures From Net"
      Check15         =   -1  'True
      Drop15          =   -1  'True
      Icon15          =   "frmTest.frx":50A0
      BeginProperty Font15 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      isPop15         =   -1  'True
      Caption16       =   "........T&his Is The First Item"
      Icon16          =   "frmTest.frx":53F2
      BeginProperty Font16 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption17       =   "....Testing Group Menu"
      Icon17          =   "frmTest.frx":5744
      BeginProperty Font17 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      sCut17          =   "Ctrl + T"
      group17         =   1
      Style17         =   1
      Caption18       =   " "
      Enabled18       =   0   'False
      Icon18          =   "frmTest.frx":5A96
      BeginProperty Font18 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BStyle18        =   0
      Caption19       =   "&Address Book"
      bkCol19         =   16711680
      frCol19         =   16777215
      Icon19          =   "frmTest.frx":66E8
      BeginProperty Font19 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Gradiant19      =   16776960
      BStyle19        =   0
      Caption20       =   "Email"
      bkCol20         =   16711680
      frCol20         =   16777215
      Icon20          =   "frmTest.frx":6A3A
      BeginProperty Font20 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Gradiant20      =   16776960
      BStyle20        =   0
      Caption21       =   "Project"
      bkCol21         =   16711680
      frCol21         =   16777215
      Icon21          =   "frmTest.frx":6D8C
      BeginProperty Font21 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      isPop21         =   -1  'True
      Gradiant21      =   16776960
      BStyle21        =   0
      Caption22       =   "....Add"
      Drop22          =   -1  'True
      Icon22          =   "frmTest.frx":70DE
      BeginProperty Font22 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      isPop22         =   -1  'True
      Caption23       =   "........Form"
      Icon23          =   "frmTest.frx":7430
      BeginProperty Font23 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      group23         =   1
      Style23         =   1
      Caption24       =   "........Class"
      Icon24          =   "frmTest.frx":7782
      BeginProperty Font24 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      group24         =   1
      Style24         =   0
      Caption25       =   "........UserControl"
      Check25         =   -1  'True
      Icon25          =   "frmTest.frx":7AD4
      BeginProperty Font25 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      group25         =   1
      Style25         =   1
      Caption26       =   "........-"
      Icon26          =   "frmTest.frx":7E26
      BeginProperty Font26 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      group26         =   1
      Caption27       =   "........Object"
      Icon27          =   "frmTest.frx":8178
      BeginProperty Font27 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      group27         =   1
      Style27         =   1
      Caption28       =   "Disabled Item"
      bkCol28         =   16711680
      Check28         =   -1  'True
      Drop28          =   -1  'True
      Enabled28       =   0   'False
      frCol28         =   16777215
      Icon28          =   "frmTest.frx":84CA
      BeginProperty Font28 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Gradiant28      =   16776960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide the Menu Bar "
      Height          =   345
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   1020
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   10290
      Picture         =   "frmTest.frx":881C
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   60
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Menu Items"
      Height          =   345
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "Right Click Me"
      Height          =   1635
      Left            =   90
      TabIndex        =   8
      Top             =   540
      Width           =   3615
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'testing the add item feature
Private Sub Command1_Click()
'Note:
'you can now add items to any menu or submenu including the main menubar.
'keep in mind however this has only limited testing and may not provide the desired results.
'known issues
'1) even though you can add items with the same caption you will only be getting the properties of the
'   first item found with that caption. You can see this best by clicking the add button
'   Go to item #3 (third from the bottom)in "testing the auto..." click it go to it again and see that
'   it is checked now click the add button again. go again to the #3 menu item you will now have two of them
'   doesn't matter which one you will see it's unchecked. click it. go to it again and it is now checked
'   go to the other #3 menu item and you will see it is also checked indicating there properties are read from
'   the same place. If you need two items with the same caption you can use the tag property to identify
'   the items in the core code but the code would need a lot of work.
'
'2) If the parent Item used when adding an item is not found the item want be created if there are two with the same
'   caption the items will be added to the first one found regardless of debth within the menu tree
'
'look in the control code for other issues and uses.

'add items to the main menu item

    Call MainMenu1.Add("Pictures", , "A new menu", False, False, False, "Testing", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Pictures", , "Another menu", False, False, False, "Testing", , , Me.Font, False, Picture1.Image)
'add items to an existing child menu
    Call MainMenu1.Add("Import Pictures From Net", , "Menu Options ", False, False, False, "Testing", , , Me.Font, False, Picture1.Image)
'add a seperator to the popup with a caption "Internet" note that the caption of a seperator is passed in the tooltip property
    Call MainMenu1.Add("Import Pictures From Net", 1, "-", , , , , , , , , , , , , , , , , "Internet")
    Call MainMenu1.Add("Import Pictures From Net", , "Testing the auto size feature", False, True, True, "Testing tag", , , Me.Font, False, Picture1.Image)
'add items to the item we just created
    Call MainMenu1.Add("Testing the auto size feature", , "Test1", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test2", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test3", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", 1, "-", , , , , , , , , , , , , , , , , "Seperator")
    Call MainMenu1.Add("Testing the auto size feature", , "Test4", True, True, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test5", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test6", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test7", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test8", True, True, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test9", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test10", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test11", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", 1, "-", , , , , , , , , , , , , , , , , "Devider")
    Call MainMenu1.Add("Testing the auto size feature", , "Test12", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test13", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test14", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test15", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test16", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test17", False, True, True, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test18", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test19", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Testing the auto size feature", , "Test20", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)


    Call MainMenu1.Add("Test17", , "Test117", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test127", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test137", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test147", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test1117", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test1127", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test1137", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test1147", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test2117", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test2127", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test2137", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test2147", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test3117", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test3127", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test3137", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test3147", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test4117", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test4127", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test4137", False, True, True, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test4147", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test5117", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test5127", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test5137", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test17", , "Test5147", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)

    Call MainMenu1.Add("Test4137", , "Test4137AA1", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA2", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA3", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA4", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA5", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA6", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA7", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA8", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA9", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA10", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA11", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA12", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA13", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA14", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA15", False, True, True, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA16", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA17", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA18", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA19", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA20", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137", , "Test4137AA21", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)

    Call MainMenu1.Add("Test4137AA15", , "AAAAAA1", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137AA15", , "AAAAAA2", False, True, True, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137AA15", , "AAAAAA3", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("Test4137AA15", , "AAAAAA4", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)

    Call MainMenu1.Add("AAAAAA2", , "AAAAAA21", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("AAAAAA2", , "AAAAAA22", False, True, True, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("AAAAAA2", , "AAAAAA23", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("AAAAAA2", , "AAAAAA24", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)

    Call MainMenu1.Add("AAAAAA22", , "AAAAAA221", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("AAAAAA22", , "AAAAAA222", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("AAAAAA22", , "AAAAAA223", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
    Call MainMenu1.Add("AAAAAA22", , "AAAAAA224", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image)
'add an item to the Main menubar just leave the parent pram blank
    Call MainMenu1.Add("", , "Remove Item", False, False, False, "Testing tag", , , Me.Font, False, Picture1.Image, True, True, "", 0, , , 0)
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0
    MainMenu1.Visible = Not MainMenu1.Visible
End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then MainMenu1.PopupMenu "Testing the auto size feature"
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then MainMenu2.PopupMenuGroup 1
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then MainMenu1.PopupMenu "Pictures"
End Sub

Private Sub MainMenu1_Click(Caption As String, Key As String)
    Debug.Print Caption 'respond to item being clicked
    Select Case LCase(Caption)
        Case "email" 'change the caption of items
            MainMenu1.Caption "Email", "View Email"
        Case "view email"
            MainMenu1.Caption "View Email", "Email"
        Case "remove item"
            MainMenu1.Remove "Test15"
            MainMenu1.Remove "Pictures"
            MainMenu1.isEnabled "Disabled Item", Not MainMenu1.isEnabled("Disabled Item"), False 'set the last param to false will set the value of that item
        Case "open"
            Form1.Show
        Case "test4"
            MainMenu1.isVisible "Pictures", Not MainMenu1.isVisible("Pictures"), False 'set the last param to false will set the value of that item
        Case "test3"
        'Example of how to check or uncheck an item
            MainMenu1.Check "Test3", Not MainMenu1.Check("Test3"), False 'set the last param to false will set the value of that item
        Case "test13"
        'Example of how to Hide or Unhide an item
            MainMenu1.isVisible "Test3", Not MainMenu1.isVisible("Test3"), False 'set the last param to false will set the value of that item
        'Example of how to Enable or Disable an item
            MainMenu1.isEnabled "Test4", Not MainMenu1.isEnabled("Test4"), False 'set the last param to false will set the value of that item
            
' in case the above is confusing the bellow lines do the same as the above line
            If MainMenu1.isEnabled("Test5") = False Then 'is the item enabled
                MainMenu1.isEnabled ("Test5"), True, False 'if not enable it
            Else
                MainMenu1.isEnabled ("Test5"), False, False 'if so disable it
            End If
    End Select
End Sub

Private Sub MainMenu2_Click(Caption As String, Key As String)
    Debug.Print Caption 'respond to item being clicked
        Select Case Caption
            Case "edit" 'react to a menuitem click before the popup is displayed
                'allowing the user to say, if no changes "undo" would be disabled. visible and checked work the same
                'toggle the enabled state
                MainMenu2.isEnabled "Copy", Not MainMenu2.isEnabled("Copy"), False
        End Select
End Sub

Private Sub MainMenu3_Click(Caption As String, Key As String)
    Debug.Print Caption 'respond to item being clicked
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'you can have popup menus based on group allowing you to have items in many locations
    'in the main menu even a main menu item.
    If Button = 2 Then
        Text1.Enabled = False
        Text1.Enabled = True
        MainMenu1.PopupMenuGroup 1
    Else
        'test the get/set style property
        MsgBox "The style Of Item 4 = " & MainMenu1.Style(3)
    End If
End Sub
