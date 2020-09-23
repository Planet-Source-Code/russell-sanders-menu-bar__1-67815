VERSION 5.00
Object = "*\AMenuBar.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin MenuBar.MainMenu MainMenu2 
      Height          =   405
      Left            =   1800
      TabIndex        =   1
      Top             =   1410
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   714
      BotBorder       =   0   'False
      RightBorder     =   0   'False
      LeftBorder      =   0   'False
      TopBorder       =   0   'False
   End
   Begin MenuBar.MainMenu MainMenu1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   714
      retainApp       =   -1  'True
      GradientDirection=   1
      BackColor       =   255
      TotalItems      =   7
      Children        =   "&Address Book\&File\Edit\View"
      Caption0        =   "&Address Book"
      bkCol0          =   255
      frCol0          =   16777215
      Icon0           =   "Form1.frx":0000
      sCut0           =   "F1"
      Gradiant0       =   12615935
      Caption1        =   "&File"
      bkCol1          =   13619151
      Icon1           =   "Form1.frx":0352
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      isPop1          =   -1  'True
      Caption2        =   "....&Open"
      bkCol2          =   13619151
      Check2          =   -1  'True
      Drop2           =   -1  'True
      Icon2           =   "Form1.frx":06A4
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      sCut2           =   "F2"
      Caption3        =   "....New"
      bkCol3          =   13619151
      Icon3           =   "Form1.frx":09F6
      BeginProperty Font3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption4        =   "Edit"
      Icon4           =   "Form1.frx":0D48
      BeginProperty Font4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      isPop4          =   -1  'True
      sCut4           =   "F3"
      BStyle4         =   0
      Caption5        =   "....&Cut"
      Icon5           =   "Form1.frx":109A
      BeginProperty Font5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      sCut5           =   "Ctrl + X"
      Caption6        =   "....Copy"
      Icon6           =   "Form1.frx":13EC
      BeginProperty Font6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      sCut6           =   "Ctrl + C"
      Caption7        =   "View"
      Icon7           =   "Form1.frx":173E
      BeginProperty Font7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BStyle7         =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MainMenu1.ProcessKeys MainMenu1.curItemM, KeyCode, Shift
End Sub

Private Sub Form_Load()
'Dim Picture1 As PictureBox
'    Set Picture1 = Me.Controls.Add("VBPictureBox", "PictureBox")
End Sub

Private Sub MainMenu1_Click(Caption As String, Key As String)
    Debug.Print Caption
End Sub

