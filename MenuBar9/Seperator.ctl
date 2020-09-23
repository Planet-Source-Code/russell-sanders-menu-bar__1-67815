VERSION 5.00
Begin VB.UserControl Seperator 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BackStyle       =   0  'Transparent
   ClientHeight    =   105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PaletteMode     =   4  'None
   ScaleHeight     =   105
   ScaleWidth      =   615
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   525
      TabIndex        =   0
      Top             =   0
      Width           =   525
   End
End
Attribute VB_Name = "Seperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private mOrentation As Long 'not implimented in this version
Private mCaption As String
Private mGradiant As OLE_COLOR
Private mForeColor As OLE_COLOR

'Public Enum Orentation 'set the orentation of the spacer bar
'    Vertical
'    Horazonal
'End Enum

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(NewValue As OLE_COLOR)
    mForeColor = NewValue
    drawMe
End Property

Public Property Get Gradiant() As OLE_COLOR
    Gradiant = mGradiant
End Property

Public Property Let Gradiant(NewValue As OLE_COLOR)
    mGradiant = NewValue
    drawMe
End Property

Private Sub drawMe()
   Call GetTextExtentPoint32(Picture1.hdc, mCaption, Len(mCaption), thisPoint)
    If mCaption = "" Then thisPoint.Y = 6
    If UserControl.Height <> thisPoint.Y * Screen.TwipsPerPixelY Then
        UserControl.Height = thisPoint.Y * Screen.TwipsPerPixelY
        Exit Sub
    End If
    Dim s1 As Long
    Dim s2 As Long
    Picture1.Cls
    If Not Gradiant = 0 Then DrawGradientFill BackColor, Gradiant, Picture1
    s1 = ((Picture1.Width - (thisPoint.X * Screen.TwipsPerPixelX)) / 2) - 60
    s2 = (((Picture1.Width \ 2) - ((thisPoint.X * Screen.TwipsPerPixelX) \ 2)) + (thisPoint.X * Screen.TwipsPerPixelX)) + 60
    Picture1.Line (s1, (Picture1.Height \ 2) - 15)-(0, (Picture1.Height \ 2) - 15), &H8000000D
    Picture1.Line (s1, (Picture1.Height \ 2))-(0, (Picture1.Height \ 2)), &H80000005
    Picture1.Line (Picture1.Width, (Picture1.Height \ 2) - 15)-(s2, (Picture1.Height \ 2) - 15), &H8000000D
    Picture1.Line (Picture1.Width, (Picture1.Height \ 2))-(s2, (Picture1.Height \ 2)), &H80000005
    Picture1.CurrentX = (Picture1.Width \ 2) - ((thisPoint.X * Screen.TwipsPerPixelX) \ 2)
    Picture1.CurrentY = -15
    Picture1.ForeColor = BackColor
    Picture1.Print mCaption
    Picture1.CurrentX = ((Picture1.Width \ 2) - ((thisPoint.X * Screen.TwipsPerPixelX) \ 2)) + 15
    Picture1.CurrentY = -15
    Picture1.ForeColor = mForeColor
    Picture1.Print mCaption
End Sub
'Public Property Get Orentation() As Orentation
'    Orentation = mOrentation
'End Property
'
'Public Property Let Orentation(NewValue As Orentation)
'    mOrentation = NewValue
'    UserControl_Resize
'End Property

Private Sub UserControl_Resize()
    Picture1.Width = UserControl.Width
    Picture1.Height = UserControl.Height
    drawMe
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Picture1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Picture1.BackColor() = New_BackColor
    drawMe
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    Caption = mCaption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    mCaption = New_Caption
    drawMe
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Picture1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Picture1.Font = New_Font
    drawMe
End Property

