Attribute VB_Name = "modGradient"
'I downloaded this module a few years back and not sure
'where but if it looks like something you wrote or someones you know
'please let me know so I can give proper credit
'
'I think its' from API viewer but not sure.
'
'I made a few changes to this allowing it to do center out grads.
'likely slowed it down some.
Option Explicit

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Private Const GRADIENT_FILL_RECT_H  As Long = &H0
Private Const GRADIENT_FILL_RECT_V  As Long = &H1
Private Const GRADIENT_FILL_TRIANGLE As Long = &H2

Public GRADIENT_FILL_RECT_DIRECTION As Long

Private Type TRIVERTEX
   X As Long
   Y As Long
   red As Integer
   green As Integer
   blue As Integer
   Alpha As Integer
End Type

Private Type GRADIENT_RECT
   UpperLeft As Long
   LowerRight As Long
End Type

Dim vert()  As TRIVERTEX
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Public Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As Any, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Public Sub DrawGradientFill(ByVal dwColour1 As Long, ByVal dwColour2 As Long, ByRef Picture1 As PictureBox)
   Dim grRc As GRADIENT_RECT
   Dim R1 As Integer, R2 As Integer
   Dim G1 As Integer, G2 As Integer
   Dim B1 As Integer, B2 As Integer
   R1 = LongToSignedShort((dwColour1 And &HFF&) * 256)
   G1 = LongToSignedShort(((dwColour1 And &HFF00&) \ &H100&) * 256)
   B1 = LongToSignedShort(((dwColour1 And &HFF0000) \ &H10000) * 256)
   R2 = LongToSignedShort((dwColour2 And &HFF&) * 256)
   G2 = LongToSignedShort(((dwColour2 And &HFF00&) \ &H100&) * 256)
   B2 = LongToSignedShort(((dwColour2 And &HFF0000) \ &H10000) * 256)
    ReDim vert(0 To 1) 'As TRIVERTEX
    If GRADIENT_FILL_RECT_DIRECTION = 0 Then
            With vert(0)  'Colour at upper-left corner
               .X = 0
               .Y = 0
               .red = R1
               .green = G1
               .blue = B1
               .Alpha = 0
            End With
            With vert(1)  'Colour at bottom-right corner
               .X = Picture1.ScaleWidth \ Screen.TwipsPerPixelX
               .Y = Picture1.ScaleHeight \ Screen.TwipsPerPixelY
               .red = R2
               .green = G2
               .blue = B2
               .Alpha = 0
            End With
            With grRc
               .LowerRight = 0
               .UpperLeft = 1
            End With
        Call GradientFill(Picture1.hdc, vert(0), 2, grRc, 1, Abs(1))
   ElseIf GRADIENT_FILL_RECT_DIRECTION = 1 Then
           With vert(0)  'Colour at upper-left corner
               .X = 0 '2
               .Y = 0
               .red = R1
               .green = G1
               .blue = B1
               .Alpha = 0
            End With
            With vert(1)  'Colour at bottom-right corner
               .X = (Picture1.ScaleWidth \ Screen.TwipsPerPixelX) '(Picture1.ScaleWidth \ Screen.TwipsPerPixelX) - 2
               .Y = (Picture1.ScaleHeight \ Screen.TwipsPerPixelY) / 2
               .red = R2
               .green = G2
               .blue = B2
               .Alpha = 0
            End With
            With grRc
               .LowerRight = 0
               .UpperLeft = 1
            End With
        Call GradientFill(Picture1.hdc, vert(0), 2, grRc, 1, Abs(1))
            With vert(0)  'Colour at upper-left corner
               .X = 0 '2
               .Y = (Picture1.ScaleHeight \ Screen.TwipsPerPixelY) / 2
               .red = R2
               .green = G2
               .blue = B2
               .Alpha = 0
            End With
            With vert(1)  'Colour at bottom-right corner
               .X = (Picture1.ScaleWidth \ Screen.TwipsPerPixelX) '(Picture1.ScaleWidth \ Screen.TwipsPerPixelX) - 2
               .Y = Picture1.ScaleHeight \ Screen.TwipsPerPixelY
               .red = R1
               .green = G1
               .blue = B1
               .Alpha = 0
            End With
            With grRc
               .LowerRight = 0
               .UpperLeft = 1
            End With
        Call GradientFill(Picture1.hdc, vert(0), 2, grRc, 1, Abs(1))
    ElseIf GRADIENT_FILL_RECT_DIRECTION = 2 Then
            With vert(0)  'Colour at upper-left corner
               .X = 0
               .Y = 0
               .red = R1
               .green = G1
               .blue = B1
               .Alpha = 0
            End With
            With vert(1)  'Colour at bottom-right corner
               .X = Picture1.ScaleWidth \ Screen.TwipsPerPixelX
               .Y = Picture1.ScaleHeight \ Screen.TwipsPerPixelY
               .red = R2
               .green = G2
               .blue = B2
               .Alpha = 0
            End With
            With grRc
               .LowerRight = 0
               .UpperLeft = 1
            End With
        Call GradientFill(Picture1.hdc, vert(0), 2, grRc, 1, Abs(0))
    ElseIf GRADIENT_FILL_RECT_DIRECTION = 3 Then
            With vert(0)  'Colour at upper-left corner
                .X = 0
                .Y = 0
                .red = R1
                .green = G1
                .blue = B1
                .Alpha = 0
             End With
             With vert(1)  'Colour at bottom-right corner
                .X = (Picture1.ScaleWidth \ Screen.TwipsPerPixelX) / 2
                .Y = Picture1.ScaleHeight \ Screen.TwipsPerPixelY
                .red = R2
                .green = G2
                .blue = B2
                .Alpha = 0
             End With
             With grRc
                .LowerRight = 0
                .UpperLeft = 1
             End With
        Call GradientFill(Picture1.hdc, vert(0), 2, grRc, 1, Abs(0))
            With vert(0)  'Colour at upper-left corner
               .X = (Picture1.ScaleWidth \ Screen.TwipsPerPixelX) / 2
               .Y = 0
               .red = R2
               .green = G2
               .blue = B2
               .Alpha = 0
            End With
            With vert(1)  'Colour at bottom-right corner
               .X = Picture1.ScaleWidth \ Screen.TwipsPerPixelX
               .Y = Picture1.ScaleHeight \ Screen.TwipsPerPixelY
               .red = R1
               .green = G1
               .blue = B1
               .Alpha = 0
            End With
            With grRc
               .LowerRight = 0
               .UpperLeft = 1
            End With
        Call GradientFill(Picture1.hdc, vert(0), 2, grRc, 1, Abs(0))
    End If
End Sub


Private Function LongToSignedShort(dwUnsigned As Long) As Integer  'convert from long to signed short
   If dwUnsigned < 32768 Then
      LongToSignedShort = CInt(dwUnsigned)
   Else
      LongToSignedShort = CInt(dwUnsigned - &H10000)
   End If
End Function


