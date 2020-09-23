Attribute VB_Name = "TransBmp"
Option Explicit
'this is all me, it is set up to work with the drawing of these icons so wouldn't work well as a snippet
'without some rework. It's also slow if you process a larg picture. well it's slow anyway; but, it's
'tollerable for 16 * 16 pictures
Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long


Public Sub DrawImageTrans(DestPic As PictureBox, SourcePic As PictureBox)
'this works to draw the icon on the menu item but is very slow.
'ACTION: removes any pixel that matches the first pixel on the first line(.point(0,0))
'when I say removes what really happens is it's skiped in the drawing process
Dim startx As Long, sx As Long: startx = 3 * Screen.TwipsPerPixelX '45 '45 sets us out 3 pixels from the left
Dim starty As Long, sy As Long: starty = 2 * Screen.TwipsPerPixelY '30 '30 is 2 pixels from the top
    For sx = 0 To 256 Step Screen.TwipsPerPixelX 'the scalemode is twips and the icon is in pixels so they should be matched to save time in the loop
    'meaning for my system there are 15 twips in a pixel setting any of those twips color in a pixel base will in effect collor
    'all the twips(15*15 in my case) in that pixel to that color. that is why the bellow loops step forward by the
    'number of twips per pixel
        For sy = 0 To 256 Step Screen.TwipsPerPixelY 'using screen twipsperpixel for differant res.
            If SourcePic.Point(sx, sy) <> SourcePic.Point(0, 0) Then
                If SourcePic.Point(sx, sy) <> -1 Then
                    'DestPic.PSet (sx + startx, sy + starty), SourcePic.Point(sx, sy)
                    'faster than PSet
                    SetPixelV DestPic.hdc, (sx + startx) / Screen.TwipsPerPixelX, (sy + starty) / Screen.TwipsPerPixelY, SourcePic.Point(sx, sy)
                End If
            End If
        Next sy
    Next sx
End Sub


