Attribute VB_Name = "mdlGrayScale"
'-----------------------------------------------------------------------------------------------------------
' !! Paint GrayScale !!
'-----------------------------------------------------------------------------------------------------------
' SourceCode : mdlGrayScale
' Auther     : Jim Jose
' Email      : jimjosev33@yahoo.com
' Date       : 19-12-2005
' Purpose    : Hi-Speed grayscale function
'            : Support for Transparent icons
' CopyRight  : JimJose © Gtech Creations - 2005
' Thanks to  : Carls p.v (my master in DIB)
'-----------------------------------------------------------------------------------------------------------
' About
' -----
'    Hi guys, This is a rock solid hi-speed code(DIB) for image
' GrayScale. GrayScales are very important especially for usercontrols
' to draw disabled states!!
'
'    I can't find any stand-alone hi-speeed code
' for this. And none of them supports icons too...(icons are very
' important for uc). So here is it... you can pass Bitmaps as well as
' icons to this function. The transparency of icons will be retained
' with absolutly no memory leak!!!
'
'    Even though this routine is for GrayScale... it can do any
' kind of picture effect with simple changes in the pixel-editing loop.
' If anyone interested... tell me. I will post some nice ones!!(later):)
'
' Regards,
' Jim Jose
'-----------------------------------------------------------------------------------------------------------
'
' Usage
'        PaintGrayScale picDraw.hdc, picIcon.Picture, 0, 0, -1, -1
'
Option Explicit

'[Apis]
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

'[Types]
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Type BITMAP
   bmType       As Long
   bmWidth      As Long
   bmHeight     As Long
   bmWidthBytes As Long
   bmPlanes     As Integer
   bmBitsPixel  As Integer
   bmBits       As Long
End Type


'------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : PaintGrayScale
' Auther    : Jim Jose
' Input     : Hdc + Picture + Position
' OutPut    : None
' Purpose   : Hi-Speed grayscale... icons supported !!
'------------------------------------------------------------------------------------------------------------------------------------------

Public Function PaintGrayScale(ByVal lHDC As Long, _
                            ByVal hPicture As Long, _
                            ByVal lLeft As Long, _
                            ByVal lTop As Long, _
                            Optional ByVal lWidth As Long = -1, _
                            Optional ByVal lHeight As Long = -1) As Boolean

 Dim BMP        As BITMAP
 Dim BMPiH      As BITMAPINFOHEADER
 Dim lBits()    As Byte 'Packed DIB
 Dim lTrans()   As Byte 'Packed DIB
 Dim TmpDC      As Long
 Dim X          As Long
 Dim xMax       As Long
 Dim TmpCol     As Long
 Dim R1         As Long
 Dim G1         As Long
 Dim B1         As Long
 Dim bIsIcon    As Boolean
 
    'Get the Image format
    If (GetObjectType(hPicture) = 0) Then
        Dim mIcon As ICONINFO
        bIsIcon = True
        GetIconInfo hPicture, mIcon
        hPicture = mIcon.hbmColor
    End If

    'Get image info
    GetObject hPicture, Len(BMP), BMP

    'Prepare DIB header and redim. lBits() array
    With BMPiH
       .biSize = Len(BMPiH) '40
       .biPlanes = 1
       .biBitCount = 24
       .biWidth = BMP.bmWidth
       .biHeight = BMP.bmHeight
       .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
        If lWidth = -1 Then lWidth = .biWidth
        If lHeight = -1 Then lHeight = .biHeight
    End With
    ReDim lBits(Len(BMPiH) + BMPiH.biSizeImage)   '[Header + Bits]

    'Create TemDC and Get the image bits
    TmpDC = CreateCompatibleDC(lHDC)
    GetDIBits TmpDC, hPicture, 0, BMP.bmHeight, lBits(0), BMPiH, 0

    'Loop through the array... (grayscale - average!!)
    xMax = BMPiH.biSizeImage - 1
    For X = 0 To xMax - 3 Step 3
        R1 = lBits(X)
        G1 = lBits(X + 1)
        B1 = lBits(X + 2)
        TmpCol = (R1 + G1 + B1) / 3
'-------------------------------------------------------------------------------
'I Added this line to keep the gray from going to black for icons this works good
        If TmpCol < 100 Then TmpCol = 100 'don't allow black
'--------------------------------------------------------------------------------
        lBits(X) = TmpCol
        lBits(X + 1) = TmpCol
        lBits(X + 2) = TmpCol
    Next X
    ' Paint it!
    If bIsIcon Then
        ReDim lTrans(Len(BMPiH) + BMPiH.biSizeImage)
        GetDIBits TmpDC, mIcon.hbmMask, 0, BMP.bmHeight, lTrans(0), BMPiH, 0  ' Get the mask
        Call StretchDIBits(lHDC, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lTrans(0), BMPiH, 0, vbSrcAnd)   ' Draw the mask
        PaintGrayScale = StretchDIBits(lHDC, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lBits(0), BMPiH, 0, vbSrcPaint)  'Draw the gray
        DeleteObject mIcon.hbmMask  'Delete the extracted images
        DeleteObject mIcon.hbmColor
    Else
        PaintGrayScale = StretchDIBits(lHDC, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lBits(0), BMPiH, 0, vbSrcCopy)
    End If
    
    'Clear memory
    DeleteDC TmpDC
    
 Erase lBits
 Erase lTrans
End Function


