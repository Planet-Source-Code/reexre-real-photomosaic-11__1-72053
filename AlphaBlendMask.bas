Attribute VB_Name = "ABM"
'Thanx to "Tecc" of PSC

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
Private Type BITMAP '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private W As Long
Private H As Long

Private msk As Long, MSKO1 As Long, MSKI As BITMAPINFO, MSKBITS() As Byte
Private nSRC As Long, nSRCO1 As Long, nSRCI As BITMAPINFO, SRCBITS() As Byte
Private DST As Long, DSTO1 As Long, DSTI As BITMAPINFO, DSTBITS() As Byte
Private BB As Long, BBO As Long

Private LX As Long, LY As Long

Public Sub ModMask_Setup(ByRef PicSRC As PictureBox, ByRef PicTar As PictureBox, ByRef Target As PictureBox)
Dim S1 As Long
Dim S2 As Long
S1 = PicSRC.hDC
S2 = PicTar.hDC

'set the width and height
W = PicSRC.ScaleWidth
H = PicSRC.ScaleHeight
'set bitmap info for the source, mask, and destination
'bitmaps
With MSKI.bmiHeader
    .biBitCount = 24 '24 bits per pixel (R,G,B per pixel)
    .biSize = Len(MSKI) 'size of this information
    .biHeight = H 'height
    .biWidth = W 'width
    .biPlanes = 1 'bitmap planes (2D, so 1)
    .biCompression = BI_RGB 'Type of color compression
End With
'the following is the same for all bitmaps
With DSTI.bmiHeader
    .biBitCount = 24
    .biSize = Len(DSTI)
    .biHeight = H
    .biWidth = W
    .biPlanes = 1
    .biCompression = BI_RGB
End With
With nSRCI.bmiHeader
    .biBitCount = 24
    .biSize = Len(nSRCI)
    .biHeight = H
    .biPlanes = 1
    .biWidth = W
    .biCompression = BI_RGB
End With

'create the device contexts
msk = CreateCompatibleDC(GetDC(0))
nSRC = CreateCompatibleDC(GetDC(0))
DST = CreateCompatibleDC(GetDC(0))
BB = CreateCompatibleDC(GetDC(0))

'variable that defines how many color bits there
'are in one bit array
'[Width * Height] (all pixels) [* 3] (R,G,B - 3 values)
'per pixel
Dim nl As Long
nl = ((W + 1) * (H + 1)) * 3

'redimension the bit color information arrays to
'fit all the color information
ReDim MSKBITS(1 To nl)
ReDim SRCBITS(1 To nl)
ReDim DSTBITS(1 To nl)

'create a DIB section based on the bitmapinfo we
'provided above. this is like creating a
'compatible bitmap, but used for modifying bitmap
'bits
MSKO1 = CreateDIBSection(GetDC(0), MSKI, DIB_RGB_COLORS, 0, 0, 0)
nSRCO1 = CreateDIBSection(GetDC(0), nSRCI, DIB_RGB_COLORS, 0, 0, 0)
DSTO1 = CreateDIBSection(GetDC(0), DSTI, DIB_RGB_COLORS, 0, 0, 0)

'create a permanent image of the form, so we can
'restore drawn-over parts
BBO = CreateCompatibleBitmap(GetDC(0), Target.ScaleWidth, Target.ScaleHeight)

'link the device contexts to thier bitmap objects
SelectObject msk, MSKO1
SelectObject DST, DSTO1
SelectObject nSRC, nSRCO1
SelectObject BB, BBO

'we want to blt from the form, so make sure it
'is visible
'target.Show
Target.Refresh

'blt the mask and source images into the bitmap
'object so we can copy the color information
BitBlt msk, 0, 0, W, H, S2, 0, 0, vbSrcCopy
BitBlt nSRC, 0, 0, W, H, S1, 0, 0, vbSrcCopy
BitBlt BB, 0, 0, Target.ScaleWidth, Target.ScaleHeight, Target.hDC, 0, 0, vbSrcCopy

'load the color information into arrays
'we only do this once because the source and mask
'images never change, but the destination image
'will change frequently, depending on where the mouse
'is on the form, so we have to update the DST bit array
'every time we alphablt.
GetDIBits msk, MSKO1, 0, H, MSKBITS(1), MSKI, DIB_RGB_COLORS
GetDIBits nSRC, nSRCO1, 0, H, SRCBITS(1), nSRCI, DIB_RGB_COLORS

End Sub

Public Sub ModMask_CleanUp()
'cleanup all the memory space we have used
DeleteDC msk
DeleteDC nSRC
DeleteDC DST
DeleteDC BB

DeleteObject BBO
DeleteObject MSKO1
DeleteObject nSRCO1
DeleteObject DSTO1

'erase any array data left over
Erase MSKBITS
Erase SRCBITS
Erase DSTBITS
End Sub

Public Sub ModMask_BLTIT(ByVal x As Long, ByVal y As Long, Target As PictureBox)
'set the cursor in the middle of the alpha-blitted
'bitmap
x = x - W / 2
y = y - H / 2

''if the area is off the form, move it back

'This is like a bug in borders
'From V7.6 Removed!

'If x >= Target.ScaleWidth - (W + 1) Then
'    x = Target.ScaleWidth - (W + 1)
'End If
'If y >= Target.ScaleHeight - (H + 1) Then
'    y = Target.ScaleHeight - (H + 1)
'End If
'If y <= 0 Then
'    y = 0
'End If
'If x <= 0 Then
'    x = 0
'End If



'copy image from the permanant image of the form
'to the destination bitmap, so we have a 'background'
'to alphablt to. This is so that we will blt to the
'area where the cursor is
BitBlt DST, 0, 0, W, H, BB, x, y, vbSrcCopy

'copy the destination image data into its bit array
'so we can process it
GetDIBits DST, DSTO1, 0, H, DSTBITS(1), DSTI, DIB_RGB_COLORS

'some processing variables
Dim SrcC(2) As Integer
Dim DstC(2) As Integer
Dim Alpha(2) As Integer
Dim tmp(2) As Integer

'temporary bit array
Dim tmpBits() As Byte

'make the temporary bit array large enough to hold
'all the color information from the resulting alpha
'blitted bitmap
ReDim tmpBits(UBound(SRCBITS))

'a for loop to loop through the pixels of the bitmaps
'we do step3 because for every pixel, there are RED,
'GREEN, and BLUE color values in the bit array
For i = 1 To UBound(SRCBITS) Step 3
    'pixel: (i) to (i+2)
    SrcC(0) = SRCBITS(i) 'blue value
    SrcC(1) = SRCBITS(i + 1) 'green value
    SrcC(2) = SRCBITS(i + 2) 'red value
    
    Alpha(0) = MSKBITS(i)
    Alpha(1) = MSKBITS(i + 1)
    Alpha(2) = MSKBITS(i + 2)
    
    DstC(0) = DSTBITS(i)
    DstC(1) = DSTBITS(i + 1)
    DstC(2) = DSTBITS(i + 2)
    
    'create alpha values based on color information
    'the transparency level is based on the current mask
    'pixel, the function DOES calculate green, red and blue
    'alpha channels, but in this example, only GRAYSCALE
    'is used (because the mask BMP is black and white).
    'but there is a true 32bit alpha channel available with
    'no decrease in speed.
    
    'when we use a GRAYSCALE mask, all alpha
    'values are the same. with a full color mask,
    'alpha values differ and so depending on a certain pixel,
    'more green, blue, or red can be forced transparent
    
    'say you have a source pixel, RGB(100,0,0)
    'a mask pixel, RGB(200,0,0)
    'and a destination pixel RGB(0,100,255)
    'the alpha pixel would be RGB(78,0,0) showing
    'only a red pixel, because in the mask, only
    'red has a visible value (200). Some very
    'interesting effects are available for you to
    'experiment with.
    tmp(0) = SrcC(0) + (((DstC(0) - SrcC(0)) / 255) * Alpha(0))
    tmp(1) = SrcC(1) + (((DstC(1) - SrcC(1)) / 255) * Alpha(1))
    tmp(2) = SrcC(2) + (((DstC(2) - SrcC(2)) / 255) * Alpha(2))
    
    'set the alpha values into the temporary bit
    'array
    tmpBits(i) = tmp(0) 'Alpha Blue
    tmpBits(i + 1) = tmp(1) 'Alpha Green
    tmpBits(i + 2) = tmp(2) 'Alpha Red
Next

'copy the previous image over where we alphablitted
'last, so we clear only that part of the screen.
BitBlt Target.hDC, LX, LY, W, H, BB, LX, LY, vbSrcCopy

'blt the alpha values to the screen
SetDIBitsToDevice Target.hDC, x, y, W, H, 0, 0, 0, H, tmpBits(1), nSRCI, DIB_RGB_COLORS

'set the Last X and Last Y values so we know where
'tp clear the screen next time.
LX = x
LY = y
End Sub
