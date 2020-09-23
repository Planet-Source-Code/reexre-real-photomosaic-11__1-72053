VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type Bitmap
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long

Private bmpBits() As Byte
Private hBmp As Bitmap

Public Sub GetBits(pBoxPicturehandle As Long)
Dim iRet As Long
    'Get the bitmap header
    iRet = GetObject(pBoxPicturehandle, Len(hBmp), hBmp)
    'Resize to hold image data
    ReDim bmpBits(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into bmpBits array
    iRet = GetBitmapBits(pBoxPicturehandle, hBmp.bmWidthBytes * hBmp.bmHeight, bmpBits(0, 0, 0))
End Sub

Public Sub SetBits(pBoxPicturehandle As Long)
Dim iRet As Long
    'Set the new image data back onto pBox
    iRet = SetBitmapBits(pBoxPicturehandle, hBmp.bmWidthBytes * hBmp.bmHeight, bmpBits(0, 0, 0))
    'Erase bmpBits because we finished with it now
    Erase bmpBits
End Sub



Public Sub MYADD(R, G, B)
Dim x As Long
Dim y As Long
Dim TR As Integer
Dim TG As Integer
Dim TB As Integer

    'Call GetBits(pDst)
    
    For x = 0 To hBmp.bmWidth - 1
        For y = 0 To hBmp.bmHeight - 1
   
            TR = bmpBits(2, x, y) + R
            TG = bmpBits(1, x, y) + G
            TB = bmpBits(0, x, y) + B
          
            If TR < 0 Then TR = 0
            If TG < 0 Then TG = 0
            If TB < 0 Then TB = 0
            
            If TR > 255 Then TR = 255
            If TG > 255 Then TG = 255
            If TB > 255 Then TB = 255
            
            bmpBits(2, x, y) = TR
            bmpBits(1, x, y) = TG
            bmpBits(0, x, y) = TB
            
        Next y
    Next x
    '''Here we set the new bits
    '''Call SetBits(pDst)
    '''And Refresh the picturebox
    '''Call pDst.Refresh
End Sub
