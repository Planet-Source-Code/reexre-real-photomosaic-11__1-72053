Attribute VB_Name = "MyTypes"
Type tCollTile
    
    filename As String
    FileDate As Double
    
    oW As Integer
    oH As Integer
    
    WW As Integer
    HH As Integer
    
    R() As Byte
    G() As Byte
    B() As Byte
    
    Mirrored As Boolean
    
End Type



Type tPMzone
    CX As Single
    CY As Single
    
    OnPmWidht As Single
    OnPmHeight As Single
    
    WZone_W As Integer
    WZone_H As Integer
    
    
    ANG As Single
    R(18, 18) As Byte
    G(18, 18) As Byte
    B(18, 18) As Byte
    
    FIT() As Single
    FitINDEXusable() As Boolean
    
    agR() As Integer
    agB() As Integer
    agG() As Integer
    
    indexBESTFIT As Single
    
    indexBFfileName As String
    
    DrawOrder As Long
    
    Mirrored As Boolean
    
    
    
    
End Type




Type tColl
    NofPhotos As Integer
    Z() As tCollTile
    
    NAME As String
    
    STARTdir As String
    
End Type

Type tFM
    TotalW As Integer
    TotalH As Integer
    
    GlobTileW As Single
    GlobTileH As Single
    
    WZone_W As Integer
    WZone_H As Integer
    
    FilePathIN As String
    FilePathOUT As String
    
    NZones As Integer
    
    toSEE() As tPMzone
    
    MaskName As String
    
    'MaskPercX As Single
    'MaskPercY As Single
    
End Type


Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long

Public Const STRETCHMODE = vbPaletteModeNone 'You can find other modes in the "PaletteModeConstants" section of your Object Browser

Public H(16) As String

Public GlobalMirrored() As Boolean


'
'Private fastSQR(0 To 1530) As Single
Public fastSQR(0 To 195075) As Single



Public Function PixelsToCentimeter(pxs, Optional DPI = 300) As Single
'dpi = pixel (dot) x pollice

'1 pollice= 2.54 cm
'1 cm = 0.39 pollici

Const CMxInch = 2.54
Const InchxCM = 0.39
Dim Inch As Single

Inch = pxs / DPI

PixelsToCentimeter = Inch * CMxInch
PixelsToCentimeter = Int(PixelsToCentimeter * 10) / 10

End Function

Public Function fnRND(Min, Max, DoRound As Boolean)

fnRND = Rnd * (Max - Min) + Min
If DoRound Then fnRND = Round(fnRND)
'Debug.Print "fnRND___ ", min, max, fnRND

End Function
