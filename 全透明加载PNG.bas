Attribute VB_Name = "PNG"
'功能：全透明加载PNG
'调用：  MakeTrans 窗体, PNGPath

Option Explicit

Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hWnd As Long, graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetDC Lib "gdiplus" (ByVal graphics As Long, hdc As Long) As GpStatus
Public Declare Function GdipReleaseDC Lib "gdiplus" (ByVal graphics As Long, ByVal hdc As Long) As GpStatus
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As String, image As Long) As GpStatus
Public Declare Function GdipCloneImage Lib "gdiplus" (ByVal image As Long, cloneImage As Long) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, BITMAP As Long) As GpStatus
Public Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal BITMAP As Long, ByVal X As Long, ByVal Y As Long, color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal BITMAP As Long, ByVal X As Long, ByVal Y As Long, ByVal color As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, BITMAP As Long) As GpStatus

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)


Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_SYSCOMMAND = &H112

Private Const ULW_OPAQUE = &H4
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOOWNERZORDER As Long = &H200

Public Type GdiplusStartupInput
   GdiplusVersion As Long
   DebugEventCallback As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs As Long
End Type

Public Enum GpStatus
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum


Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
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

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long

Dim mDC As Long
Dim mainBitmap As Long
Dim blendFunc32bpp As BLENDFUNCTION
Dim token As Long
Dim oldBitmap As Long

Private Function MakeTrans2(Yourform As Object, pngPath As String) As Boolean

   Dim tempBI As BITMAPINFO
   Dim tempBlend As BLENDFUNCTION
   Dim lngHeight As Long, lngWidth As Long
   Dim curWinLong As Long
   Dim img As Long
   Dim graphics As Long
   Dim winSize As Size
   Dim srcPoint As POINTAPI
   
   With tempBI.bmiHeader
      .biSize = Len(tempBI.bmiHeader)
      .biBitCount = 32
      .biHeight = Yourform.ScaleHeight
      .biWidth = Yourform.ScaleWidth
      .biPlanes = 1
      .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8)
   End With
   mDC = CreateCompatibleDC(Yourform.hdc)
   mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
   oldBitmap = SelectObject(mDC, mainBitmap)
    
   Call GdipCreateFromHDC(mDC, graphics)
   Call GdipLoadImageFromFile(StrConv(pngPath, vbUnicode), img)
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   Call GdipDrawImageRect(graphics, img, 0, 0, lngWidth, lngHeight)

   curWinLong = GetWindowLong(Yourform.hWnd, GWL_EXSTYLE)
   
   SetWindowLong Yourform.hWnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED
   SetWindowPos Yourform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER
   
   srcPoint.X = 0
   srcPoint.Y = 0
   winSize.cx = Yourform.ScaleWidth
   winSize.cy = Yourform.ScaleHeight
    
   With blendFunc32bpp
      .AlphaFormat = AC_SRC_ALPHA
      .BlendFlags = 0
      .BlendOp = AC_SRC_OVER
      .SourceConstantAlpha = 255
   End With
    
   Call GdipDisposeImage(img)
   Call GdipDeleteGraphics(graphics)
   Call UpdateLayeredWindow(Yourform.hWnd, Yourform.hdc, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
   
End Function

Public Function MakeTrans(Yourform As Object, pngPath As String) As Boolean
    Dim GpInput As GdiplusStartupInput
    GpInput.GdiplusVersion = 1
    If GdiplusStartup(token, GpInput) <> 0 Then MakeTrans = False
    Yourform.ScaleMode = 3
    MakeTrans2 Yourform, pngPath
    
    Call GdiplusShutdown(token)
    SelectObject mDC, oldBitmap
    DeleteObject mainBitmap
    DeleteObject oldBitmap
    DeleteDC mDC
End Function

