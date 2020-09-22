Attribute VB_Name = "modGDI"
Option Explicit

' API Declares:
    
' The traditional rectangle structure:
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RGB
    r As Integer
    G As Integer             '  //--Selects a red, green, blue (RGB) color based on the arguments supplied
    B As Integer
End Type

'Public Const GRADIENT_FILL_TRIANGLE As Long = &H2
Public Const GRADIENT_FILL_RECT_H As Long = &H0
Public Const GRADIENT_FILL_RECT_V As Long = &H1

Public Enum GradientType
   gtHorz = GRADIENT_FILL_RECT_H
   gtVert = GRADIENT_FILL_RECT_V
End Enum

Public Const STRETCHMODE = vbPaletteModeContainer

Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Type PAINTSTRUCT
   hDC                     As Long
   fErase                  As Long
   rcPaint                 As RECT
   fRestore                As Long
   fIncUpdate              As Long
   rgbReserved(1 To 32)    As Byte
End Type

Public Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long


Public Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type



Public Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type


Public Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long


' This is most useful but Win32 only.  Particularly try the
' LOADMAP3DCOLORS for a quick way to sort out those
' embarassing gray backgrounds in your fixed bitmaps!
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" ( _
    ByVal hInst As Long, _
    ByVal lpsz As String, _
    ByVal un1 As Long, _
    ByVal n1 As Long, ByVal n2 As Long, _
    ByVal un2 As Long _
    ) As Long

Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2
Public Const LR_COLOR = &H2
Public Const LR_COPYDELETEORG = &H8
Public Const LR_COPYFROMRESOURCE = &H4000
Public Const LR_COPYRETURNORG = &H4
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_MONOCHROME = &H1
Public Const LR_SHARED = &H8000

' Creates a memory DC
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

' Creates a bitmap in memory:
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

' Places a GDI object into DC, returning the previous one:
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

' Deletes a GDI object:
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' Copies Bitmaps from one DC to another, can also perform
' raster operations during the transfer:
Public Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long _
    ) As Long

Public Declare Function PatBlt Lib "gdi32" ( _
   ByVal hDC As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal dwRop As Long) As Long

Public Declare Function TransparentBlt Lib "msimg32.dll" ( _
   ByVal hdcDest As Long, _
   ByVal nXOriginDest As Long, _
   ByVal nYOriginDest As Long, _
   ByVal nWidthDest As Long, _
   ByVal nHeightDest As Long, _
   ByVal hdcSrc As Long, _
   ByVal nXOriginSrc As Long, _
   ByVal nYOriginSrc As Long, _
   ByVal nWidthSrc As Long, _
   ByVal nHeightSrc As Long, _
   ByVal crTransparent As Long) As Long

Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCINVERT = &H660046

' Structure used to hold bitmap information about Bitmaps
' created using GDI in memory:
Public Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

' Get information relating to a GDI Object
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    lpObject As Any _
    ) As Long

' Fills a rectangle in a DC with a specified brush
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

' Create a brush of a certain colour:
'Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

'Pass in 0 as HDC to use Desktop HDC
Public Function BitmapFromDesktop(ByVal lWidth As Long, ByVal lHeight As Long) As cBitMap
  Dim lHwnd As Long
  Dim lCDC As Long
  
  lHwnd = GetDesktopWindow()
  lCDC = GetDC(lHwnd)
        
  Set BitmapFromDesktop = CreateBitmapFromHDC(lCDC, lWidth, lHeight)
  
  'Need to release DC if it was the desktop
   ReleaseDC lHwnd, lCDC
   
End Function

Public Function BitmapFromHDC(ByVal hDC As Long, ByVal lWidth As Long, ByVal lHeight As Long) As cBitMap
                
   Set BitmapFromHDC = CreateBitmapFromHDC(hDC, lWidth, lHeight)

End Function

Private Function CreateBitmapFromHDC(ByVal hDC As Long, ByVal lWidth As Long, ByVal lHeight As Long) As cBitMap
   Dim hBmp As Long
   Dim hBmpOld As Long
   Dim lHDC As Long
   
   lHDC = CreateCompatibleDC(hDC)
   
   Set CreateBitmapFromHDC = New cBitMap

   If (lHDC <> 0) Then
      ' If we get one, then time to make the bitmap:
      hBmp = CreateCompatibleBitmap(hDC, lWidth, lHeight)
      ' If we succeed in creating the bitmap:
      If (hBmp <> 0) Then
         hBmpOld = SelectObject(lHDC, hBmp)
         ' Success:
         CreateBitmapFromHDC.Initialize lHDC, hBmp, hBmpOld, lWidth, lHeight
      End If
   End If

End Function

Public Function BitmapFromFile(ByVal sBitmapPath As String) As cBitMap

   Dim tBM As Bitmap
   Dim hInst As Long
   Dim hBmp As Long
   Dim hBmpOld As Long
   Dim lWidth As Long, lHeight As Long
   Dim hDC As Long, hDCBasis As Long
   Dim lHwnd As Long
   
   Set BitmapFromFile = New cBitMap
      
   hInst = App.hInstance
    
   hBmp = LoadImage(hInst, sBitmapPath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    
   If (hBmp <> 0) Then
      lHwnd = GetDesktopWindow()
      hDCBasis = GetDC(lHwnd)
      hDC = CreateCompatibleDC(hDCBasis)
      If (hDC <> 0) Then
         ' If DC Is created, select the bitmap into it:
         hBmpOld = SelectObject(hDC, hBmp)
         'Get info about bitmap
         GetObjectAPI hBmp, Len(tBM), tBM
         lWidth = tBM.bmWidth
         lHeight = tBM.bmHeight
         BitmapFromFile.Initialize hDC, hBmp, hBmpOld, lWidth, lHeight
      End If
      ReleaseDC lHwnd, hDCBasis
   End If
      
End Function

Public Function BitmapFromPicture(ByRef picThis As StdPicture) As cBitMap
   Dim tBM As Bitmap
   Dim hInst As Long
   Dim hBmp As Long
   Dim hBmpOld As Long
   Dim lWidth As Long, lHeight As Long
   Dim hDC As Long, hDCBasis As Long
   Dim hDCTemp As Long
   Dim hBmpTemp As Long
   Dim hBmpTempOld As Long
   Dim lHwnd As Long
   
   Set BitmapFromPicture = New cBitMap
   
    ' Initialise byref variables:
    hDC = 0: hBmp = 0: hBmpOld = 0
        
    ' Create a DC to hold the sprite, and select
    ' the sprite into it:
    lHwnd = GetDesktopWindow()
    hDCBasis = GetDC(lHwnd)
    
    hDCTemp = CreateCompatibleDC(hDCBasis)
        
    If (hDCTemp <> 0) Then
       hBmpTempOld = SelectObject(hDCTemp, picThis.Handle)
    
        hDC = CreateCompatibleDC(hDCBasis)
        If (hDC <> 0) Then
            ' If we get one, then time to make the bitmap:
            GetObjectAPI picThis.Handle, Len(tBM), tBM
            
            hBmp = CreateCompatibleBitmap(hDCBasis, tBM.bmWidth, tBM.bmHeight)
            
            If (hBmp <> 0) Then
                hBmpOld = SelectObject(hDC, hBmp)
                
                BitBlt hDC, 0, 0, tBM.bmWidth, tBM.bmHeight, hDCTemp, 0, 0, SRCCOPY
                
                BitmapFromPicture.Initialize hDC, hBmp, hBmpOld, tBM.bmWidth, tBM.bmHeight
            End If
        End If
        
        SelectObject hDCTemp, hBmpTempOld
        DeleteObject hDCTemp
        
    End If
    
    ReleaseDC lHwnd, hDCBasis

End Function

Public Sub GDIClearDCBitmap( _
        ByRef hDC As Long, _
        ByRef hBmp As Long, _
        ByVal hBmpOld As Long _
    )
' **********************************************************
' GDI Helper function: Goes through the steps required
' to clear up a bitmap within a DC.
' **********************************************************
    ' If we have a valid DC:
    If (hDC <> 0) Then
        ' If there is a valid bitmap in it:
        If (hBmp <> 0) Then
            ' Select the original bitmap into the DC:
            SelectObject hDC, hBmpOld
            ' Now delete the unreferenced bitmap:
            DeleteObject hBmp
            ' Byref so set the value to invalid BMP:
            hBmp = 0
        End If
        ' Delete the memory DC:
        DeleteObject hDC
        ' Byref so set the value to invalid DC:
        hDC = 0
    End If
End Sub

Public Sub DrawGradientcBitmap(ByVal cBmp As cBitMap, ByVal Color1 As OLE_COLOR, ByVal Color2 As OLE_COLOR, Optional ByVal Direction As GradientType = gtHorz)
   DrawGradientHDC cBmp.hDC, 0, 0, cBmp.Width, cBmp.Height, Color1, Color2, Direction
End Sub

Public Sub DrawGradientHDC(cHdc As Long, x As Long, y As Long, x2 As Long, y2 As Long, ByVal Color1 As OLE_COLOR, ByVal Color2 As OLE_COLOR, Optional ByVal Direction As GradientType = gtHorz)

   Dim Vert(1) As TRIVERTEX   '2 Colors
   Dim gRect As GRADIENT_RECT
   
   Dim clr1RGB As RGB
   Dim clr2RGB As RGB
   clr1RGB = GetRGBColors(GetLngColor(Color1))
   clr2RGB = GetRGBColors(GetLngColor(Color2))
   
   With Vert(0)
      .x = x
      .y = y
      .Red = clr1RGB.r
      .Green = clr1RGB.G
      .Blue = clr1RGB.B
      .Alpha = 0&
   End With

   With Vert(1)
      .x = Vert(0).x + x2
      .y = Vert(0).y + y2
      .Red = clr2RGB.r
      .Green = clr2RGB.G
      .Blue = clr2RGB.B
      .Alpha = 0&
   End With

   gRect.UpperLeft = 0
   gRect.LowerRight = 1

   GradientFillRect cHdc, Vert(0), 2, gRect, 1, Direction

End Sub

Private Function GetRGBColors(Color As Long) As RGB

Dim HexColor As String
        
    HexColor = String(6 - Len(Hex(Color)), "0") & Hex(Color)
    GetRGBColors.r = "&H" & Mid(HexColor, 5, 2) & "00"
    GetRGBColors.G = "&H" & Mid(HexColor, 3, 2) & "00"
    GetRGBColors.B = "&H" & Mid(HexColor, 1, 2) & "00"
End Function

Private Function GetLngColor(Color As Long) As Long
    
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
    
End Function
