VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBitMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' Class:    cBitmap
' Filename: cBitmap.cls
' Author:   SP McMahon
' Date:     14 July 1998
'
' A class to store a bitmap.  Similar to cSpriteBitmap but doesn't
' use a mask.
' ======================================================================

' Memory DC & Bitmap to hold the bitmap
Private m_hdcBitmap As Long
Private m_hBmpBitmap As Long
Private m_hBmpBitmapOld As Long
' Width & Height of the bitmap
Private m_lDX As Long
Private m_lDY As Long

Friend Sub Initialize(ByVal lhdc As Long, ByVal lBitmap As Long, ByVal lOldBitmap As Long, ByVal lwidth As Long, ByVal lheight As Long)
   m_hdcBitmap = lhdc
   m_hBmpBitmap = lBitmap
   m_hBmpBitmapOld = lOldBitmap
   m_lDX = lwidth
   m_lDY = lheight
End Sub

Property Get Width() As Long
    Width = m_lDX
End Property

Property Get Height() As Long
    Height = m_lDY
End Property

Property Get HDC() As Long
    HDC = m_hdcBitmap
End Property

Private Sub ClearUp()
    ' Here we ensure all our Objects are destroyed:
    Graphics.ClearDCBitmap m_hdcBitmap, m_hBmpBitmap, m_hBmpBitmapOld
    m_hdcBitmap = -1
End Sub

Private Sub Class_Initialize()
   m_hdcBitmap = -1
End Sub

Private Sub Class_Terminate()
    ClearUp
End Sub


