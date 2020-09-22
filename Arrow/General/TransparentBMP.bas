Attribute VB_Name = "TransparentBMP"
Option Explicit

'No matter where this code goes, or what happens to it,
'I, Craig Andrews, still retain all rights to it.
'You may modify the code for your own use, but
'may not claim the original concept as your own.
'Please credit me in your about box and in any accompanying
'documentation.
'Released open source on July 1st, 2000

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Declare Function BitBlt Lib "gdi32" _
  (ByVal hDCDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal XSrc As Long, ByVal YSrc As Long, _
   ByVal dwRop As Long) As Long

Public Declare Function CreateBitmap Lib "gdi32" _
  (ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal nPlanes As Long, _
   ByVal nBitCount As Long, _
   lpBits As Any) As Long

Public Declare Function SetBkColor Lib "gdi32" _
   (ByVal hdc As Long, ByVal crColor As Long) As Long

Public Declare Function SelectObject Lib "gdi32" _
   (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" _
   (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" _
   (ByVal hdc As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" _
   (ByVal hdc As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" _
   (ByVal hObject As Long) As Long

Public Sub TransparentBlt(OutDstDC As Long, _
  DstDC As Long, SrcDC As Long, SrcRect As RECT, _
  DstX As Integer, DstY As Integer, TransColor As Long)
   
  'DstDC- Device context into which image must be
  'drawn transparently

  'OutDstDC- Device context into image is actually drawn,
  'even though it is made transparent in terms of DstDC

  'Src- Device context of source to be made transparent
  'in color TransColor

  'SrcRect- Rectangular region within SrcDC to be made
  'transparent in terms of DstDC, and drawn to OutDstDC

  'DstX, DstY - Coordinates in OutDstDC (and DstDC)
  'where the transparent bitmap must go. In most
  'cases, OutDstDC and DstDC will be the same
   
  Dim nRet As Long, W As Integer, H As Integer
  Dim MonoMaskDC As Long, hMonoMask As Long
  Dim MonoInvDC As Long, hMonoInv As Long
  Dim ResultDstDC As Long, hResultDst As Long
  Dim ResultSrcDC As Long, hResultSrc As Long
  Dim hPrevMask As Long, hPrevInv As Long
  Dim hPrevSrc As Long, hPrevDst As Long

  W = SrcRect.Right - SrcRect.Left + 1
  H = SrcRect.Bottom - SrcRect.Top + 1
   
 'create monochrome mask and inverse masks
  MonoMaskDC = CreateCompatibleDC(DstDC)
  MonoInvDC = CreateCompatibleDC(DstDC)
  hMonoMask = CreateBitmap(W, H, 1, 1, ByVal 0&)
  hMonoInv = CreateBitmap(W, H, 1, 1, ByVal 0&)
  hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
  hPrevInv = SelectObject(MonoInvDC, hMonoInv)
   
 'create keeper DCs and bitmaps
  ResultDstDC = CreateCompatibleDC(DstDC)
  ResultSrcDC = CreateCompatibleDC(DstDC)
  hResultDst = CreateCompatibleBitmap(DstDC, W, H)
  hResultSrc = CreateCompatibleBitmap(DstDC, W, H)
  hPrevDst = SelectObject(ResultDstDC, hResultDst)
  hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
   
'copy src to monochrome mask
  Dim OldBC As Long
  OldBC = SetBkColor(SrcDC, TransColor)
  nRet = BitBlt(MonoMaskDC, 0, 0, W, H, SrcDC, _
                SrcRect.Left, SrcRect.Top, vbSrcCopy)
  TransColor = SetBkColor(SrcDC, OldBC)
   
 'create inverse of mask
  nRet = BitBlt(MonoInvDC, 0, 0, W, H, _
                MonoMaskDC, 0, 0, vbNotSrcCopy)
   
 'get background
  nRet = BitBlt(ResultDstDC, 0, 0, W, H, _
                DstDC, DstX, DstY, vbSrcCopy)
   
 'AND with Monochrome mask
  nRet = BitBlt(ResultDstDC, 0, 0, W, H, _
                MonoMaskDC, 0, 0, vbSrcAnd)
   
 'get overlapper
  nRet = BitBlt(ResultSrcDC, 0, 0, W, H, SrcDC, _
                SrcRect.Left, SrcRect.Top, vbSrcCopy)
   
 'AND with inverse monochrome mask
  nRet = BitBlt(ResultSrcDC, 0, 0, W, H, _
                MonoInvDC, 0, 0, vbSrcAnd)
   
'XOR these two
  nRet = BitBlt(ResultDstDC, 0, 0, W, H, _
                ResultSrcDC, 0, 0, vbSrcInvert)
   
 'output results
  nRet = BitBlt(OutDstDC, DstX, DstY, W, H, _
                ResultDstDC, 0, 0, vbSrcCopy)
   
 'clean up
  hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
  DeleteObject hMonoMask

  hMonoInv = SelectObject(MonoInvDC, hPrevInv)
  DeleteObject hMonoInv

  hResultDst = SelectObject(ResultDstDC, hPrevDst)
  DeleteObject hResultDst

  hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
  DeleteObject hResultSrc

  DeleteDC MonoMaskDC
  DeleteDC MonoInvDC
  DeleteDC ResultDstDC
  DeleteDC ResultSrcDC

End Sub


