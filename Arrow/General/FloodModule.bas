Attribute VB_Name = "FloodModule"
Option Explicit

'No matter where this code goes, or what happens to it,
'I, Craig Andrews, still retain all rights to it.
'You may modify the code for your own use, but
'may not claim the original concept as your own.
'Please credit me in your about box and in any accompanying
'documentation.
'Released open source on July 1st, 2000

Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wfilltype As Long) As Long

Public Sub Flood(ByVal WhatPic As Object, style As Integer, FillColor As Long, StopColor As Long, X As Long, Y As Long)
'Dim TempAuto As Boolean
'TempAuto = WhatPic.AutoRedraw
'WhatPic.AutoRedraw = False
Dim TempColor As Long
TempColor = WhatPic.FillColor
WhatPic.FillColor = FillColor
Dim tempStyle As Integer
tempStyle = WhatPic.FillStyle
WhatPic.FillStyle = 0
'Dim TempMode As Integer
'TempMode = WhatPic.ScaleMode
'WhatPic.ScaleMode = 3


' Make sure that the FillStyle is not transparent.
 ' crColor& specifies the color for the boundary.
  'Const FLOODFILLBORDER = 0  ' Fill until crColor& color encountered.
  'Const FLOODFILLSURFACE = 1 ' Fill surface until crColor& color not
                             ' encountered.
  'crColor = RGB(0, 0, 0)
  Dim wfilltype As Long
  'wfilltype = FLOODFILLBORDER 'FLOODFILLSURFACE
  Dim suc As Long
  suc = ExtFloodFill(WhatPic.hdc, X, Y, StopColor, style)
  'Picture1.Print "Hello"
  'MsgBox suc
'WhatPic.AutoRedraw = TempAuto
WhatPic.FillColor = TempColor
WhatPic.FillStyle = tempStyle
'WhatPic.ScaleMode = TempMode
End Sub
