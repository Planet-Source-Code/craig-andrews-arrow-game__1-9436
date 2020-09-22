Attribute VB_Name = "Module1"
Option Explicit

Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" _
       (SoundName As Any, ByVal Flags As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public hill_ht As Double, hill_at As Double, hill_spread As Double, target_at As Double, target_ht As Double, you_at As Double, you_ht As Double

Public inputmode As String
Public InputBuffer As String
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wfilltype As Long) As Long
Public ShotDelay As Integer
Public SoundOn As Boolean

Public Const g = 32.16
Public Const wLeft = -50                    '! World coordinates, in feet
Public Const wRight = 200
Public Const wBottom = -50
Public Const wTop = 100

Sub MakeScene()
'Form1.Picture1.Picture = LoadResPicture("clouds", 0)
    Randomize
    'MsgBox Form1.Picture1.Width
    Form1.Picture1.Scale (wLeft, wTop)-(wRight, wBottom)
    Let hill_at = Int(50 * Rnd + 50) '! Hill peak location
    Let hill_ht = Int(30 * Rnd + 30) '! Hill height
    Let hill_spread = Int(10 * Rnd + 20) '! Hill spread

    Let target_at = Int(100 * Rnd + hill_at)
    Let target_ht = Hill(target_at)

    Let you_at = 0
    Let you_ht = Hill(you_at)

    Call DrawHill
    Call DrawYou
    Call DrawTarget
    'Form1.Picture1.Refresh
End Sub

Sub DrawHill()
'Const GroundColor = vbGreen
Const HillColor = vbBlack
Const GroundColor = 44800
'GroundColor = RGB(0, 175, 0)
Form1.Picture1.CurrentX = wLeft
Form1.Picture1.CurrentY = Hill(wLeft)

'Draw Sky
Dim i As Double
Dim j As Double
For j = wTop To wBottom Step -1 * Form1.Sky.Height
    For i = wLeft To wRight Step Form1.Sky.Width
        Form1.Picture1.PaintPicture Form1.Sky.Picture, i, j
    Next i
Next j

Dim X As Double
    For X = wLeft To wRight
        'PLOT LINES: x,hill(x);
        'Form1.Picture1.CurrentX = Form1.Picture1.Width
        'Form1.Picture1.CurrentY = Form1.Picture1.Height
        Form1.Picture1.Line -(X, Hill(X)), HillColor
        'Form1.Picture1.Line (X, wBottom)-(X, Hill(X)), QBColor(2)
        'Form1.Picture1.Line (X, wTop)-(X, Hill(X)), QBColor(3)
    Next X
    'PLOT
    'Draw Sky and Hill
    Const FLOODFILLBORDER = 0  ' Fill until crColor& color encountered.
    Const FLOODFILLSURFACE = 1 ' Fill surface until crColor& color not
                             ' encountered.
'Flood Form1.Picture1, FLOODFILLBORDER, vbBlue, vbBlack, 0, 0
'Draw Ground
Form1.Picture1.CurrentX = wLeft + 1
Form1.Picture1.CurrentY = wBottom + 1
Form1.Picture1.ScaleMode = 3 'Pixel
Flood Form1.Picture1, FLOODFILLBORDER, GroundColor, HillColor, Form1.Picture1.CurrentX, Form1.Picture1.CurrentY
Form1.Picture1.Scale (wLeft, wTop)-(wRight, wBottom)
End Sub

Sub DrawYou()
    Form1.Picture1.DrawWidth = 2
    'MsgBox you_at
    'Form1.Picture1.Line (you_at - 2, you_ht - 2)-(you_at - 2, you_ht + 2), QBColor(3)
    'Form1.Picture1.Line (you_at - 2, you_ht - 2)-(you_at + 2, you_ht - 2), QBColor(3)
    'Form1.Picture1.Line (you_at + 2, you_ht + 2)-(you_at - 2, you_ht + 2), QBColor(3)
    'Form1.Picture1.Line (you_at + 2, you_ht + 2)-(you_at + 2, you_ht - 2), QBColor(3)
    'Form1.Picture1.PaintPicture Form1.PetePic.Picture, you_at - 4, Hill(you_at - 2) + Form1.PetePic.Height
    Form1.Picture1.CurrentX = you_at - 4
    Form1.Picture1.CurrentY = Hill(you_at - 4) + Form1.PetePic.Height
    Form1.Picture1.ScaleMode = 3 'Pixel
      Dim R As RECT
  With R
   .Left = 0
   .Top = 0
   .Right = Form1.PetePic.ScaleWidth - 0.5
   .Bottom = Form1.PetePic.ScaleHeight - 1
  End With

  TransparentBlt Form1.Picture1.hdc, Form1.Picture1.hdc, Form1.PetePic.hdc, R, Form1.Picture1.CurrentX, Form1.Picture1.CurrentY, vbMagenta
    Form1.Picture1.Scale (wLeft, wTop)-(wRight, wBottom)

    
    Form1.Picture1.DrawWidth = 1
    Form1.Picture1.CurrentX = you_at + 2
    Form1.Picture1.CurrentY = Hill(you_at + 2) + 10
    Form1.Picture1.ForeColor = vbRed
    Form1.Picture1.Print "You"
End Sub

Sub DrawTarget(Optional Hit As Boolean)
If Hit = False Then
    Form1.TargetPic.Picture = LoadResPicture("targetnorm", vbResBitmap)
Else
    Form1.TargetPic.Picture = LoadResPicture("targethit", vbResBitmap)
End If
    Form1.Picture1.CurrentX = target_at - Form1.TargetPic.Width / 2
    Form1.Picture1.CurrentY = Hill(target_at) + Form1.TargetPic.Height / 2
    Form1.Picture1.ScaleMode = 3 'Pixel
      Dim R As RECT
  With R
   .Left = 0
   .Top = 0
   .Right = Form1.TargetPic.ScaleWidth - 1
   .Bottom = Form1.TargetPic.ScaleHeight - 1
  End With

  TransparentBlt Form1.Picture1.hdc, Form1.Picture1.hdc, Form1.TargetPic.hdc, R, Form1.Picture1.CurrentX, Form1.Picture1.CurrentY, vbWhite
    Form1.Picture1.Scale (wLeft, wTop)-(wRight, wBottom)

'Form1.Picture1.Circle (target_at, target_ht), 2, QBColor(4)
Form1.Picture1.CurrentX = target_at + 2
Form1.Picture1.CurrentY = Hill(target_at + 2) + 6
Form1.Picture1.ForeColor = vbRed
Form1.Picture1.Print "Target"
End Sub

Sub PlayGame()
Dim again As String
Dim tries As Integer
Do
Form1.Picture1.Cls
MakeScene
    Let tries = 0
    Do
       Let tries = tries + 1
       'MsgBox tries
    Loop Until MakeShot = True
    PrintText "Congratulations, you hit the target!"
    PrintText "It took you " & tries & " tries."
    Do
    PrintText "Do you want to play again? ", 1
    again = GetString(1)
    PrintText
    again = LCase(again)
    Loop Until again = "y" Or again = "n"
    If again = "n" Then Exit Do
    Form1.Text1.Text = ""
Loop
MsgBox "Thanks for Playing!"
EndAll
End Sub

Function Hill(qx As Double)
    '! Hill_ht and hill_at are defined in MakeScene
    Let Hill = hill_ht / (1 + ((qx - hill_at) / hill_spread) ^ 2)
End Function

Function MakeShot() As Boolean
Dim velx As Double, vely As Double, angle As Double, vel As Double
Dim X As Double
Dim Y As Double
Const delta = 0.01
Dim pi As Double
Let pi = 4 * Atn(1)
    'WINDOW #2
    'INPUT prompt "Enter angle, velocity (e.g., 25, 30): ": angle, vel
    'Let angle = 60
    'Let vel = 70
    Do
        PrintText "Input Angle: ", 1
        Let angle = GetNumber
        If angle < 180 And angle > 0 Then Exit Do
        PrintText "Angle must be between 0 and 180"
    Loop
    PrintText "Input Velocity: ", 1
    Let vel = GetNumber
    PrintText "Press CTRL+C to stop the shot..."
    inputmode = "all"
    Let velx = vel * Cos(angle * pi / 180)
    Let vely = vel * Sin(angle * pi / 180)

    'WINDOW #1
    Let X = you_at
    Let Y = you_ht
    Dim t As Double
    For t = 0 To 1000 Step delta  '! More than enough time (seconds)
    If InStr(InputBuffer, Chr(3)) > 0 Then
        PrintText "CTRL+C Pressed. Shot Stopped!"
        Exit For
    End If
    If X > wRight Or X < wLeft Then
        MakeShot = False
        Exit For
    End If
        'PLOT x,y;
        Form1.Picture1.PSet (X, Y), vbMagenta
        Form1.Picture1.Refresh

        '! Check for hitting the ground, or hitting the target

        If (X - target_at) ^ 2 + (Y - target_ht) ^ 2 < 4 Then

           '! Within two units of the target, so score a hit.
           DrawTarget True
           DoEvents
           'DRAW Burst with shift(target_at, target_ht)
           'LET bullseye = 1       ! Score a hit
            
            If SoundOn = True Then
                Dim bSound() As Byte
                Dim R As Long
                bSound = LoadResData("Hit", "WAVE")
                R = sndPlaySound(bSound(0), 4)
            End If

           MakeShot = True
           Exit For               '! ... and leave the loop

        ElseIf t > 0 And Y < Hill(X) Then  '! Just hit the ground
           'Let bullseye = 0       '! Score a miss
           MakeShot = False
           Exit For               '! ... and leave the loop

        End If

        '! Not yet at the ground, so try again.

        Let X = X + velx * delta  '! Newton's Laws of Motion
        Let Y = Y + vely * delta
        Let vely = vely - g * delta    '! ... effect of gravity
        DoEvents
        Call Sleep(ShotDelay)
    
    Next t
    'PLOT
'MsgBox "Done!"
End Function

Public Function GetNumber() As Double
inputmode = "number"
InputBuffer = ""
Dim i As Integer
Do
    'MsgBox "test"
    DoEvents
    Sleep 1
Loop Until Right(InputBuffer, 2) = vbCrLf
'InputBuffer = 10 & vbCrLf
Let GetNumber = Val(Left(InputBuffer, Len(InputBuffer) - 2))
'MsgBox GetNumber
inputmode = ""
InputBuffer = ""
End Function

Public Function GetString(Length As Integer) As String
inputmode = "string"
InputBuffer = ""
Dim i As Integer
Do
    DoEvents
    Sleep 1
Loop Until Right(InputBuffer, 2) = vbCrLf Or Len(InputBuffer) = Length
If InputBuffer = vbCrLf Then
    GetString = vbCrLf
    inputmode = ""
    Exit Function
End If
If Len(InputBuffer) = Length Then GetString = InputBuffer
If Right(InputBuffer, 2) = vbCrLf Then GetString = Left(InputBuffer, Len(InputBuffer) - 2)
inputmode = ""
InputBuffer = ""
End Function

Public Sub PrintText(Optional what As String, Optional DontPressEnter As Integer)
Form1.Text1.Text = Form1.Text1.Text & what
If DontPressEnter = 0 Then Form1.Text1.Text = Form1.Text1.Text & vbCrLf
End Sub

'Public Sub Burst()
'    Dim radius As Integer
'    Let radius = 8
'    Dim a As Double
'    For a = 0 To 360 Step 22.5    '! Eight-pointed star
'        'Form1.Picture1.Line -(radius * Cos(a * pi / 180), radius * Sin(a))
'        Form1.Picture1.Line (target_at, target_ht)-(radius * Cos(a * pi / 180), radius * Sin(a * pi / 180))
'        Let radius = 12 - radius
'    Next a
'End Sub
