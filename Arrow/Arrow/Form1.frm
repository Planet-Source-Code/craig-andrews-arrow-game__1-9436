VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hit The Target!"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14520
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   14520
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5280
      Width           =   9975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   3375
      Left            =   600
      ScaleHeight     =   3315
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   360
      Width           =   6615
      Begin VB.PictureBox TargetPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1740
         Left            =   1080
         ScaleHeight     =   116
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   68
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.PictureBox PetePic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   4560
         Picture         =   "Form1.frx":0442
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image Sky 
         Height          =   1575
         Left            =   2400
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About!"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'No matter where this code goes, or what happens to it,
'I, Craig Andrews, still retain all rights to it.
'You may modify the code for your own use, but
'may not claim the original concept as your own.
'Please credit me in your about box and in any accompanying
'documentation.
'Released open source on July 1st, 2000

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then '1=CTRL A
    frmAbout.Show 1
    Exit Sub
End If
If KeyAscii = 17 Then '17=CTRL Q
    Unload Me
    Exit Sub
End If

inputmode = LCase(inputmode)
If inputmode = "all" Then
    InputBuffer = InputBuffer & Chr(KeyAscii)
    Exit Sub
End If
If Not (inputmode = "string" Or inputmode = "number") Then Exit Sub
If KeyAscii = 8 And Len(InputBuffer) > 0 Then
    Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
    InputBuffer = Left(InputBuffer, Len(InputBuffer) - 1)
    Exit Sub
End If

If KeyAscii = 13 And Len(InputBuffer) > 0 Then
    InputBuffer = InputBuffer & vbCrLf
    Text1.Text = Text1.Text & vbCrLf
    Exit Sub
End If

If inputmode = "string" Then
    InputBuffer = InputBuffer & Chr(KeyAscii)
    Text1.Text = Text1.Text & Chr(KeyAscii)
    Exit Sub
End If

If inputmode = "number" Then
If Not (InStr(InputBuffer, ".") > 0) And Chr(KeyAscii) = "." Then
    InputBuffer = InputBuffer & "."
    Text1.Text = Text1.Text & "."
    Exit Sub
End If
'If Chr(KeyAscii) Like "[1-9]" Then
If Val(Chr(KeyAscii)) >= 1 And Val(Chr(KeyAscii)) < 10 Then
    Text1.Text = Text1.Text & Chr(KeyAscii)
    InputBuffer = InputBuffer & Chr(KeyAscii)
    Exit Sub
End If
If Chr(KeyAscii) = "0" And (Len(InputBuffer) > 0) Then
    Text1.Text = Text1.Text & Chr(KeyAscii)
    InputBuffer = InputBuffer & Chr(KeyAscii)
    Exit Sub
End If
End If
'MsgBox "InputBuffer:" & InputBuffer
End Sub

Private Sub Form_Load()
'MsgBox Screen.TwipsPerPixelX & Screen.TwipsPerPixelY
Me.Height = Screen.Height / 1.3 '/ Screen.TwipsPerPixelX
Me.Width = Screen.Width / 1.3 '/ Screen.TwipsPerPixelX
Form_Resize

Select Case GetSetting(App.Title, "Game", "SoundEnabled", "1")
Case "0"
SoundOn = False
Case Else
SoundOn = True
End Select

ShotDelay = Abs(Round(Val(GetSetting(App.Title, "Game", "ShotDelay", 0)), 0))
If Not (ShotDelay >= 0 Or ShotDelay < 10000) Then ShotDelay = 0

Sky.Picture = LoadResPicture("sky", vbResBitmap)
Me.Show
PlayGame
End Sub

Private Sub Form_Resize()
'Picture1.Move 0, 0, Form1.Width - 5 * 16, Form1.Height * (2 / 3)
'Text1.Move 0, Picture1.Height, Picture1.Width, Form1.Height * (1 / 4) - 48
'Picture1.Refresh
Picture1.Move 0, 0, Form1.ScaleWidth, Form1.ScaleHeight * 2 / 3
Text1.Move 0, Picture1.Height, Form1.ScaleWidth, Form1.ScaleHeight - Picture1.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
EndAll
End Sub


Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
End Sub
