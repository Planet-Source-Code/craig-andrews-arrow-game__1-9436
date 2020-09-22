VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   2745
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5130
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   183
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   342
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox chkSound 
      Caption         =   "Check1"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   195
   End
   Begin VB.TextBox txtDelay 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Text            =   "10000"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3720
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   3720
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sound Enabled"
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label GotoWeb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visit my website!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "http://www.compucrafters.com/software"
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shot Delay Time (in milliseconds-default is 0):"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   3165
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   3960
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label EmailLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email me at candrews@compucrafters.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   2985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   8
      X2              =   336
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   8
      X2              =   336
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   360
      Width           =   525
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning: ..."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3945
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
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

Private Const SW_SHOW = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
If Abs(Round(Val(txtDelay.Text), 0)) = txtDelay.Text And 10000 > Abs(Round(Val(txtDelay.Text), 0)) And 0 <= Abs(Round(Val(txtDelay.Text), 0)) Then
    SaveSetting App.Title, "Game", "ShotDelay", txtDelay.Text
    ShotDelay = Abs(Round(Val(txtDelay.Text), 0))
Else
    MsgBox "Delay must be a positive integer between 0 and 10,000 (0 is off)"
    txtDelay.Text = Abs(Round(Val(txtDelay.Text)))
End If
SaveSetting App.Title, "Game", "SoundEnabled", chkSound.Value
Select Case chkSound.Value
    Case 1
        SoundOn = True
    Case Else
        SoundOn = False
End Select
Unload Me
End Sub

Private Sub EmailLink_Click()
EmailTo "candrews@compucrafters.net"
End Sub

Private Sub EmailLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
EmailLink.ForeColor = &HFF&
GotoWeb.ForeColor = &HC00000
End Sub

Private Sub Form_Load()
Select Case SoundOn
Case True
chkSound.Value = 1
Case False
chkSound.Value = 0
End Select

txtDelay.Text = ShotDelay
Image1.Picture = LoadResPicture("sky", vbResBitmap)
'frmAbout.Show
Dim i As Integer
Dim j As Integer
For j = 0 To frmAbout.ScaleHeight Step Image1.Height
    For i = 0 To frmAbout.ScaleWidth Step Image1.Width
        frmAbout.PaintPicture Image1.Picture, i, j
    Next i
Next j
    Me.Icon = Form1.Icon
    Me.Caption = "About " & App.Title
    lblDescription = App.FileDescription
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDisclaimer.Caption = "This product is public domain, and is freely copyable. For more info, email candrews@mediaone.net"
    PaintPicture Form1.Icon, 16, 8
    
Dim R As RECT
With R
    .Left = 0
    .Top = 0
    .Right = Form1.PetePic.ScaleWidth - 0.5
    .Bottom = Form1.PetePic.ScaleHeight - 1
End With
TransparentBlt frmAbout.hdc, frmAbout.hdc, Form1.PetePic.hdc, R, 216, 140, vbMagenta
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
EmailLink.ForeColor = &HC00000
GotoWeb.ForeColor = &HC00000
End Sub

Private Sub GotoWeb_Click()
BrowseTo GotoWeb.ToolTipText
End Sub

Private Sub GotoWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
GotoWeb.ForeColor = &HFF&
EmailLink.ForeColor = &HC00000
End Sub

Private Sub Label2_Click()
Select Case chkSound.Value
Case 0
chkSound.Value = 1
Case Else
chkSound.Value = 0
End Select
End Sub
