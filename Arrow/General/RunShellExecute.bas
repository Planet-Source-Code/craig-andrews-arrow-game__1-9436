Attribute VB_Name = "RunShellExec"
Option Explicit

'No matter where this code goes, or what happens to it,
'I, Craig Andrews, still retain all rights to it.
'You may modify the code for your own use, but
'may not claim the original concept as your own.
'Please credit me in your about box and in any accompanying
'documentation.
'Released open source on July 1st, 2000

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10

Public Sub EmailTo(who As String)
Call RunShellExecute("Open", "mailto:" & who, 0&, 0&, SW_SHOWNORMAL)
End Sub

Public Sub BrowseTo(what As String)
Call RunShellExecute("Open", what, 0&, 0&, SW_SHOWNORMAL)
End Sub

Public Sub RunShellExecute(sTopic As String, sFile As Variant, _
                           sParams As Variant, sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  If success < 32 Then
     Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  End If
   
End Sub


