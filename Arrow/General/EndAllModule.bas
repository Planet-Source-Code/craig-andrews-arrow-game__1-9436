Attribute VB_Name = "EndAllModule"
Option Explicit

Public Sub EndAll()
Dim ret As Object
For Each ret In Forms
    Unload ret
    Set ret = Nothing
Next
End
End Sub
