Attribute VB_Name = "SUBMODE"
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub Main()
FrmLoad.Show
End Sub
