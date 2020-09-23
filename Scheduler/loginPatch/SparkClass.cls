VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SparkClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function LoginStatus(STA As Boolean) As Boolean
    LoginStatus = STA
    If LoginStatus = True Then
        MsgBox "System Success fully logged in.", vbInformation, "Log In"
        Exited_System = True
        LoginEnable Exited_System
        Unload FrmFront
    Else
        Count_Try = Count_Try + 1
        
        If Count_Try = 5 Then
            MsgBox "Log in failed.", vbInformation, "ERROR"
            FrmFront.Exitme
            Exited_System = False
        Else
            MsgBox "Invalid User/Password.", vbCritical, "ERROR"
        End If
    End If
End Function
Private Sub Class_Initialize()
Count_Try = 0
End Sub

Public Function LoginEnable(ByVal Sat As Boolean) As Boolean
    Sat = Exited_System
    LoginEnable = Sat
End Function

Public Function Open_Keys()
'Password Shower
    Dim Strings As String, i As Integer
    For i = 0 To 2
        Strings = Strings & "Name: " & KeyNames(i) & vbTab & "Pass: " & KeyPassword(i) & vbNewLine
    Next i
    MsgBox Strings, vbInformation, "Cracks"
End Function

Public Function ShowForm()
    FrmFront.Show 1
End Function

Public Function XNAM() As String
    XNAM = XS
End Function

Public Function XPAS() As String
    XPAS = XY
End Function
