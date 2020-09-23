VERSION 5.00
Begin VB.Form FrmLock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Unlock"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtLock 
      Alignment       =   2  'Center
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   0
      OLEDropMode     =   2  'Automatic
      PasswordChar    =   "®"
      TabIndex        =   0
      Top             =   0
      Width           =   4590
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "System is currently locked by user. Please enter the password to unlock system. Software security by iehjsucker."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   75
      TabIndex        =   1
      Top             =   450
      Width           =   4440
   End
End
Attribute VB_Name = "FrmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Caption = "Enter Password to unlock. Current User: " & Namx
End Sub

Private Sub TxtLock_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TxtLock.Text = PASX Then
        MsgBox "System Unlocked.", vbInformation, "Unlock"
        Unload Me
    Else
        MsgBox "Invalid Password.", vbInformation, "Unlock"
        TxtLock.Text = ""
        TxtLock.SetFocus
    End If
End If
End Sub
