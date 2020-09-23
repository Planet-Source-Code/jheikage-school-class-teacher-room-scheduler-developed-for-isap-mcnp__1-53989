VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "loading..."
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label LBL1 
      Caption         =   "Loading Help file. Please wait..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "FrmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PB_MOVE_R As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Form_Load()
ShowCursor False
'Couse_Sleep
SetWindowPos Me.hwnd, -1, Screen.Width / 2 - Me.Width / 2, Screen.Height / 2 - Me.Height / 2, 330, 90, &H10
End Sub

Private Sub TX_Timer()
PB_MOVE_R = Rnd(100)
PB.Value = PB.Value + PB_MOVE_R
If PB.Value >= 100 Then
    LBL1.Caption = "Complete. Please wait..."
    FRMHELP.Show
    Unload Me
End If
End Sub

Sub Couse_Sleep()
PB_MOVE_R = 10
Do Until PB.Value = 100
PB.Value = PB.Value + PB_MOVE_R
If PB.Value >= 100 Then
    LBL1.Caption = "Complete. Please wait..."
    FRMHELP.Show
    Unload Me
End If
Sleep 100000
Loop
End Sub
