VERSION 5.00
Begin VB.Form FrmFront 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5430
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FrmFront.frx":0000
   ScaleHeight     =   5430
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrCha 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change User Name and Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   2040
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   3855
      Begin VB.ListBox Lchange 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         ItemData        =   "FrmFront.frx":0152
         Left            =   120
         List            =   "FrmFront.frx":0154
         TabIndex        =   11
         Top             =   600
         Width           =   2415
      End
      Begin nigol.chameleonButton Ccancel 
         Height          =   495
         Left            =   2760
         TabIndex        =   12
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "MS Sans Serif"
         SIZE            =   0
         UND             =   0   'False
         BTYPE           =   5
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nigol.chameleonButton Cok 
         Height          =   495
         Left            =   2760
         TabIndex        =   13
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "MS Sans Serif"
         SIZE            =   0
         UND             =   0   'False
         BTYPE           =   5
         TX              =   "OK"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select a User Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2040
      End
   End
   Begin VB.Frame FrLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log in Form"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   3855
      Begin VB.ListBox LNames 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         ItemData        =   "FrmFront.frx":0156
         Left            =   480
         List            =   "FrmFront.frx":0158
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Tpass 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Tnam 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin nigol.chameleonButton Ccclose 
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "MS Sans Serif"
         SIZE            =   0
         UND             =   0   'False
         BTYPE           =   5
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nigol.chameleonButton Bok 
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "MS Sans Serif"
         SIZE            =   0
         UND             =   0   'False
         BTYPE           =   5
         TX              =   "OK"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin nigol.chameleonButton CLog 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BOLD            =   0   'False
      ITA             =   0   'False
      INAME           =   "MS Sans Serif"
      SIZE            =   0
      UND             =   0   'False
      BTYPE           =   6
      TX              =   "Log In"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   8421504
      FCOLO           =   8421504
      MCOL            =   12632256
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nigol.chameleonButton Copt 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BOLD            =   0   'False
      ITA             =   0   'False
      INAME           =   "MS Sans Serif"
      SIZE            =   0
      UND             =   0   'False
      BTYPE           =   6
      TX              =   "Log in Options"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   8421504
      FCOLO           =   8421504
      MCOL            =   12632256
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nigol.chameleonButton Cexit 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BOLD            =   0   'False
      ITA             =   0   'False
      INAME           =   "MS Sans Serif"
      SIZE            =   0
      UND             =   0   'False
      BTYPE           =   6
      TX              =   "Exit System"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   8421504
      FCOLO           =   8421504
      MCOL            =   12632256
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   240
      Picture         =   "FrmFront.frx":015A
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   4440
      Picture         =   "FrmFront.frx":19B4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All right reserved"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   840
      Width           =   1470
   End
   Begin VB.Label LVerse 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   960
      TabIndex        =   16
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEM SECURITY"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3900
   End
End
Attribute VB_Name = "FrmFront"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Xparks As New SparkClass
Private Sub Bok_Click()
'Check Name and Password
Dim i As Integer, logs As Boolean
For i = 0 To 2
If Tnam.Text = KeyNames(i) And _
    Tpass.Text = KeyPassword(i) Then
    logs = True
    XS = KeyNames(i)
    XY = KeyPassword(i)
    Call Xparks.XNAM
    Call Xparks.XPAS
    GoTo NOWCHECK
    Exit For
End If
Next
NOWCHECK:

Call Xparks.LoginStatus(logs)
End Sub

Private Sub Ccancel_Click()
FrCha.Visible = False
End Sub

Private Sub Ccclose_Click()
LNames.Visible = False
FrLog.Visible = False
Tnam.Text = ""
Tpass.Text = ""
End Sub

Private Sub Cexit_Click()
Exitme
End Sub

Public Sub Exitme()
Exited_System = False
Unload Me

End Sub
Private Sub CLog_Click()
FrLog.Visible = True
FrCha.Visible = False
End Sub

Private Sub Cok_Click()
'Select Change
Dim inNa As String, inNew As String
Dim pPass As String, pconfirm As String, pNew As String
inNa = InputBox("Enter Current User name: ", "User name", Lchange)
If inNa <> KeyNames(Indexman) Then Exit Sub
pPass = InputBox("Enter Current Password: ", "Passoword", "Password")
If pPass <> KeyPassword(Indexman) Then Exit Sub
inNew = InputBox("Enter new User name: ", "User name", Lchange)
If inNew = "" Then Exit Sub
pconfirm = InputBox("Enter new Password: ", "Passoword", "Password")
If pconfirm = "" Then Exit Sub
pNew = InputBox("Confirm Password: ", "Passoword", "Password")
If pNew <> pconfirm Then MsgBox "Confirmation do not match.Please Try again next time.", vbCritical, "ERROR": Exit Sub
'"NIGOL", "H_Key_Login", "KEYNAMES" & i, "Administrator" & i
'"Nigol", "H_Key_Login", "KEYPASS" & i, "crack" & i
SaveSetting "NIGOL", "H_Key_Login", "KEYNAMES" & Indexman, inNew
SaveSetting "Nigol", "H_Key_Login", "KEYPASS" & Indexman, pNew
MsgBox "You have changed a user information. This will take effect now.", vbInformation, "CHANGED"
Get_All_Keys
Pop_Hrd LNames
Pop_Hrd Lchange
End Sub

Private Sub Copt_Click()
FrCha.Visible = True
FrLog.Visible = False
Lchange.ListIndex = Indexman
End Sub

Private Sub Form_Load()
LVerse.Caption = App.Major & "." & App.Minor & "." & App.Revision
LX.Caption = "Copyright " & App.LegalCopyright & " all rights reserved."
Get_All_Keys
Pop_Hrd LNames
Pop_Hrd Lchange
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 5 Then
        FrmPassName.Show 1
    End If
End Sub

Private Sub Lchange_Click()
Indexman = Lchange.ListIndex
End Sub

Private Sub LNames_Click()
    Tnam.Text = LNames.Text
    LNames.Visible = False
End Sub

Private Sub Tnam_Click()
    LNames.Visible = True
End Sub
