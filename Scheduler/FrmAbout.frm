VERSION 5.00
Begin VB.Form FrmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About HX Scheduler 2"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CBOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OKAY"
      Height          =   465
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmAbout.frx":27A2
      ForeColor       =   &H00FFFFFF&
      Height          =   1140
      Left            =   1350
      TabIndex        =   3
      Top             =   1350
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   75
      Picture         =   "FrmAbout.frx":2852
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   1185
   End
   Begin VB.Label lbltx 
      BackStyle       =   0  'Transparent
      Caption         =   "Scheduling System for Schools. A revision of the HX Scheduler 2.0.2."
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   1350
      TabIndex        =   2
      Top             =   525
      Width           =   4815
   End
   Begin VB.Label Lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HX Scheduler 2.00003995"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1275
      TabIndex        =   1
      Top             =   150
      Width           =   4020
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   75
      Picture         =   "FrmAbout.frx":4BCF7
      Stretch         =   -1  'True
      Top             =   75
      Width           =   1215
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function xxx()

Dim i As Integer, X As Double
        ScaleMode = vbPixels
        ScaleWidth = 256
        DrawWidth = 4
        For i = 0 To 255
        Line (X, 0)-(X, Height), RGB(0, 0, i), BF
        X = X + 1
        Next

End Function

Private Sub CBOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Lbltitle.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
xxx
End Sub
