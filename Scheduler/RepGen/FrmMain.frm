VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Generator"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnroom 
      Caption         =   "Room Schedule"
      Height          =   735
      Left            =   6840
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton BtnTeach 
      Caption         =   "Teacher Schedule"
      Height          =   735
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton BtnClass 
      Caption         =   "Class Schedule"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "REPORT GENERATOR. PLEASE THE TYPE OF REPORT TO GENERATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnClass_Click()
FrmViewer.FrmReps.Visible = 1
FrmViewer.Show 1
End Sub


Private Sub btnroom_Click()
FrmViewer.FrmRoom.Visible = 1
FrmViewer.Show 1
End Sub

Private Sub BtnTeach_Click()
FrmViewer.FrmTeacher.Visible = 1
FrmViewer.Show 1
End Sub

