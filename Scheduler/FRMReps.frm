VERSION 5.00
Begin VB.Form FRMReps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTS"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMReps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FSubjects 
      Height          =   2340
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Visible         =   0   'False
      Width           =   5865
      Begin VB.ComboBox CBYears 
         Enabled         =   0   'False
         Height          =   360
         ItemData        =   "FRMReps.frx":57E2
         Left            =   5325
         List            =   "FRMReps.frx":57F5
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   825
         Width           =   465
      End
      Begin VB.TextBox TCourses 
         Enabled         =   0   'False
         Height          =   390
         Left            =   3900
         TabIndex        =   9
         Top             =   825
         Width           =   1440
      End
      Begin VB.TextBox Tsubs 
         Enabled         =   0   'False
         Height          =   390
         Left            =   3900
         TabIndex        =   7
         Top             =   300
         Width           =   1890
      End
      Begin VB.Frame Frame1 
         Caption         =   "Report Option"
         Height          =   1515
         Left            =   75
         TabIndex        =   2
         Top             =   75
         Width           =   1965
         Begin VB.OptionButton O1 
            Caption         =   "Select Course/Yr"
            Height          =   240
            Index           =   2
            Left            =   75
            TabIndex        =   5
            Top             =   1125
            Width           =   1815
         End
         Begin VB.OptionButton O1 
            Caption         =   "Select Subject"
            Height          =   240
            Index           =   1
            Left            =   75
            TabIndex        =   4
            Top             =   750
            Width           =   1665
         End
         Begin VB.OptionButton O1 
            Caption         =   "Report All"
            Height          =   240
            Index           =   0
            Left            =   75
            TabIndex        =   3
            Top             =   375
            Value           =   -1  'True
            Width           =   1290
         End
      End
      Begin VB.CommandButton CbRepSub 
         Caption         =   "OK"
         Height          =   540
         Left            =   4575
         TabIndex        =   1
         Top             =   1725
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enter Course Here:"
         Height          =   240
         Left            =   2175
         TabIndex        =   8
         Top             =   825
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Enter Subject Here:"
         Height          =   240
         Left            =   2175
         TabIndex        =   6
         Top             =   300
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FRMReps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mysql1 As String, mysql2 As String, myshape As String
Private Enab As Integer
Private Sub CbRepSub_Click()
Create_SubjectReport
End Sub

Private Sub Form_Load()
Select Case RepActive
Case 1
    FSubjects.Visible = True
End Select
End Sub

Private Sub O1_Click(Index As Integer)
Select Case Index
    Case 0 'No Enable
        Tsubs.Enabled = False
        TCourses.Enabled = False
        CBYears.Enabled = False
    Case 1 'Enable Subjects
        Tsubs.Enabled = True
        TCourses.Enabled = False
        CBYears.Enabled = False
    Case 2 'Enable Course/yr
        Tsubs.Enabled = False
        TCourses.Enabled = True
        CBYears.Enabled = True
End Select
Enab = Index
End Sub

Function Create_SubjectReport()
With Denver
    If .rsSubHead.State <> 0 Then .rsSubHead.Close
    Select Case Enab
        Case 0
            mysql1 = "SELECT Course, Yr FROM Subjects GROUP BY Course, yr ORDER BY course,yr"
            mysql2 = "Select * from Subjects"
        Case 1
            mysql1 = "SELECT Course, Yr FROM Subjects GROUP BY Course, yr ORDER BY course,yr"
            mysql2 = "Select * from Subjects where SC = " & Quoted(Tsubs.Text)
        Case 2
            mysql1 = "SELECT Course, Yr FROM Subjects where course = " & Quoted(TCourses.Text) & " and yr = " & Quoted(CBYears.Text) & "Group by course,yr"
            mysql2 = "Select * from Subjects "
    End Select
    myshape = " SHAPE {" & mysql1 & "}  AS SubHead APPEND ({" & mysql2 & "}  AS Subjects RELATE 'Course' TO 'Course','Yr' TO 'Yr') AS Subjects"
    .rsSubHead.Open myshape, .MyCON, adOpenDynamic, adLockOptimistic
End With
With RPTSubject
    .Sections("PageHeader").Controls("Lyear").Caption = MainForm.TSy.Text
    .Sections("PageHeader").Controls("Lschool").Caption = UCase(InputBox("Enter School:", "School"))
    .Sections("PageHeader").Controls("LSem").Caption = UCase(InputBox("Enter Semester:", "Semester"))
    .Refresh
    .Show 1
End With
End Function
