VERSION 5.00
Begin VB.Form FrmAddNeed 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Record"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAddNeed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CBCancel 
      Caption         =   "CANCEL"
      Height          =   465
      Left            =   3450
      TabIndex        =   7
      Top             =   2550
      Width           =   1215
   End
   Begin VB.CommandButton CBOk 
      Caption         =   "OK"
      Height          =   465
      Left            =   2160
      TabIndex        =   6
      Top             =   2550
      Width           =   1215
   End
   Begin VB.PictureBox PSec 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   120
      ScaleHeight     =   2310
      ScaleWidth      =   4560
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   4590
      Begin VB.ComboBox CbxSec 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmAddNeed.frx":57E2
         Left            =   3960
         List            =   "FrmAddNeed.frx":57E4
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   480
         Width           =   570
      End
      Begin VB.ComboBox CBYrsec 
         Height          =   360
         ItemData        =   "FrmAddNeed.frx":57E6
         Left            =   3240
         List            =   "FrmAddNeed.frx":57F9
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   480
         Width           =   690
      End
      Begin VB.ComboBox CBCoursesec 
         Height          =   360
         ItemData        =   "FrmAddNeed.frx":580C
         Left            =   1650
         List            =   "FrmAddNeed.frx":581F
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   450
         Width           =   1125
      End
      Begin VB.TextBox TSectionsec 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1875
         Width           =   3165
      End
      Begin VB.Label Label10 
         Caption         =   "Note: You must put spaces between parameters (COURSE YR SECTION) ex: BSIT 1 B"
         Height          =   735
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Yr:"
         Height          =   240
         Left            =   2880
         TabIndex        =   21
         Top             =   450
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Available Course:"
         Height          =   240
         Left            =   75
         TabIndex        =   19
         Top             =   450
         Width           =   1500
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Section Record Manipulation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   75
         TabIndex        =   16
         Top             =   75
         Width           =   4440
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Section Name:"
         Height          =   240
         Left            =   75
         TabIndex        =   15
         Top             =   1950
         Width           =   1260
      End
   End
   Begin VB.PictureBox Proom 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   120
      ScaleHeight     =   2310
      ScaleWidth      =   4560
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   4590
      Begin VB.TextBox Troomr 
         Height          =   360
         Left            =   975
         MaxLength       =   70
         TabIndex        =   29
         Top             =   450
         Width           =   3540
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Room Record Manipulation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   75
         TabIndex        =   31
         Top             =   75
         Width           =   4440
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Room:"
         Height          =   240
         Left            =   75
         TabIndex        =   30
         Top             =   525
         Width           =   570
      End
   End
   Begin VB.PictureBox PTeach 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   120
      ScaleHeight     =   2310
      ScaleWidth      =   4560
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   4590
      Begin VB.TextBox TSubT 
         Height          =   360
         Left            =   975
         MaxLength       =   12
         TabIndex        =   26
         Top             =   900
         Width           =   3540
      End
      Begin VB.TextBox TTeacherT 
         Height          =   360
         Left            =   975
         MaxLength       =   70
         TabIndex        =   23
         Top             =   450
         Width           =   3540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Subject:"
         Height          =   240
         Left            =   75
         TabIndex        =   27
         Top             =   975
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Teacher:"
         Height          =   240
         Left            =   75
         TabIndex        =   25
         Top             =   525
         Width           =   780
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Teacher Record Manipulation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   75
         TabIndex        =   24
         Top             =   75
         Width           =   4440
      End
   End
   Begin VB.PictureBox PSubs 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   120
      ScaleHeight     =   2310
      ScaleWidth      =   4560
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   4590
      Begin VB.ComboBox CBYrSub 
         Height          =   360
         ItemData        =   "FrmAddNeed.frx":5832
         Left            =   3825
         List            =   "FrmAddNeed.frx":5845
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   825
         Width           =   690
      End
      Begin VB.TextBox TUnitsSub 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   360
         Left            =   1350
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1800
         Width           =   690
      End
      Begin VB.TextBox TDescSub 
         Height          =   360
         Left            =   1350
         TabIndex        =   3
         Top             =   1425
         Width           =   3165
      End
      Begin VB.TextBox TCourseSub 
         Height          =   360
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   1
         Top             =   825
         Width           =   2115
      End
      Begin VB.TextBox TScSub 
         Height          =   360
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   0
         Top             =   450
         Width           =   3165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Units:"
         Height          =   240
         Left            =   75
         TabIndex        =   13
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descriptive Title:"
         Height          =   240
         Left            =   75
         TabIndex        =   12
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Yr:"
         Height          =   240
         Left            =   3525
         TabIndex        =   11
         Top             =   825
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Course:"
         Height          =   240
         Left            =   75
         TabIndex        =   10
         Top             =   900
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subject Code:"
         Height          =   240
         Left            =   75
         TabIndex        =   9
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Subject Record Manipulation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   75
         TabIndex        =   8
         Top             =   75
         Width           =   4440
      End
   End
End
Attribute VB_Name = "FrmAddNeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Private Sub CBCancel_Click()
Select Case NeedActive
Case 1
SubjectQuery "Select * From Subjects order by sc"
Case 2
SectionQuery "Select * From Sections order by Sectionname"
Case 3
TeachersQuery "Select * From Teachers order by Teacher"
Case 4
RoomQuery "Select * From rooms order by room"
End Select
AddnewNeed = False
Unload Me
End Sub


Private Sub CBCoursesec_Click()
'FIXIT: Replace 'Trim' function with 'Trim$' function                                      FixIT90210ae-R9757-R1B8ZE
'FIXIT: Replace 'Trim' function with 'Trim$' function                                      FixIT90210ae-R9757-R1B8ZE
TSectionsec.Text = CBCoursesec.Text & " " & CBYrsec.Text & " " & CbxSec.Text
End Sub

Private Sub CBOk_Click()
Select Case NeedActive
    Case 1
        SubjectQueryx
    Case 2
        If Len(Trim(TSectionsec.Text)) = 0 Then Exit Sub
        SectionQueryx
    Case 3
        TeacherQueryx
    Case 4
        RoomQueryx
End Select
End Sub

Private Sub SubjectQueryx()
On Error GoTo ErrorSubjectQuery
Select Case AddnewNeed
Case True   'new record
    With RS_Subjects
        .AddNew
    End With
Case False
    'select the file
    With MainForm.LVSubjects
        Dim Sql As String
        Sql = "Select * from Subjects where Sc = " & Quoted(.SelectedItem.Text) & " and Course = " & Quoted(.SelectedItem.SubItems(1)) & " and Yr = " & Quoted(.SelectedItem.SubItems(2))
        SubjectQuery Sql
    End With
End Select
With RS_Subjects
'FIXIT: Replace 'Trim' function with 'Trim$' function                                      FixIT90210ae-R9757-R1B8ZE
    .Fields!sc = Trim(TScSub.Text)
'FIXIT: Replace 'Trim' function with 'Trim$' function                                      FixIT90210ae-R9757-R1B8ZE
    .Fields!course = Trim(TCourseSub.Text)
    .Fields!yr = CBYrSub.Text
'FIXIT: Replace 'Trim' function with 'Trim$' function                                      FixIT90210ae-R9757-R1B8ZE
    .Fields!Desc = Trim(TDescSub.Text)
'FIXIT: Replace 'Trim' function with 'Trim$' function                                      FixIT90210ae-R9757-R1B8ZE
    .Fields!Units = Trim(TUnitsSub.Text)
    .Update
    CBCancel_Click
End With
Exit Sub
ErrorSubjectQuery:
    MainModule_Exception
    RS_Subjects.Cancel
    RS_Subjects.CancelUpdate
    CBCancel_Click
End Sub

Private Sub SectionQueryx()
On Error GoTo ErrorSubjectQuery
Select Case AddnewNeed
    Case True   'new record
        With RS_Sections
            .AddNew
        End With
    Case False
        With MainForm.LVSection
            Dim Sql As String
            Sql = "Select * From Sections where SectionName = " & Quoted(.SelectedItem.Text)
            SectionQuery Sql
        End With
End Select
With RS_Sections
'FIXIT: Replace 'Trim' function with 'Trim$' function                              FixIT90210ae-R9757-R1B8ZE
    .Fields!Sectionname = Trim(TSectionsec.Text)
    .Update
    CBCancel_Click
End With
Exit Sub
ErrorSubjectQuery:
    MainModule_Exception
    RS_Sections.Cancel
    RS_Sections.CancelUpdate
    CBCancel_Click
End Sub

Private Sub TeacherQueryx()
On Error GoTo ErrorTeacherQuery
Select Case AddnewNeed
    Case True   'new record
        With RS_Teachers
            .AddNew
        End With
    Case False
        With MainForm.LVTeachers
        Dim Sql As String
            Sql = "Select * From Teachers where teacher = " & Quoted(.SelectedItem.Text) & _
                " and Subject = " & Quoted(.SelectedItem.SubItems(1))
            TeachersQuery Sql
        End With
End Select
With RS_Teachers
'FIXIT: Replace 'Trim' function with 'Trim$' function                                      FixIT90210ae-R9757-R1B8ZE
    .Fields!Teacher = Trim(TTeacherT.Text)
'FIXIT: Replace 'Trim' function with 'Trim$' function                                      FixIT90210ae-R9757-R1B8ZE
    .Fields!Subject = Trim(TSubT.Text)
    .Update
    CBCancel_Click
End With
Exit Sub
ErrorTeacherQuery:
    MainModule_Exception
    RS_Teachers.Cancel
    RS_Teachers.CancelUpdate
    CBCancel_Click
End Sub

Private Sub RoomQueryx()
On Error GoTo ErrorRoomQuery
Select Case AddnewNeed
    Case True   'new record
        With RS_Room
            .AddNew
        End With
    Case False
        With MainForm.LVRooms
        Dim Sql As String
            Sql = "Select * From rooms where room = " & Quoted(.SelectedItem.Text)
            RoomQuery Sql
        End With
End Select
With RS_Room
'FIXIT: Replace 'Trim' function with 'Trim$' function                                      FixIT90210ae-R9757-R1B8ZE
    .Fields!room = Trim(Troomr.Text)
    .Update
    CBCancel_Click
End With
Exit Sub
ErrorRoomQuery:
    MainModule_Exception
    RS_Room.Cancel
    RS_Room.CancelUpdate
    CBCancel_Click
End Sub


Private Sub CbxSec_Click()
TSectionsec.Text = CBCoursesec.Text & " " & CBYrsec.Text & " " & CbxSec.Text
End Sub

Private Sub CBYrsec_Click()
'FIXIT: Replace 'Trim' function with 'Trim$' function                                      FixIT90210ae-R9757-R1B8ZE
'FIXIT: Replace 'Trim' function with 'Trim$' function                                      FixIT90210ae-R9757-R1B8ZE
TSectionsec.Text = CBCoursesec.Text & " " & CBYrsec.Text & " " & CbxSec.Text
End Sub

Private Sub Form_Load()
Select Case NeedActive
    Case 1
        PSubs.Visible = True
    Case 2
        PSec.Visible = True
        LoadSectionstoCombo CBCoursesec
        'set the values for the cbxsec
        With CbxSec
            Dim i As Integer
            For i = 65 To 74
                .AddItem Chr(i)
            Next
        End With
    Case 3
        PTeach.Visible = True
    Case 4
        Proom.Visible = True
End Select
LoadTemp
End Sub

Private Sub LoadTemp()
If AddnewNeed = False Then
    Select Case NeedActive
        Case 1
            With MainForm.LVSubjects
                TScSub.Text = .SelectedItem.Text
                TCourseSub.Text = .SelectedItem.SubItems(1)
                CBYrSub.Text = .SelectedItem.SubItems(2)
                TDescSub.Text = .SelectedItem.SubItems(3)
                TUnitsSub.Text = .SelectedItem.SubItems(4)
            End With
        Case 2
            With MainForm.LVSection
                TSectionsec.Text = .SelectedItem.Text
            End With
        Case 3
            With MainForm.LVTeachers
                TTeacherT.Text = .SelectedItem.Text
                TSubT.Text = .SelectedItem.SubItems(1)
            End With
        Case 4
            With MainForm.LVRooms
                Troomr.Text = .SelectedItem.Text
            End With
    End Select
End If
End Sub
