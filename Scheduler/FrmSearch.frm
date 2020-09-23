VERSION 5.00
Begin VB.Form FrmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Records"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CBOk 
      Caption         =   "OK"
      Height          =   465
      Left            =   2175
      TabIndex        =   1
      Top             =   2550
      Width           =   1215
   End
   Begin VB.CommandButton CBCancel 
      Caption         =   "CANCEL"
      Height          =   465
      Left            =   3450
      TabIndex        =   0
      Top             =   2550
      Width           =   1215
   End
   Begin VB.PictureBox Pteach 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   75
      ScaleHeight     =   2310
      ScaleWidth      =   4560
      TabIndex        =   16
      Top             =   75
      Visible         =   0   'False
      Width           =   4590
      Begin VB.TextBox tsubt 
         Height          =   360
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   18
         Top             =   825
         Width           =   3165
      End
      Begin VB.TextBox tteachert 
         Height          =   360
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   17
         Top             =   450
         Width           =   3165
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Subject:"
         Height          =   240
         Left            =   75
         TabIndex        =   22
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Teacher:"
         Height          =   240
         Left            =   75
         TabIndex        =   21
         Top             =   525
         Width           =   780
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Search Teacher"
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
         TabIndex        =   20
         Top             =   75
         Width           =   4440
      End
      Begin VB.Label Label9 
         Caption         =   "NOTE:To Select all, Leave all as blank. Use % for wildcard search of characters and # for numbers."
         Height          =   840
         Left            =   150
         TabIndex        =   19
         Top             =   1350
         Width           =   4290
      End
   End
   Begin VB.PictureBox PSec 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   75
      ScaleHeight     =   2310
      ScaleWidth      =   4560
      TabIndex        =   11
      Top             =   75
      Visible         =   0   'False
      Width           =   4590
      Begin VB.TextBox Tsectionsec 
         Height          =   360
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   12
         Top             =   450
         Width           =   3165
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Section Name:"
         Height          =   240
         Left            =   75
         TabIndex        =   15
         Top             =   525
         Width           =   1260
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Search Section"
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
         TabIndex        =   14
         Top             =   75
         Width           =   4440
      End
      Begin VB.Label Label6 
         Caption         =   "NOTE:To Select all, Leave all as blank. Use % for wildcard search of characters and # for numbers."
         Height          =   840
         Left            =   150
         TabIndex        =   13
         Top             =   1350
         Width           =   4290
      End
   End
   Begin VB.PictureBox PSubs 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   75
      ScaleHeight     =   2310
      ScaleWidth      =   4560
      TabIndex        =   2
      Top             =   75
      Visible         =   0   'False
      Width           =   4590
      Begin VB.TextBox TScSub 
         Height          =   360
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   5
         Top             =   450
         Width           =   3165
      End
      Begin VB.TextBox TCourseSub 
         Height          =   360
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   4
         Top             =   825
         Width           =   2115
      End
      Begin VB.ComboBox CBYrSub 
         Height          =   360
         ItemData        =   "FrmSearch.frx":6852
         Left            =   3825
         List            =   "FrmSearch.frx":6865
         TabIndex        =   3
         Top             =   825
         Width           =   690
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Search Subject"
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
         TabIndex        =   9
         Top             =   75
         Width           =   4440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subject Code:"
         Height          =   240
         Left            =   75
         TabIndex        =   8
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Course:"
         Height          =   240
         Left            =   75
         TabIndex        =   7
         Top             =   900
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Yr:"
         Height          =   240
         Left            =   3525
         TabIndex        =   6
         Top             =   825
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "NOTE:To Select all, Leave all as blank. Use % for wildcard search of characters and # for numbers."
         Height          =   840
         Left            =   150
         TabIndex        =   10
         Top             =   1350
         Width           =   4290
      End
   End
   Begin VB.PictureBox Proom 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   75
      ScaleHeight     =   2310
      ScaleWidth      =   4560
      TabIndex        =   23
      Top             =   75
      Visible         =   0   'False
      Width           =   4590
      Begin VB.TextBox Troomr 
         Height          =   360
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   24
         Top             =   450
         Width           =   3165
      End
      Begin VB.Label Label16 
         Caption         =   "NOTE:To Select all, Leave all as blank. Use % for wildcard search of characters and # for numbers."
         Height          =   840
         Left            =   150
         TabIndex        =   27
         Top             =   1350
         Width           =   4290
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Search Room"
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
         TabIndex        =   26
         Top             =   75
         Width           =   4440
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Room:"
         Height          =   240
         Left            =   75
         TabIndex        =   25
         Top             =   525
         Width           =   570
      End
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBCancel_Click()
Select Case NeedActive
Case 1
SubjectQuery "Select * From Subjects order by sc"
Case 2
SectionQuery "Select * From Sections order by Sectionname"
Case 3
TeachersQuery "Select * From Teachers order by Teacher"
Case 4
RoomQuery "Select * From Rooms order by room"
End Select
Unload Me
End Sub

Private Sub CBOk_Click()
Select Case NeedActive
Case 1
    SubjectQBuilder
Case 2
    SectionQBuilder
Case 3
    TeacherBuilder
Case 4
    RoomQBuilder
End Select
End Sub

Private Sub SubjectQBuilder()
Dim Sql As String
Sql = "Select * From Subjects where "
If Len(Trim(TScSub.Text)) <> 0 Then
    Sql = Sql & "sc like " & Quoted(Trim(TScSub.Text)) & " "
End If
If Len(Trim(TCourseSub.Text)) <> 0 Then
    Sql = Sql & "Course like " & Quoted(Trim(TCourseSub.Text)) & " "
End If
If Len(Trim(CBYrSub.Text)) <> 0 Then
    Sql = Sql & "Yr like " & Quoted(Trim(CBYrSub.Text)) & " "
End If
If Len(Sql) < 30 Then
    Sql = "Select * From Subjects"
End If
SubjectQuery Sql
Unload Me
End Sub

Private Sub SectionQBuilder()
Dim Sql As String
Sql = "Select * From Sections where "
If Len(Trim(Tsectionsec.Text)) <> 0 Then
    Sql = Sql & "SectionName like " & Quoted(Trim(Tsectionsec.Text))
End If
If Len(Sql) < 30 Then
    Sql = "Select * From Sections"
End If
SectionQuery Sql
Unload Me
End Sub

Private Sub TeacherBuilder()
Dim Sql As String
Sql = "Select * From Teachers where "
If Len(Trim(tteachert.Text)) <> 0 Then
    Sql = Sql & "teacher like " & Quoted(Trim(tteachert.Text))
End If
If Len(Trim(tsubt.Text)) <> 0 Then
    Sql = Sql & "Subject like " & Quoted(Trim(tsubt.Text))
End If
If Len(Sql) < 30 Then
    Sql = "Select * From Teachers"
End If
TeachersQuery Sql
Unload Me
End Sub

Private Sub RoomQBuilder()
Dim Sql As String
Sql = "Select * From Rooms where "
If Len(Trim(Troomr.Text)) <> 0 Then
    Sql = Sql & "room like " & Quoted(Trim(Troomr.Text))
End If
If Len(Sql) < 27 Then
    Sql = "Select * From Rooms"
End If
RoomQuery Sql
Unload Me
End Sub

Private Sub Form_Load()
Select Case NeedActive
    Case 1
        PSubs.Visible = True
    Case 2
        PSec.Visible = True
    Case 3
        Pteach.Visible = True
    Case 4
        Proom.Visible = True
End Select
End Sub
