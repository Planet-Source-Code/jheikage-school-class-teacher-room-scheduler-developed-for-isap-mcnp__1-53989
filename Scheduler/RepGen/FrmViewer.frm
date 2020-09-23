VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Setup"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   5145
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmRoom 
      Caption         =   "Room Report Setup"
      Height          =   5055
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton BtnCrte 
         Caption         =   "&Create Report"
         Height          =   495
         Left            =   6960
         TabIndex        =   24
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton BtnPrev 
         Caption         =   "&Preview"
         Height          =   495
         Left            =   5280
         TabIndex        =   23
         Top             =   4440
         Width           =   1575
      End
      Begin VB.ListBox LstRoom 
         Height          =   3900
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton Btnsrc3 
         Caption         =   "..."
         Height          =   360
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "Set Data Source"
         Top             =   240
         Width           =   375
      End
      Begin MSComctlLib.ListView LVRooms 
         Height          =   3375
         Left            =   3120
         TabIndex        =   25
         Top             =   960
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "DAY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SUBJECT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DESCRIPTION"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "UNITS"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "TIME"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "PROFESSOR/INSTRUCTOR"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Class"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   1080
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Select Project"
         Filter          =   "Scheduling Project (*.mdb)|*.mdb"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Schedules"
         Height          =   240
         Index           =   8
         Left            =   3120
         TabIndex        =   29
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Room List"
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   885
      End
      Begin VB.Label LblSoruce 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   240
         Left            =   1920
         TabIndex        =   27
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Source:"
         Height          =   240
         Index           =   6
         Left            =   720
         TabIndex        =   26
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame FrmTeacher 
      Caption         =   "Teacher Report Setup"
      Height          =   5055
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton BtnSetDS 
         Caption         =   "..."
         Height          =   360
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Set Data Source"
         Top             =   240
         Width           =   375
      End
      Begin VB.ListBox LstTeacher 
         Height          =   3900
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton BtnPreview 
         Caption         =   "&Preview"
         Height          =   495
         Left            =   5280
         TabIndex        =   12
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton BtnTCreate 
         Caption         =   "&Create Report"
         Height          =   495
         Left            =   6960
         TabIndex        =   11
         Top             =   4440
         Width           =   1575
      End
      Begin MSComctlLib.ListView LVTSched 
         Height          =   3375
         Left            =   3120
         TabIndex        =   13
         Top             =   960
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "DAY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SUBJECT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DESCRIPTION"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "UNITS"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "TIME"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "CLASS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ROOM"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1080
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Select Project"
         Filter          =   "Scheduling Project (*.mdb)|*.mdb"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Source:"
         Height          =   240
         Index           =   5
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   240
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Teacher's List"
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Schedules"
         Height          =   240
         Index           =   3
         Left            =   3120
         TabIndex        =   16
         Top             =   720
         Width           =   915
      End
   End
   Begin VB.Frame FrmReps 
      Caption         =   "Class Report Setup"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton BtnCreate 
         Caption         =   "&Create Report"
         Height          =   495
         Left            =   6960
         TabIndex        =   9
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton btnView 
         Caption         =   "&Preview"
         Height          =   495
         Left            =   5280
         TabIndex        =   8
         Top             =   4440
         Width           =   1575
      End
      Begin MSComctlLib.ListView LVScheds 
         Height          =   3375
         Left            =   3120
         TabIndex        =   6
         Top             =   960
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "DAY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SUBJECT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DESCRIPTION"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "UNITS"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "TIME"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "PROFESSOR/INSTRUCTOR"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "REMARKS"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComDlg.CommonDialog DLG 
         Left            =   1080
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Select Project"
         Filter          =   "Scheduling Project (*.mdb)|*.mdb"
      End
      Begin VB.ListBox LstClass 
         Height          =   3900
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton btndsset 
         Caption         =   "..."
         Height          =   360
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Set Data Source"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Schedules"
         Height          =   240
         Index           =   2
         Left            =   3120
         TabIndex        =   7
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Class List"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   870
      End
      Begin VB.Label LblSource 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   240
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Source:"
         Height          =   240
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "FrmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function connect(Source As String) As Boolean
On Error GoTo ERB
    With Mycon
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .ConnectionString = "provider = MSDATASHAPE.1;data provider=microsoft.jet.oledb.4.0;persist security info=false;Jet Oledb:Database password=vip;data source=" & Source
        .Open
    End With
    connect = True
Exit Function
ERB:
MsgBox Err.Description, vbCritical, "Error"
connect = False
If Mycon.State <> 0 Then Mycon.Close
End Function

Function PopulateList(Lb As ListBox, Sql As String, Field As String)
    On Error GoTo ERB
    Lb.Clear
    Dim rs As New Recordset
    With rs
    .CursorLocation = adUseClient
    .Open Sql, Mycon, adOpenDynamic, adLockOptimistic
    If .RecordCount <> 0 Then
        Do Until .EOF
            Lb.AddItem .Fields(0).Value
            .MoveNext
        Loop
    End If
    End With
    Set rs = Nothing
    Exit Function
ERB:
    MsgBox Err.Description, vbCritical, "Error"
    Set rs = Nothing
    Lb.AddItem "Error occured. Loading Failed."
    Mycon.Close
    
End Function

Private Sub BtnCreate_Click()
If Mycon.State <> 0 Then
Dim Report As New CRepClass
With FrmPrint
    btnView_Click
    If DE.CON.State <> 0 Then DE.CON.Close
    DE.CON.Open
    
    .Vrpc.ReportSource = Report
    .Vrpc.Refresh
    .Vrpc.ViewReport
End With
FrmPrint.Show 1
Else
    MsgBox "Please select a data source first.", vbInformation, "Error"
End If
End Sub

Private Sub BtnCrte_Click()
If Mycon.State <> 0 Then
Dim Report As New CrepRooms
With FrmPrint
    BtnPrev_Click
    If DE.CON.State <> 0 Then DE.CON.Close
    DE.CON.Open
    .Vrpc.ReportSource = Report
    .Vrpc.Refresh
    .Vrpc.ViewReport
End With
FrmPrint.Show 1
Else
    MsgBox "Please select a data source first.", vbInformation, "Error"
End If

End Sub

Private Sub btndsset_Click()
DLG.ShowOpen
If Trim(DLG.FileName) = "" Then Exit Sub
LblSource.Caption = DLG.FileName
If connect(LblSource.Caption) = True Then
    PopulateList Me.LstClass, "Select * from sections", "sectionname"
Else
    If Mycon.State <> 0 Then Mycon.Close
End If
End Sub

Private Sub BtnPrev_Click()
'load the data
If Mycon.State <> 0 Then
    If LstRoom.ListIndex < 0 Then LstRoom.ListIndex = 0
        roomsched LstRoom.Text
        RoomLoad LVRooms
Else
    MsgBox "Please select a data source first.", vbInformation, "Error"
End If
End Sub

Private Sub BtnPreview_Click()
'load the data
If Mycon.State <> 0 Then
    If LstTeacher.ListIndex < 0 Then LstTeacher.ListIndex = 0
        TeacherScheds LstTeacher.Text
        TeacherLoad LVTSched
Else
    MsgBox "Please select a data source first.", vbInformation, "Error"
End If
End Sub

Private Sub BtnSetDS_Click()
DLG.ShowOpen
If Trim(DLG.FileName) = "" Then Exit Sub
LblSource.Caption = DLG.FileName
If connect(LblSource.Caption) = True Then
    PopulateList Me.LstTeacher, "Select Teacher from Teachers group by teacher", "Teacher"
Else
    If Mycon.State <> 0 Then Mycon.Close
End If
End Sub

Private Sub Btnsrc3_Click()
DLG.ShowOpen
If Trim(DLG.FileName) = "" Then Exit Sub
LblSoruce.Caption = DLG.FileName
If connect(Me.LblSoruce.Caption) = True Then
    PopulateList Me.LstRoom, "Select * from rooms", "room"
Else
    If Mycon.State <> 0 Then Mycon.Close
End If
End Sub

Private Sub BtnTCreate_Click()
If Mycon.State <> 0 Then
Dim Report As New CrepTeacher
With FrmPrint
    BtnPreview_Click
    If DE.CON.State <> 0 Then DE.CON.Close
    DE.CON.Open
    .Vrpc.ReportSource = Report
    .Vrpc.Refresh
    .Vrpc.ViewReport
End With
FrmPrint.Show 1
Else
    MsgBox "Please select a data source first.", vbInformation, "Error"
End If

End Sub

Private Sub btnView_Click()
'load the data
If Mycon.State <> 0 Then
    If LstClass.ListIndex < 0 Then LstClass.ListIndex = 0
        SchedHandler LstClass.Text
        classLoad LVScheds
Else
    MsgBox "Please select a data source first.", vbInformation, "Error"
End If
End Sub
