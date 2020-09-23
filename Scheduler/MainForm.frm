VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form MainForm 
   Caption         =   "ISAP MCNP SCHEDULER"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CmdNew 
      Left            =   7125
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "New Project"
      Filter          =   "Access Files(*.MDB)|*.mdb"
   End
   Begin MSComctlLib.ImageList Subjects 
      Left            =   7125
      Top             =   1725
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":23D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":7BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":E42A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":146C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CMDOpen 
      Left            =   7125
      Top             =   2175
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Schedule"
      Filter          =   "Access Files(*.MDB)|*.mdb"
   End
   Begin VB.PictureBox FrmProp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   7680
      ScaleHeight     =   7425
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton mpr 
         BackColor       =   &H000000FF&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2250
         MaskColor       =   &H000000C0&
         Style           =   1  'Graphical
         TabIndex        =   259
         Top             =   1350
         Width           =   240
      End
      Begin VB.CommandButton mbc 
         BackColor       =   &H000000FF&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2025
         MaskColor       =   &H000000C0&
         Style           =   1  'Graphical
         TabIndex        =   258
         Top             =   1350
         Width           =   240
      End
      Begin VB.ListBox LSubs 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5685
         ItemData        =   "MainForm.frx":19EBA
         Left            =   0
         List            =   "MainForm.frx":19EC7
         TabIndex        =   11
         Top             =   1575
         Width           =   3975
      End
      Begin VB.CommandButton CBprjUpdate 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1725
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   900
         Width           =   765
      End
      Begin VB.TextBox TSem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1050
         MaxLength       =   1
         TabIndex        =   7
         Top             =   675
         Width           =   180
      End
      Begin VB.TextBox TSy 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1050
         MaxLength       =   10
         TabIndex        =   5
         Top             =   465
         Width           =   1455
      End
      Begin VB.TextBox TPrjName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Subject List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   1350
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2550
         Y1              =   1275
         Y2              =   1275
      End
      Begin VB.Label LblCaps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1275
         TabIndex        =   8
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semester"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   6
         Top             =   675
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   4
         Top             =   450
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   2
         Top             =   225
         Width           =   960
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Project Properties"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2535
      End
   End
   Begin TabDlg.SSTab STAB 
      Height          =   7215
      Left            =   75
      TabIndex        =   13
      Top             =   300
      Visible         =   0   'False
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   12726
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   529
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "SUBJECTS"
      TabPicture(0)   =   "MainForm.frx":19F0E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrmSub"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LVSubjects"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "SECTIONS"
      TabPicture(1)   =   "MainForm.frx":19F2A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frmsec"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LVSection"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "TEACHERS"
      TabPicture(2)   =   "MainForm.frx":19F46
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrmTeach"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "LVTeachers"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "ROOMS"
      TabPicture(3)   =   "MainForm.frx":19F62
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrmR"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "LVRooms"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "SCHEDULES"
      TabPicture(4)   =   "MainForm.frx":19F7E
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Pback"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin MSComctlLib.ListView LVSubjects 
         Height          =   5940
         Left            =   -74925
         TabIndex        =   14
         Top             =   450
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   10478
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "Subjects"
         SmallIcons      =   "Subjects"
         ColHdrIcons     =   "Subjects"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SUBJECT CODE"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "COURSE"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "YR"
            Object.Width           =   988
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DESCRIPTIVE TITLE"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "UNITS"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Frame FrmSub 
         Height          =   840
         Left            =   -74925
         TabIndex        =   15
         Top             =   6300
         Width           =   7365
         Begin VB.CommandButton CBPrintSubs 
            Caption         =   "&Print"
            Height          =   540
            Left            =   5625
            Style           =   1  'Graphical
            TabIndex        =   261
            Top             =   225
            Width           =   1290
         End
         Begin VB.CommandButton BdelSub 
            Caption         =   "&Delete"
            Height          =   540
            Left            =   2925
            Style           =   1  'Graphical
            TabIndex        =   260
            Top             =   225
            Width           =   1290
         End
         Begin VB.CommandButton BNewSub 
            Caption         =   "&New"
            Height          =   540
            Left            =   75
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   225
            Width           =   1365
         End
         Begin VB.CommandButton BsearchSub 
            Caption         =   "&Search"
            Height          =   540
            Left            =   4275
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   225
            Width           =   1290
         End
         Begin VB.CommandButton BeditSub 
            Caption         =   "&Edit"
            Height          =   540
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   225
            Width           =   1365
         End
      End
      Begin MSComctlLib.ListView LVSection 
         Height          =   5940
         Left            =   -74925
         TabIndex        =   19
         Top             =   450
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   10478
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "Subjects"
         SmallIcons      =   "Subjects"
         ColHdrIcons     =   "Subjects"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SECTIONS"
            Object.Width           =   4410
            ImageIndex      =   2
         EndProperty
      End
      Begin VB.Frame Frmsec 
         Height          =   840
         Left            =   -74925
         TabIndex        =   20
         Top             =   6300
         Width           =   7365
         Begin VB.CommandButton CBPrintSection 
            Caption         =   "&Print"
            Height          =   540
            Left            =   5625
            Style           =   1  'Graphical
            TabIndex        =   262
            Top             =   225
            Width           =   1290
         End
         Begin VB.CommandButton CBEditSec 
            Caption         =   "&Edit"
            Height          =   540
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   225
            Width           =   1365
         End
         Begin VB.CommandButton CBDelSec 
            Caption         =   "&Delete"
            Height          =   540
            Left            =   2925
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   225
            Width           =   1290
         End
         Begin VB.CommandButton CBSearchSec 
            Caption         =   "&Search"
            Height          =   540
            Left            =   4275
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   225
            Width           =   1290
         End
         Begin VB.CommandButton CBNewSec 
            Caption         =   "&New"
            Height          =   540
            Left            =   75
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   225
            Width           =   1365
         End
      End
      Begin MSComctlLib.ListView LVTeachers 
         Height          =   5940
         Left            =   -74925
         TabIndex        =   25
         Top             =   450
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   10478
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "Subjects"
         SmallIcons      =   "Subjects"
         ColHdrIcons     =   "Subjects"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Teacher"
            Object.Width           =   5292
            ImageIndex      =   3
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Subject"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame FrmTeach 
         Height          =   840
         Left            =   -74925
         TabIndex        =   26
         Top             =   6300
         Width           =   7365
         Begin VB.CommandButton CBPrintTeacher 
            Caption         =   "&Print"
            Height          =   540
            Left            =   5625
            Style           =   1  'Graphical
            TabIndex        =   263
            Top             =   225
            Width           =   1290
         End
         Begin VB.CommandButton CBNewT 
            Caption         =   "&New"
            Height          =   540
            Left            =   75
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   225
            Width           =   1365
         End
         Begin VB.CommandButton CBSearchT 
            Caption         =   "&Search"
            Height          =   540
            Left            =   4275
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   225
            Width           =   1290
         End
         Begin VB.CommandButton CBDeleteT 
            Caption         =   "&Delete"
            Height          =   540
            Left            =   2925
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   225
            Width           =   1290
         End
         Begin VB.CommandButton CBEditT 
            Caption         =   "&Edit"
            Height          =   540
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   225
            Width           =   1365
         End
      End
      Begin MSComctlLib.ListView LVRooms 
         Height          =   5940
         Left            =   -74925
         TabIndex        =   31
         Top             =   450
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   10478
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "Subjects"
         SmallIcons      =   "Subjects"
         ColHdrIcons     =   "Subjects"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Rooms"
            Object.Width           =   4410
            ImageIndex      =   4
         EndProperty
      End
      Begin VB.Frame FrmR 
         Height          =   840
         Left            =   -74925
         TabIndex        =   32
         Top             =   6300
         Width           =   7365
         Begin VB.CommandButton CbPrintRoom 
            Caption         =   "&Print"
            Height          =   540
            Left            =   5625
            Style           =   1  'Graphical
            TabIndex        =   264
            Top             =   225
            Width           =   1290
         End
         Begin VB.CommandButton CBNewR 
            Caption         =   "&New"
            Height          =   540
            Left            =   75
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   225
            Width           =   1365
         End
         Begin VB.CommandButton CBSearchR 
            Caption         =   "&Search"
            Height          =   540
            Left            =   4275
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   225
            Width           =   1290
         End
         Begin VB.CommandButton CBDeleteR 
            Caption         =   "&Delete"
            Height          =   540
            Left            =   2925
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   225
            Width           =   1290
         End
         Begin VB.CommandButton CBEditR 
            Caption         =   "&Edit"
            Height          =   540
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   225
            Width           =   1365
         End
      End
      Begin VB.Frame Pback 
         Height          =   6765
         Left            =   75
         TabIndex        =   37
         Top             =   375
         Width           =   7365
         Begin VB.CommandButton cbCreateReport 
            Caption         =   "&PRINT"
            Height          =   390
            Left            =   6450
            TabIndex        =   266
            Top             =   300
            Width           =   765
         End
         Begin VB.CommandButton CBTeacherLoads 
            Caption         =   "&LOADS"
            Height          =   390
            Left            =   5625
            TabIndex        =   265
            Top             =   300
            Width           =   765
         End
         Begin VB.ComboBox TOUT 
            Height          =   360
            ItemData        =   "MainForm.frx":19F9A
            Left            =   675
            List            =   "MainForm.frx":19FF5
            Style           =   2  'Dropdown List
            TabIndex        =   256
            Top             =   0
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.ComboBox CBView 
            Height          =   360
            ItemData        =   "MainForm.frx":1A104
            Left            =   1200
            List            =   "MainForm.frx":1A111
            Style           =   2  'Dropdown List
            TabIndex        =   137
            Top             =   300
            Width           =   1365
         End
         Begin VB.ComboBox CBSelect 
            Height          =   360
            ItemData        =   "MainForm.frx":1A12B
            Left            =   3900
            List            =   "MainForm.frx":1A132
            TabIndex        =   136
            Text            =   "SELECT DATA"
            Top             =   300
            Width           =   1665
         End
         Begin VB.PictureBox PCon1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   5790
            Left            =   75
            ScaleHeight     =   5760
            ScaleWidth      =   7110
            TabIndex        =   38
            Top             =   750
            Width           =   7140
            Begin VB.HScrollBar Hscrlclass 
               Height          =   240
               LargeChange     =   5
               Left            =   0
               Max             =   48
               TabIndex        =   39
               Top             =   5550
               Width           =   7140
            End
            Begin VB.PictureBox pxx 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   6840
               Left            =   0
               ScaleHeight     =   6810
               ScaleWidth      =   735
               TabIndex        =   106
               Top             =   0
               Width           =   765
               Begin VB.Label Label40 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "9:00 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   135
                  Top             =   6600
                  Width           =   765
               End
               Begin VB.Label Label39 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "8:30 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   134
                  Top             =   6375
                  Width           =   765
               End
               Begin VB.Label Label38 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "8:00 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   133
                  Top             =   6150
                  Width           =   765
               End
               Begin VB.Label Label37 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "7:30 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   132
                  Top             =   5925
                  Width           =   765
               End
               Begin VB.Label Label36 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "7:00 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   131
                  Top             =   5700
                  Width           =   765
               End
               Begin VB.Label Label35 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "6:30 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   130
                  Top             =   5475
                  Width           =   765
               End
               Begin VB.Label Label34 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "6:00 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   129
                  Top             =   5250
                  Width           =   765
               End
               Begin VB.Label Label33 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "5:30 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   128
                  Top             =   5025
                  Width           =   765
               End
               Begin VB.Label Label32 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "5:00 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   127
                  Top             =   4800
                  Width           =   765
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "4:30 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   126
                  Top             =   4575
                  Width           =   765
               End
               Begin VB.Label Label30 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "4:00 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   125
                  Top             =   4350
                  Width           =   765
               End
               Begin VB.Label Label29 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "3:30 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   124
                  Top             =   4125
                  Width           =   765
               End
               Begin VB.Label Label28 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "3:00 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   123
                  Top             =   3900
                  Width           =   765
               End
               Begin VB.Label Label27 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "2:30 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   122
                  Top             =   3675
                  Width           =   765
               End
               Begin VB.Label Label26 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "2:00 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   121
                  Top             =   3450
                  Width           =   765
               End
               Begin VB.Label Label25 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "1:30 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   120
                  Top             =   3225
                  Width           =   765
               End
               Begin VB.Label Label24 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "1:00 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   119
                  Top             =   3000
                  Width           =   765
               End
               Begin VB.Label Label23 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "12:30 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   118
                  Top             =   2775
                  Width           =   765
               End
               Begin VB.Label Label22 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "12:00 PM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   117
                  Top             =   2550
                  Width           =   765
               End
               Begin VB.Label Label21 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "11:30 AM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   116
                  Top             =   2325
                  Width           =   765
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "11:00 AM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   115
                  Top             =   2100
                  Width           =   765
               End
               Begin VB.Label Label19 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "10:30 AM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   114
                  Top             =   1875
                  Width           =   765
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "10:00 AM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   113
                  Top             =   1650
                  Width           =   765
               End
               Begin VB.Label Label17 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "9:30 AM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   112
                  Top             =   1425
                  Width           =   765
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "9:00 AM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   111
                  Top             =   1200
                  Width           =   765
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "8:30 AM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   110
                  Top             =   975
                  Width           =   765
               End
               Begin VB.Label Label14 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "8:00 AM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   109
                  Top             =   750
                  Width           =   765
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "7:30 AM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   108
                  Top             =   525
                  Width           =   765
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "7:00 AM"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   107
                  Top             =   300
                  Width           =   765
               End
            End
            Begin VB.VScrollBar VScrclass 
               Height          =   5565
               LargeChange     =   5
               Left            =   6900
               Max             =   51
               TabIndex        =   40
               Top             =   0
               Width           =   240
            End
            Begin VB.PictureBox PClass 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   6840
               Left            =   750
               ScaleHeight     =   6810
               ScaleWidth      =   8085
               TabIndex        =   41
               Top             =   0
               Width           =   8115
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   0
                  Left            =   6750
                  TabIndex        =   227
                  Top             =   300
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   1
                  Left            =   6750
                  TabIndex        =   228
                  Top             =   525
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   2
                  Left            =   6750
                  TabIndex        =   229
                  Top             =   750
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   3
                  Left            =   6750
                  TabIndex        =   230
                  Top             =   975
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   4
                  Left            =   6750
                  TabIndex        =   231
                  Top             =   1200
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   5
                  Left            =   6750
                  TabIndex        =   232
                  Top             =   1425
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   6
                  Left            =   6750
                  TabIndex        =   233
                  Top             =   1650
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   7
                  Left            =   6750
                  TabIndex        =   234
                  Top             =   1875
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   8
                  Left            =   6750
                  TabIndex        =   235
                  Top             =   2100
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   9
                  Left            =   6750
                  TabIndex        =   236
                  Top             =   2325
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   10
                  Left            =   6750
                  TabIndex        =   237
                  Top             =   2550
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   11
                  Left            =   6750
                  TabIndex        =   238
                  Top             =   2775
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   12
                  Left            =   6750
                  TabIndex        =   239
                  Top             =   3000
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   13
                  Left            =   6750
                  TabIndex        =   240
                  Top             =   3225
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   14
                  Left            =   6750
                  TabIndex        =   241
                  Top             =   3450
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   15
                  Left            =   6750
                  TabIndex        =   242
                  Top             =   3675
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   16
                  Left            =   6750
                  TabIndex        =   243
                  Top             =   3900
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   17
                  Left            =   6750
                  TabIndex        =   244
                  Top             =   4125
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   18
                  Left            =   6750
                  TabIndex        =   245
                  Top             =   4350
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   19
                  Left            =   6750
                  TabIndex        =   246
                  Top             =   4575
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   20
                  Left            =   6750
                  TabIndex        =   247
                  Top             =   4800
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   21
                  Left            =   6750
                  TabIndex        =   248
                  Top             =   5025
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   22
                  Left            =   6750
                  TabIndex        =   249
                  Top             =   5250
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   23
                  Left            =   6750
                  TabIndex        =   250
                  Top             =   5475
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   24
                  Left            =   6750
                  TabIndex        =   251
                  Top             =   5700
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   25
                  Left            =   6750
                  TabIndex        =   252
                  Top             =   5925
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   26
                  Left            =   6750
                  TabIndex        =   253
                  Top             =   6150
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   27
                  Left            =   6750
                  TabIndex        =   254
                  Top             =   6375
                  Width           =   1290
               End
               Begin VB.Label SAT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   28
                  Left            =   6750
                  TabIndex        =   255
                  Top             =   6600
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   0
                  Left            =   4050
                  TabIndex        =   169
                  Top             =   300
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   1
                  Left            =   4050
                  TabIndex        =   170
                  Top             =   525
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   2
                  Left            =   4050
                  TabIndex        =   171
                  Top             =   750
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   3
                  Left            =   4050
                  TabIndex        =   172
                  Top             =   975
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   4
                  Left            =   4050
                  TabIndex        =   173
                  Top             =   1200
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   5
                  Left            =   4050
                  TabIndex        =   174
                  Top             =   1425
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   6
                  Left            =   4050
                  TabIndex        =   175
                  Top             =   1650
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   7
                  Left            =   4050
                  TabIndex        =   176
                  Top             =   1875
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   8
                  Left            =   4050
                  TabIndex        =   177
                  Top             =   2100
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   9
                  Left            =   4050
                  TabIndex        =   178
                  Top             =   2325
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   10
                  Left            =   4050
                  TabIndex        =   179
                  Top             =   2550
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   11
                  Left            =   4050
                  TabIndex        =   180
                  Top             =   2775
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   12
                  Left            =   4050
                  TabIndex        =   181
                  Top             =   3000
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   13
                  Left            =   4050
                  TabIndex        =   182
                  Top             =   3225
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   14
                  Left            =   4050
                  TabIndex        =   183
                  Top             =   3450
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   15
                  Left            =   4050
                  TabIndex        =   184
                  Top             =   3675
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   16
                  Left            =   4050
                  TabIndex        =   185
                  Top             =   3900
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   17
                  Left            =   4050
                  TabIndex        =   186
                  Top             =   4125
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   18
                  Left            =   4050
                  TabIndex        =   187
                  Top             =   4350
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   19
                  Left            =   4050
                  TabIndex        =   188
                  Top             =   4575
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   20
                  Left            =   4050
                  TabIndex        =   189
                  Top             =   4800
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   21
                  Left            =   4050
                  TabIndex        =   190
                  Top             =   5025
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   22
                  Left            =   4050
                  TabIndex        =   191
                  Top             =   5250
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   23
                  Left            =   4050
                  TabIndex        =   192
                  Top             =   5475
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   24
                  Left            =   4050
                  TabIndex        =   193
                  Top             =   5700
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   25
                  Left            =   4050
                  TabIndex        =   194
                  Top             =   5925
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   26
                  Left            =   4050
                  TabIndex        =   195
                  Top             =   6150
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   27
                  Left            =   4050
                  TabIndex        =   196
                  Top             =   6375
                  Width           =   1290
               End
               Begin VB.Label TH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   28
                  Left            =   4050
                  TabIndex        =   197
                  Top             =   6600
                  Width           =   1290
               End
               Begin VB.Label Label6 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "MONDAY"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   0
                  TabIndex        =   105
                  Top             =   0
                  Width           =   1290
               End
               Begin VB.Label Label7 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "TUESDAY"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   1350
                  TabIndex        =   104
                  Top             =   0
                  Width           =   1290
               End
               Begin VB.Label Label8 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "WEDNESDAY"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   2700
                  TabIndex        =   103
                  Top             =   0
                  Width           =   1290
               End
               Begin VB.Label Label9 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "THURSDAY"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   4050
                  TabIndex        =   102
                  Top             =   0
                  Width           =   1290
               End
               Begin VB.Label Label10 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "FRIDAY"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   5400
                  TabIndex        =   101
                  Top             =   0
                  Width           =   1290
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "SATURDAY"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   6750
                  TabIndex        =   100
                  Top             =   0
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   0
                  Left            =   1350
                  TabIndex        =   70
                  Top             =   300
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   1
                  Left            =   1350
                  TabIndex        =   69
                  Top             =   525
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   2
                  Left            =   1350
                  TabIndex        =   68
                  Top             =   750
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   3
                  Left            =   1350
                  TabIndex        =   67
                  Top             =   975
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   4
                  Left            =   1350
                  TabIndex        =   66
                  Top             =   1200
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   5
                  Left            =   1350
                  TabIndex        =   65
                  Top             =   1425
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   6
                  Left            =   1350
                  TabIndex        =   64
                  Top             =   1650
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   7
                  Left            =   1350
                  TabIndex        =   63
                  Top             =   1875
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   8
                  Left            =   1350
                  TabIndex        =   62
                  Top             =   2100
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   9
                  Left            =   1350
                  TabIndex        =   61
                  Top             =   2325
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   10
                  Left            =   1350
                  TabIndex        =   60
                  Top             =   2550
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   11
                  Left            =   1350
                  TabIndex        =   59
                  Top             =   2775
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   12
                  Left            =   1350
                  TabIndex        =   58
                  Top             =   3000
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   13
                  Left            =   1350
                  TabIndex        =   57
                  Top             =   3225
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   14
                  Left            =   1350
                  TabIndex        =   56
                  Top             =   3450
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   15
                  Left            =   1350
                  TabIndex        =   55
                  Top             =   3675
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   16
                  Left            =   1350
                  TabIndex        =   54
                  Top             =   3900
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   17
                  Left            =   1350
                  TabIndex        =   53
                  Top             =   4125
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   18
                  Left            =   1350
                  TabIndex        =   52
                  Top             =   4350
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   19
                  Left            =   1350
                  TabIndex        =   51
                  Top             =   4575
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   20
                  Left            =   1350
                  TabIndex        =   50
                  Top             =   4800
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   21
                  Left            =   1350
                  TabIndex        =   49
                  Top             =   5025
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   22
                  Left            =   1350
                  TabIndex        =   48
                  Top             =   5250
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   23
                  Left            =   1350
                  TabIndex        =   47
                  Top             =   5475
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   24
                  Left            =   1350
                  TabIndex        =   46
                  Top             =   5700
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   25
                  Left            =   1350
                  TabIndex        =   45
                  Top             =   5925
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   26
                  Left            =   1350
                  TabIndex        =   44
                  Top             =   6150
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   27
                  Left            =   1350
                  TabIndex        =   43
                  Top             =   6375
                  Width           =   1290
               End
               Begin VB.Label Tu 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   28
                  Left            =   1350
                  TabIndex        =   42
                  Top             =   6600
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   0
                  Left            =   0
                  TabIndex        =   99
                  Top             =   300
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   1
                  Left            =   0
                  TabIndex        =   98
                  Top             =   525
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   2
                  Left            =   0
                  TabIndex        =   97
                  Top             =   750
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   3
                  Left            =   0
                  TabIndex        =   96
                  Top             =   975
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   4
                  Left            =   0
                  TabIndex        =   95
                  Top             =   1200
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   5
                  Left            =   0
                  TabIndex        =   94
                  Top             =   1425
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   6
                  Left            =   0
                  TabIndex        =   93
                  Top             =   1650
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   7
                  Left            =   0
                  TabIndex        =   92
                  Top             =   1875
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   8
                  Left            =   0
                  TabIndex        =   91
                  Top             =   2100
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   9
                  Left            =   0
                  TabIndex        =   90
                  Top             =   2325
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   10
                  Left            =   0
                  TabIndex        =   89
                  Top             =   2550
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   11
                  Left            =   0
                  TabIndex        =   88
                  Top             =   2775
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   12
                  Left            =   0
                  TabIndex        =   87
                  Top             =   3000
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   13
                  Left            =   0
                  TabIndex        =   86
                  Top             =   3225
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   14
                  Left            =   0
                  TabIndex        =   85
                  Top             =   3450
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   15
                  Left            =   0
                  TabIndex        =   84
                  Top             =   3675
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   16
                  Left            =   0
                  TabIndex        =   83
                  Top             =   3900
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   17
                  Left            =   0
                  TabIndex        =   82
                  Top             =   4125
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   18
                  Left            =   0
                  TabIndex        =   81
                  Top             =   4350
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   19
                  Left            =   0
                  TabIndex        =   80
                  Top             =   4575
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   20
                  Left            =   0
                  TabIndex        =   79
                  Top             =   4800
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   21
                  Left            =   0
                  TabIndex        =   78
                  Top             =   5025
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   22
                  Left            =   0
                  TabIndex        =   77
                  Top             =   5250
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   23
                  Left            =   0
                  TabIndex        =   76
                  Top             =   5475
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   24
                  Left            =   0
                  TabIndex        =   75
                  Top             =   5700
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   25
                  Left            =   0
                  TabIndex        =   74
                  Top             =   5925
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   26
                  Left            =   0
                  TabIndex        =   73
                  Top             =   6150
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   27
                  Left            =   0
                  TabIndex        =   72
                  Top             =   6375
                  Width           =   1290
               End
               Begin VB.Label MN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   28
                  Left            =   0
                  TabIndex        =   71
                  Top             =   6600
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   0
                  Left            =   2700
                  TabIndex        =   140
                  Top             =   300
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   1
                  Left            =   2700
                  TabIndex        =   141
                  Top             =   525
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   2
                  Left            =   2700
                  TabIndex        =   142
                  Top             =   750
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   3
                  Left            =   2700
                  TabIndex        =   143
                  Top             =   975
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   4
                  Left            =   2700
                  TabIndex        =   144
                  Top             =   1200
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   5
                  Left            =   2700
                  TabIndex        =   145
                  Top             =   1425
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   6
                  Left            =   2700
                  TabIndex        =   146
                  Top             =   1650
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   7
                  Left            =   2700
                  TabIndex        =   147
                  Top             =   1875
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   8
                  Left            =   2700
                  TabIndex        =   148
                  Top             =   2100
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   9
                  Left            =   2700
                  TabIndex        =   149
                  Top             =   2325
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   10
                  Left            =   2700
                  TabIndex        =   150
                  Top             =   2550
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   11
                  Left            =   2700
                  TabIndex        =   151
                  Top             =   2775
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   12
                  Left            =   2700
                  TabIndex        =   152
                  Top             =   3000
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   13
                  Left            =   2700
                  TabIndex        =   153
                  Top             =   3225
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   14
                  Left            =   2700
                  TabIndex        =   154
                  Top             =   3450
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   15
                  Left            =   2700
                  TabIndex        =   155
                  Top             =   3675
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   16
                  Left            =   2700
                  TabIndex        =   156
                  Top             =   3900
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   17
                  Left            =   2700
                  TabIndex        =   157
                  Top             =   4125
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   18
                  Left            =   2700
                  TabIndex        =   158
                  Top             =   4350
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   19
                  Left            =   2700
                  TabIndex        =   159
                  Top             =   4575
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   20
                  Left            =   2700
                  TabIndex        =   160
                  Top             =   4800
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   21
                  Left            =   2700
                  TabIndex        =   161
                  Top             =   5025
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   22
                  Left            =   2700
                  TabIndex        =   162
                  Top             =   5250
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   23
                  Left            =   2700
                  TabIndex        =   163
                  Top             =   5475
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   24
                  Left            =   2700
                  TabIndex        =   164
                  Top             =   5700
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   25
                  Left            =   2700
                  TabIndex        =   165
                  Top             =   5925
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   26
                  Left            =   2700
                  TabIndex        =   166
                  Top             =   6150
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   27
                  Left            =   2700
                  TabIndex        =   167
                  Top             =   6375
                  Width           =   1290
               End
               Begin VB.Label WD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   28
                  Left            =   2700
                  TabIndex        =   168
                  Top             =   6600
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   0
                  Left            =   5400
                  TabIndex        =   198
                  Top             =   300
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   1
                  Left            =   5400
                  TabIndex        =   199
                  Top             =   525
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   2
                  Left            =   5400
                  TabIndex        =   200
                  Top             =   750
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   3
                  Left            =   5400
                  TabIndex        =   201
                  Top             =   975
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   4
                  Left            =   5400
                  TabIndex        =   202
                  Top             =   1200
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   5
                  Left            =   5400
                  TabIndex        =   203
                  Top             =   1425
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   6
                  Left            =   5400
                  TabIndex        =   204
                  Top             =   1650
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   7
                  Left            =   5400
                  TabIndex        =   205
                  Top             =   1875
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   8
                  Left            =   5400
                  TabIndex        =   206
                  Top             =   2100
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   9
                  Left            =   5400
                  TabIndex        =   207
                  Top             =   2325
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   10
                  Left            =   5400
                  TabIndex        =   208
                  Top             =   2550
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   11
                  Left            =   5400
                  TabIndex        =   209
                  Top             =   2775
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   12
                  Left            =   5400
                  TabIndex        =   210
                  Top             =   3000
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   13
                  Left            =   5400
                  TabIndex        =   211
                  Top             =   3225
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   14
                  Left            =   5400
                  TabIndex        =   212
                  Top             =   3450
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   15
                  Left            =   5400
                  TabIndex        =   213
                  Top             =   3675
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   16
                  Left            =   5400
                  TabIndex        =   214
                  Top             =   3900
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   17
                  Left            =   5400
                  TabIndex        =   215
                  Top             =   4125
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   18
                  Left            =   5400
                  TabIndex        =   216
                  Top             =   4350
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   19
                  Left            =   5400
                  TabIndex        =   217
                  Top             =   4575
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   20
                  Left            =   5400
                  TabIndex        =   218
                  Top             =   4800
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   21
                  Left            =   5400
                  TabIndex        =   219
                  Top             =   5025
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   22
                  Left            =   5400
                  TabIndex        =   220
                  Top             =   5250
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   23
                  Left            =   5400
                  TabIndex        =   221
                  Top             =   5475
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   24
                  Left            =   5400
                  TabIndex        =   222
                  Top             =   5700
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   25
                  Left            =   5400
                  TabIndex        =   223
                  Top             =   5925
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   26
                  Left            =   5400
                  TabIndex        =   224
                  Top             =   6150
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   27
                  Left            =   5400
                  TabIndex        =   225
                  Top             =   6375
                  Width           =   1290
               End
               Begin VB.Label FR 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   28
                  Left            =   5400
                  TabIndex        =   226
                  Top             =   6600
                  Width           =   1290
               End
            End
         End
         Begin VB.ComboBox TIN 
            Enabled         =   0   'False
            Height          =   360
            ItemData        =   "MainForm.frx":1A13B
            Left            =   0
            List            =   "MainForm.frx":1A196
            Style           =   2  'Dropdown List
            TabIndex        =   257
            Top             =   0
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Select Record:"
            Height          =   240
            Left            =   2625
            TabIndex        =   139
            Top             =   300
            Width           =   1260
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "View Type:"
            Height          =   240
            Left            =   150
            TabIndex        =   138
            Top             =   300
            Width           =   975
         End
      End
   End
   Begin VB.Label LBFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Active Project"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   75
      TabIndex        =   12
      Top             =   0
      Width           =   7560
   End
   Begin VB.Menu MBFile 
      Caption         =   "&FILE"
      Begin VB.Menu FNew 
         Caption         =   "&New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu fOpen 
         Caption         =   "&Open Project"
      End
      Begin VB.Menu fClose 
         Caption         =   "&Close Project"
      End
      Begin VB.Menu FBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu fLock 
         Caption         =   "&Lock Application"
         Shortcut        =   ^L
      End
      Begin VB.Menu fBack 
         Caption         =   "&Back up Project"
      End
      Begin VB.Menu fBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu fExit 
         Caption         =   "&Quit System"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&HELP"
      Begin VB.Menu hHelp 
         Caption         =   "System Help File"
         Shortcut        =   {F1}
      End
      Begin VB.Menu hAbout 
         Caption         =   "&About us"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BdelSub_Click()
If LVSubjects.SelectedItem Is Nothing Then
    MsgBox "please select a record.", vbInformation, "Select"
Else
    'delete value
    If MsgBox("Delete Record?", vbQuestion + vbYesNo, "Delete?") = vbYes Then
        With LVSubjects
            SubjectQuery "Delete from Subjects where SC = " & Quoted(.SelectedItem.Text) & _
                " and Course = " & Quoted(.SelectedItem.SubItems(1)) & " and Yr = " & Quoted(.SelectedItem.SubItems(2))
            ConnectSubject
            SubjectsToLV MainForm.LVSubjects
        End With
    End If
End If
End Sub

Private Sub BeditSub_Click()
If LVSubjects.SelectedItem Is Nothing Then
    MsgBox "please select a record.", vbInformation, "Select"
Else
    AddnewNeed = False
    NeedActive = 1
    'Edit value
    FrmAddNeed.Show 1
    ConnectSubject
    SubjectsToLV LVSubjects
End If
End Sub

Private Sub BNewSub_Click()
    AddnewNeed = True
    NeedActive = 1 'set subject as active
    FrmAddNeed.Show 1
    ConnectSubject
    SubjectsToLV LVSubjects
End Sub

Private Sub BsearchSub_Click()
NeedActive = 1
FrmSearch.Show 1
SubjectsToLV LVSubjects
End Sub

Private Sub cbCreateReport_Click()
MsgBox "Disabled command. Use the binary file Repgen.exe at directory 'ProgramPath\RepGen\RepGen.exe'. Please Exit this application first before using the report generator to avoid inconvinience.", vbInformation, "Disabled patch"
Exit Sub
Screen.MousePointer = vbHourglass
'On Error GoTo xxx


Select Case CBView.ListIndex
    Case 0
    With Denver.rsSched_List_Grouping
    Dim mx As String
    NeedModule.SchedHandler CBSelect.Text
    If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        Dim mo As String
        ' SHAPE {Select Sc,Units,Descs From ReportTable}  AS Sched_List_Grouping APPEND ({Select * From REPORTTABLE}  AS Sched_List RELATE 'Sc' TO 'SC','Units' TO 'UNITS','Descs' TO 'DESCS') AS Sched_List
        mo = "  SHAPE {Select Sc,Units,Descs From ReportTable where sn = '" & CBSelect.Text & "'}  AS Sched_List_Grouping APPEND ({Select * From REPORTTABLE}  AS Sched_List RELATE 'Sc' TO 'SC','Units' TO 'UNITS','Descs' TO 'DESCS') AS Sched_List"
        Call .Open(mo, Denver.MyCON, adOpenDynamic, adLockOptimistic)
    End With
    With DRPSchedlist
        .Sections("pageheader").Controls("LSchool").Caption = UCase(InputBox("Enter School:", "SCHOOL"))
        .Sections("PageHeader").Controls("Lsection").Caption = CBSelect.Text
        .Sections("pageheader").Controls("LSem").Caption = UCase(InputBox("Enter Semester:", "SEMESTER"))
        .Sections("pageheader").Controls("LsY").Caption = TSy.Text
        .Sections("pageheader").Controls("Lunits").Caption = NeedModule.Count_Units(CBSelect.Text)
        Set .DataSource = Denver
        Screen.MousePointer = vbDefault
        .Show 1
    End With
    
    Case 1
        'Print teachers
        'SHAPE {Select * From SORTER order by Numbers}  AS HEADSt APPEND ({Select * From Schedules}  AS SubSt RELATE 'sorter' TO 'Day') AS SubSt
    With Denver.rsHEADSt
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        Dim mr As String
        mr = " SHAPE {Select * From SORTER order by Numbers}  AS HEADSt APPEND ({Select * From Schedules where teacher = '" & CBSelect.Text & "' order by tin}  AS SubSt RELATE 'sorter' TO 'Day') AS SubSt"
        Call .Open(mr, Denver.MyCON, adOpenDynamic, adLockOptimistic)
    End With
    With SchedTeachers
        .Sections("pageheader").Controls("LSchool").Caption = UCase(InputBox("Enter School:", "SCHOOL"))
        .Sections("PageHeader").Controls("LTeacher").Caption = CBSelect.Text
        .Sections("pageheader").Controls("LSem").Caption = UCase(InputBox("Enter Semester:", "SEMESTER"))
        .Sections("pageheader").Controls("LYear").Caption = TSy.Text
        Screen.MousePointer = vbDefault
        Set .DataSource = Denver
        .Show 1
    End With
    Case 2
        'Print Rooms
        'SHAPE {Select * From SORTER order by Numbers}  AS HEADSt APPEND ({Select * From Schedules}  AS SubSt RELATE 'sorter' TO 'Day') AS SubSt
    With Denver.rsHEAD_ROOMS
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        Dim myx1 As String
        myx = " SHAPE {Select * From SORTER order by Numbers}  AS HEAD_ROOMS APPEND ({Select * From Schedules where Room = '" & CBSelect.Text & "' order by tin}  AS ROOM_CONNECT RELATE 'sorter' TO 'Day') AS ROOM_CONNECT"
        Call .Open(myx, Denver.MyCON, adOpenDynamic, adLockOptimistic)
    End With
    With REPROOMS
        .Sections("pageheader").Controls("LSchool").Caption = UCase(InputBox("Enter School:", "SCHOOL"))
        .Sections("PageHeader").Controls("Lroom").Caption = CBSelect.Text
        .Sections("pageheader").Controls("LSem").Caption = UCase(InputBox("Enter Semester:", "SEMESTER"))
        .Sections("pageheader").Controls("LYear").Caption = TSy.Text
        Set .DataSource = Denver
        Screen.MousePointer = vbDefault
        .Show 1
    End With
End Select
Screen.MousePointer = vbDefault
Exit Sub
xxx:
If Err.Number = 8542 Then
    MsgBox "Please set your printer page layout to landscape to print this report.", vbInformation, "Landscape needed"
Else
MsgBox "Error: " & Err.Description, vbOKCancel + vbCritical, "Error: " & Err.Number
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub CBDeleteR_Click()
If LVRooms.SelectedItem Is Nothing Then
    MsgBox "please select a record.", vbInformation, "Select"
Else
    'delete value
    If MsgBox("Delete Record?", vbQuestion + vbYesNo, "Delete?") = vbYes Then
        With LVRooms
            RoomQuery "Delete from rooms where room = " & Quoted(.SelectedItem.Text)
            ConnectRooms
            RoomToLV LVRooms
        End With
    End If
End If
End Sub

Private Sub CBDeleteT_Click()
If LVTeachers.SelectedItem Is Nothing Then
    MsgBox "please select a record.", vbInformation, "Select"
Else
    'delete value
    If MsgBox("Delete Record?", vbQuestion + vbYesNo, "Delete?") = vbYes Then
        With LVTeachers
            TeachersQuery "Delete from teachers where teacher = " & Quoted(.SelectedItem.Text) & _
                " and subject = " & Quoted(.SelectedItem.SubItems(1))
            ConnectTeachers
            TeachersToLV LVTeachers
        End With
    End If
End If
End Sub

Private Sub CBDelSec_Click()
If LVSection.SelectedItem Is Nothing Then
    MsgBox "please select a record.", vbInformation, "Select"
Else
    'delete value
    If MsgBox("Delete Record?", vbQuestion + vbYesNo, "Delete?") = vbYes Then
        With LVSection
            SectionQuery "Delete from Sections where sectionname = " & Quoted(.SelectedItem.Text)
            ConnectSection
            SectionToLV LVSection
        End With
    End If
End If
End Sub

Private Sub CBEditR_Click()
If LVRooms.SelectedItem Is Nothing Then
    MsgBox "please select a record.", vbInformation, "Select"
Else
    AddnewNeed = False
    NeedActive = 4
    'Edit value
    FrmAddNeed.Show 1
    ConnectRooms
    RoomToLV LVRooms
End If
End Sub

Private Sub CBEditSec_Click()
If LVSection.SelectedItem Is Nothing Then
    MsgBox "please select a record.", vbInformation, "Select"
Else
    AddnewNeed = False
    NeedActive = 2
    'Edit value
    FrmAddNeed.Show 1
    ConnectSection
    SectionToLV LVSection
End If
End Sub

Private Sub CBEditT_Click()
If LVTeachers.SelectedItem Is Nothing Then
    MsgBox "please select a record.", vbInformation, "Select"
Else
    AddnewNeed = False
    NeedActive = 3
    'Edit value
    FrmAddNeed.Show 1
    ConnectTeachers
    TeachersToLV LVTeachers
End If
End Sub

Private Sub CBNewR_Click()
    AddnewNeed = True
    NeedActive = 4 'set subject as active
    FrmAddNeed.Show 1
    ConnectRooms
    RoomToLV LVRooms
End Sub

Private Sub CBNewSec_Click()
    AddnewNeed = True
    NeedActive = 2 'set subject as active
    FrmAddNeed.Show 1
    ConnectSection
    SectionToLV LVSection
End Sub

Private Sub CBNewT_Click()
    AddnewNeed = True
    NeedActive = 3 'set subject as active
    FrmAddNeed.Show 1
    ConnectTeachers
    TeachersToLV LVTeachers
End Sub

Private Sub CbPrintRoom_Click()
With Denver
    If .rsCmdRooms.State <> 0 Then .rsCmdRooms.Close
    
    .rsCmdRooms.Open "Select * From Rooms", .MyCON, adOpenDynamic, adLockOptimistic
End With
With DrepRooms
    .Sections("PageHeader").Controls("Lyear").Caption = MainForm.TSy.Text
    .Sections("PageHeader").Controls("Lschool").Caption = UCase(InputBox("Enter School:", "School"))
    .Sections("PageHeader").Controls("LSem").Caption = UCase(InputBox("Enter Semester:", "Semester"))
    .Refresh
    .Show 1
End With
End Sub

Private Sub CBPrintSection_Click()
With Denver
    If .rsSubSections.State <> 0 Then .rsSubSections.Close
    
    .rsSubSections.Open "Sections", .MyCON, adOpenDynamic, adLockOptimistic
End With
With DRepSections
    .Sections("PageHeader").Controls("Lyear").Caption = MainForm.TSy.Text
    .Sections("PageHeader").Controls("Lschool").Caption = UCase(InputBox("Enter School:", "School"))
    .Sections("PageHeader").Controls("LSem").Caption = UCase(InputBox("Enter Semester:", "Semester"))
    .Refresh
    .Show 1
End With
End Sub

Private Sub CBPrintSubs_Click()
RepActive = 1
FRMReps.Show 1
End Sub

Private Sub CBPrintTeacher_Click()
Dim mysql1 As String, mysql2 As String, myshape As String
With Denver
    If .rsTeachers.State <> 0 Then .rsTeachers.Close
            mysql1 = "SELECT teacher from teachers group by teacher"
            mysql2 = "Select * from teachers"
    myshape = " SHAPE {" & mysql1 & "} AS teachers APPEND ({" & mysql2 & "}  AS subTeachers RELATE 'teacher' TO 'Teacher') AS subTeachers"
    .rsTeachers.Open myshape, .MyCON, adOpenDynamic, adLockOptimistic
End With
With DrepTeachers
    .Sections("PageHeader").Controls("Lyear").Caption = MainForm.TSy.Text
    .Sections("PageHeader").Controls("Lschool").Caption = UCase(InputBox("Enter School:", "School"))
    .Sections("PageHeader").Controls("LSem").Caption = UCase(InputBox("Enter Semester:", "Semester"))
    .Refresh
    .Show 1
End With
End Sub

Private Sub CBprjUpdate_Click()
If Denver.MyCON.State = 0 Then Exit Sub
If isNothing(TPrjName.Text) = True Then
    MsgBox "Please enter a valid Project Name.", vbCritical, "Error"
    TPrjName.SetFocus
    Exit Sub
End If
If isNothing(TSy.Text) = True Then
    MsgBox "Please enter a valid School Year.", vbCritical, "Error"
    TSy.SetFocus
    Exit Sub
End If
If isNothing(TSem.Text) = True Then
    MsgBox "Please enter a valid Semester mode. 0 for Regular Semester and 1 for summer.", vbCritical, "Error"
    TSem.SetFocus
    Exit Sub
End If
On Error GoTo ErrorInfoUpdate
With RS_PrjInfo
    .Fields("ProjectName").Value = Trim(TPrjName.Text)
    .Fields("SY").Value = Trim(TSy.Text)
    .Fields("Mode").Value = Trim(TSem.Text)
    .Update
End With
Exit Sub
ErrorInfoUpdate:
    MainModule_Exception
    RS_PrjInfo.Cancel
    LoadInfo
    LoadInfoToForm
End Sub

Private Sub CBSearchR_Click()
NeedActive = 4
FrmSearch.Show 1
RoomToLV LVRooms
End Sub

Private Sub CBSearchSec_Click()
NeedActive = 2
FrmSearch.Show 1
SectionToLV LVSection
End Sub

Private Sub CBSearchT_Click()
NeedActive = 3
FrmSearch.Show 1
TeachersToLV LVTeachers
End Sub

Private Sub CBSelect_Click()
Select Case CBView.ListIndex
Case 0
    Loaders "M", CBSelect.Text, TIN, TOUT
    Loaders "T", CBSelect.Text, TIN, TOUT
    Loaders "W", CBSelect.Text, TIN, TOUT
    Loaders "TH", CBSelect.Text, TIN, TOUT
    Loaders "F", CBSelect.Text, TIN, TOUT
    Loaders "S", CBSelect.Text, TIN, TOUT
    'load the subject
    'get the yr and course
Dim course As String, yr As String, brk
brk = Split(CBSelect.Text, " ", , vbTextCompare)
course = brk(LBound(brk))
For i = LBound(brk) To UBound(brk)
    If IsNumeric(brk(i)) Then yr = brk(i): Exit For
Next
    'loadSubjectUnits LSubs, course, yr, TIN, TOUT, TSem.Text
Case 1
    OtherLoader "Select * From schedules where teacher = " & Quoted(CBSelect.Text) & " and day = 'M' order by tin", "M", TIN, TOUT
    OtherLoader "Select * From schedules where teacher = " & Quoted(CBSelect.Text) & " and day = 'T' order by tin", "T", TIN, TOUT
    OtherLoader "Select * From schedules where teacher = " & Quoted(CBSelect.Text) & " and day = 'W' order by tin", "W", TIN, TOUT
    OtherLoader "Select * From schedules where teacher = " & Quoted(CBSelect.Text) & " and day = 'TH' order by tin", "TH", TIN, TOUT
    OtherLoader "Select * From schedules where teacher = " & Quoted(CBSelect.Text) & " and day = 'F' order by tin", "F", TIN, TOUT
    OtherLoader "Select * From schedules where teacher = " & Quoted(CBSelect.Text) & " and day = 'S' order by tin", "S", TIN, TOUT
Case 2
    OtherLoader "Select * From schedules where room = " & Quoted(CBSelect.Text) & " and day = 'M' order by tin", "M", TIN, TOUT
    OtherLoader "Select * From schedules where room = " & Quoted(CBSelect.Text) & " and day = 'T' order by tin", "T", TIN, TOUT
    OtherLoader "Select * From schedules where room = " & Quoted(CBSelect.Text) & " and day = 'W' order by tin", "W", TIN, TOUT
    OtherLoader "Select * From schedules where room = " & Quoted(CBSelect.Text) & " and day = 'TH' order by tin", "TH", TIN, TOUT
    OtherLoader "Select * From schedules where room = " & Quoted(CBSelect.Text) & " and day = 'F' order by tin", "F", TIN, TOUT
    OtherLoader "Select * From schedules where room = " & Quoted(CBSelect.Text) & " and day = 'S' order by tin", "S", TIN, TOUT
End Select
End Sub

Private Sub CBTeacherLoads_Click()
FrmTeachersLoad.Show 1
End Sub

Private Sub CBView_Click()
Select Case CBView.ListIndex
Case 0
    whatLoad CBSelect, "Sections"
Case 1
    whatLoad CBSelect, "Teachers"
Case 2
    whatLoad CBSelect, "Rooms"
End Select
End Sub

Private Sub fBack_Click()
MsgBox "You can backup your Project by copying and pasting only.", vbInformation, "Backup"
End Sub

Private Sub fClose_Click()
    LSubs.Clear
    Close_Connection
End Sub

Private Sub fExit_Click()
Close_Connection
Unload Me
End
End Sub

Private Sub fLock_Click()
FrmLock.Show 1
End Sub

Private Sub FNew_Click()

Dim x As String, fso
CmdNew.ShowSave
    If CmdNew.FileName = "" Then Exit Sub
    If Right(App.Path, 1) = "\" Then
        x = App.Path & "default\default.mdb"
    Else
        x = App.Path & "\default\default.mdb"
    End If
    Set fso = CreateObject("Scripting.filesystemobject")
    If fso.fileexists(CmdNew.FileName) Then
        MsgBox "Can't Overwrite while creating New Project.", vbCritical, "ERROR"
        Exit Sub
    End If
    fso.copyfile x, CmdNew.FileName
    'Connect new Project
    If Connect(CmdNew.FileName) = 1 Then
    'LoadInfo
        LoadInfo
        LoadInfoToForm
        LoadAllNeeds
    End If
End Sub

Private Sub fOpen_Click()

    If Denver.MyCON.State <> 0 Then
        MsgBox "Please close your current connection first.", vbInformation, "Error"
        Exit Sub
    End If
    ClearInfo
    CMDOpen.ShowOpen
    If CMDOpen.FileName = "" Then Exit Sub
    If UCase(CMDOpen.FileName) = UCase(App.Path & "\default\default.mdb") Then
        MsgBox "Cannot use your default database.", vbCritical, "Error"
        Exit Sub
    End If
    If UCase(CMDOpen.FileName) = UCase(App.Path & "\Reports\default.mdb") Then
        MsgBox "Cannot use your default database.", vbCritical, "Error"
        Exit Sub
    End If
    If Connect(CMDOpen.FileName) = 1 Then
    'loadinfo
        LoadInfo
        LoadInfoToForm
        LoadAllNeeds
        MyActive_File = CMDOpen.FileName
    End If

End Sub

Private Sub Form_Resize()
'Check the Property window
'FORM minimum HGHT=8235 : WTDH=10395
'Property window minimum HGHT=7455 : LEFT=7680
On Error Resume Next
With Me
    If .Height < 8235 Then .Height = 8235
    If .Width < 10395 Then .Width = 10395
End With
With FrmProp
    .Height = Me.Height - 780
    .Left = Me.Width - 2715
End With
With LSubs
    .Height = FrmProp.Height - 1600
End With
With LBFile
    .Width = Me.Width - 2835
End With
With STAB
    .Width = Me.Width - 2880
    .Height = Me.Height - 1020
End With
With LVSubjects
    .Width = STAB.Width - 150
    .Height = STAB.Height - 1275
    FrmSub.Width = .Width
    FrmSub.Top = .Height + 365
    'sections
    LVSection.Width = .Width
    LVSection.Height = .Height
    Frmsec.Width = .Width
    Frmsec.Top = FrmSub.Top
    'teachers
    LVTeachers.Width = .Width
    LVTeachers.Height = .Height
    FrmTeach.Width = .Width
    FrmTeach.Top = FrmSub.Top
    'rooms
    LVRooms.Width = .Width
    LVRooms.Height = .Height
    FrmR.Width = .Width
    FrmR.Top = FrmSub.Top
End With
With Pback
    .Left = STAB.Width / 2 - .Width / 2
    .Top = STAB.Height / 2 - .Height / 2 + 100
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close_Connection
End Sub

Private Sub FR_Click(Index As Integer)
If CBView.ListIndex = 0 And CBSelect.ListIndex <> -1 Then
    Dim HSG As Integer, Sql As String, lxt As String
TempInd = Index
'Get the height
    If FR(Index).BackColor = vbWhite Then addSched = True Else addSched = False
    HSG = FR(Index).Height / 240
    Load FrmSchedadd
    Dim SendWhat, i As Integer, x As String, Y As String, z As String
    If FR(Index).Caption = "." Then
        x = "": Y = "": z = ""
    Else
        SendWhat = Split(FR(Index).Caption, "|", , vbTextCompare)
        x = SendWhat(0): Y = SendWhat(1): z = SendWhat(2)
    End If
    If Index >= 28 Then
        MsgBox "You cannot Create Schedule here.", vbInformation, "ERROR"
        Exit Sub
    End If
    'Set RS_Scheds = New Recordset
    With RS_Scheds
        
        TIN.ListIndex = Index
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .ActiveConnection = Denver.MyCON
        Sql = "Select * From Schedules where day = 'F' and sn like '%" & CBSelect.Text & "%' " & _
            " and tin =" & TimeorDate(TIN.Text)
        Call .Open(Sql, , adOpenDynamic, adLockOptimistic)
        
        If .RecordCount = 0 Then
            lxt = CBSelect.Text
        Else
            lxt = .Fields!sn
        End If

    End With
    
    Get_Class lxt, Index, HSG + Index, "F", x, z, Y
    FrmSchedadd.Caption = CBSelect.Text
    FrmSchedadd.Show 1
    Dim yr As String, course As String
    yr = Right(CBSelect.Text, 1)
    course = Left(CBSelect.Text, Len(CBSelect.Text) - 1)
    CBSelect_Click
End If
End Sub

Private Sub hAbout_Click()
FrmAbout.Show 1
End Sub

Private Sub hHelp_Click()
Shell App.Path & "\PRJ_HELP\prjhelp.exe", vbNormalFocus
End Sub

Private Sub Hscrlclass_Change()
'750 - 1200
With PClass
    .Left = 750 - (Hscrlclass.Value * 40)
End With
End Sub

Private Sub LSubs_Click()
LSubs.ToolTipText = LSubs.Text
End Sub

Private Sub LVRooms_DblClick()
    If LVRooms.SelectedItem Is Nothing Then Exit Sub
    CBEditR_Click
End Sub

Private Sub LVSection_DblClick()
    If LVSection.SelectedItem Is Nothing Then Exit Sub
    CBEditSec_Click
End Sub

Private Sub LVSubjects_DblClick()
    If LVSubjects.SelectedItem Is Nothing Then Exit Sub
    BeditSub_Click
End Sub

Private Sub LVTeachers_DblClick()
    If LVTeachers.SelectedItem Is Nothing Then Exit Sub
    CBEditT_Click
End Sub

Private Sub mbc_Click()
If LSubs.Left <= -1500 Then
    LSubs.Left = -1500
Else
    LSubs.Left = LSubs.Left - 100
End If
End Sub

Private Sub MN_Click(Index As Integer)
If CBView.ListIndex = 0 And CBSelect.ListIndex <> -1 Then
    Dim HSG As Integer, Sql As String, lxt As String
    TempInd = Index
    
'Get the height
    If MN(Index).BackColor = vbWhite Then addSched = True Else addSched = False
    HSG = MN(Index).Height / 240
    Load FrmSchedadd
    Dim SendWhat, i As Integer, x As String, Y As String, z As String
    If MN(Index).Caption = "." Then
        x = "": Y = "": z = ""
    Else
        SendWhat = Split(MN(Index).Caption, "|", , vbTextCompare)
        x = SendWhat(0): Y = SendWhat(1): z = SendWhat(2)
    End If
    If Index >= 28 Then
        MsgBox "You cannot Create Schedule here.", vbInformation, "ERROR"
        Exit Sub
    End If
    'Set RS_Scheds = New Recordset
    With RS_Scheds
        
        TIN.ListIndex = Index
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .ActiveConnection = Denver.MyCON
        Sql = "Select * From Schedules where day = 'M' and sn like '%" & CBSelect.Text & "%' " & _
            " and tin =" & TimeorDate(TIN.Text)
        Call .Open(Sql, , adOpenDynamic, adLockOptimistic)
        
        If .RecordCount = 0 Then
            lxt = CBSelect.Text
        Else
            lxt = .Fields!sn
            FrmSchedadd.LUnits.Caption = .Fields!Units
            If IsNull(.Fields!Attribs) Then FrmSchedadd.TAttrib.Text = "" Else FrmSchedadd.TAttrib.Text = .Fields!Attribs
            If IsNull(.Fields!Descs) Then FrmSchedadd.LDT.Caption = "" Else FrmSchedadd.LDT.Caption = .Fields!Descs
        End If

    End With
    
    Get_Class lxt, Index, HSG + Index, "M", x, z, Y
    FrmSchedadd.Caption = CBSelect.Text
    FrmSchedadd.Show 1
    Dim course As String, yr As String, brk
brk = Split(CBSelect.Text, " ", , vbTextCompare)
course = brk(LBound(brk))
For i = LBound(brk) To UBound(brk)
    If IsNumeric(brk(i)) Then yr = brk(i): Exit For
Next
    loadSubjectUnits LSubs, course, yr, TIN, TOUT, TSem.Text
    CBSelect_Click
End If
End Sub

Private Sub mpr_Click()
If LSubs.Left >= 0 Then
    LSubs.Left = 0
Else
    LSubs.Left = LSubs.Left + 100
End If
End Sub

Private Sub SAT_Click(Index As Integer)
If CBView.ListIndex = 0 And CBSelect.ListIndex <> -1 Then
    Dim HSG As Integer, Sql As String, lxt As String
TempInd = Index
'Get the height
    If SAT(Index).BackColor = vbWhite Then addSched = True Else addSched = False
    HSG = SAT(Index).Height / 240
    Load FrmSchedadd
    Dim SendWhat, i As Integer, x As String, Y As String, z As String
    If SAT(Index).Caption = "." Then
        x = "": Y = "": z = ""
    Else
        SendWhat = Split(SAT(Index).Caption, "|", , vbTextCompare)
        x = SendWhat(0): Y = SendWhat(1): z = SendWhat(2)
    End If
    If Index >= 28 Then
        MsgBox "You cannot Create Schedule here.", vbInformation, "ERROR"
        Exit Sub
    End If
    'Set RS_Scheds = New Recordset
    With RS_Scheds
        
        TIN.ListIndex = Index
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .ActiveConnection = Denver.MyCON
        Sql = "Select * From Schedules where day = 'S' and sn like '%" & CBSelect.Text & "%' " & _
            " and tin = " & TimeorDate(TIN.Text)
        Call .Open(Sql, , adOpenDynamic, adLockOptimistic)
        
        If .RecordCount = 0 Then
            lxt = CBSelect.Text
        Else
            lxt = .Fields!sn
        End If

    End With
    
    Get_Class lxt, Index, HSG + Index, "S", x, z, Y
    FrmSchedadd.Caption = CBSelect.Text
    FrmSchedadd.Show 1
    Dim yr As String, course As String
    yr = Right(CBSelect.Text, 1)
    course = Left(CBSelect.Text, Len(CBSelect.Text) - 1)
    CBSelect_Click
End If
End Sub

Private Sub STAB_Click(PreviousTab As Integer)
If STAB.Tab = 5 And CBSelect.ListIndex = 0 Then
Else
    'clear the subject area
    LSubs.Clear
End If
End Sub

Private Sub TH_Click(Index As Integer)
If CBView.ListIndex = 0 And CBSelect.ListIndex <> -1 Then
    Dim HSG As Integer, Sql As String, lxt As String
TempInd = Index
'Get the height
    If TH(Index).BackColor = vbWhite Then addSched = True Else addSched = False
    HSG = TH(Index).Height / 240
    Load FrmSchedadd
    Dim SendWhat, i As Integer, x As String, Y As String, z As String
    If TH(Index).Caption = "." Then
        x = "": Y = "": z = ""
    Else
        SendWhat = Split(TH(Index).Caption, "|", , vbTextCompare)
        x = SendWhat(0): Y = SendWhat(1): z = SendWhat(2)
    End If
    If Index >= 28 Then
        MsgBox "You cannot Create Schedule here.", vbInformation, "ERROR"
        Exit Sub
    End If
    'Set RS_Scheds = New Recordset
    With RS_Scheds
        
        TIN.ListIndex = Index
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .ActiveConnection = Denver.MyCON
        Sql = "Select * From Schedules where day = 'TH' and sn like '%" & CBSelect.Text & "%' " & _
            " and tin =" & TimeorDate(TIN.Text)
        Call .Open(Sql, , adOpenDynamic, adLockOptimistic)
        
        If .RecordCount = 0 Then
            lxt = CBSelect.Text
        Else
            lxt = .Fields!sn
        End If

    End With
    
    Get_Class lxt, Index, HSG + Index, "TH", x, z, Y
    FrmSchedadd.Caption = CBSelect.Text
    FrmSchedadd.Show 1
    Dim yr As String, course As String
    yr = Right(CBSelect.Text, 1)
    course = Left(CBSelect.Text, Len(CBSelect.Text) - 1)
    CBSelect_Click
End If
End Sub

Private Sub TSem_Change()
    Select Case TSem.Text
    Case "3"
        LblCaps.Caption = "Summer"
    Case "2"
        LblCaps.Caption = "2nd Sem"
    Case "1"
        LblCaps.Caption = "First Sem"
    Case ""
        
    Case Else
        MsgBox "Enter 1,2,3 only.", vbCritical, "Error"
        TSem.SetFocus
    End Select
    
End Sub

Private Sub Tu_Click(Index As Integer)
If CBView.ListIndex = 0 And CBSelect.ListIndex <> -1 Then
    Dim HSG As Integer, Sql As String, lxt As String
TempInd = Index
'Get the height
    If Tu(Index).BackColor = vbWhite Then addSched = True Else addSched = False
    HSG = Tu(Index).Height / 240
    Load FrmSchedadd
    Dim SendWhat, i As Integer, x As String, Y As String, z As String
    If Tu(Index).Caption = "." Then
        x = "": Y = "": z = ""
    Else
        SendWhat = Split(Tu(Index).Caption, "|", , vbTextCompare)
        x = SendWhat(0): Y = SendWhat(1): z = SendWhat(2)
    End If
    If Index >= 28 Then
        MsgBox "You cannot Create Schedule here.", vbInformation, "ERROR"
        Exit Sub
    End If
    'Set RS_Scheds = New Recordset
    With RS_Scheds
        
        TIN.ListIndex = Index
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .ActiveConnection = Denver.MyCON
        Sql = "Select * From Schedules where day = 'T' and sn like '%" & CBSelect.Text & "%' " & _
            " and tin =" & TimeorDate(TIN.Text)
        Call .Open(Sql, , adOpenDynamic, adLockOptimistic)
        
        If .RecordCount = 0 Then
            lxt = CBSelect.Text
        Else
            lxt = .Fields!sn
        End If

    End With
    
    Get_Class lxt, Index, HSG + Index, "T", x, z, Y
    FrmSchedadd.Caption = CBSelect.Text
    FrmSchedadd.Show 1
Dim course As String, yr As String, brk
    brk = Split(CBSelect.Text, " ", , vbTextCompare)
    course = brk(LBound(brk))
    For i = LBound(brk) To UBound(brk)
        If IsNumeric(brk(i)) Then yr = brk(i): Exit For
    Next

    loadSubjectUnits LSubs, course, yr, TIN, TOUT, TSem.Text
    CBSelect_Click
End If
End Sub

Private Sub VScrclass_Change()
'-1275
With PClass
    .Top = 0 - (VScrclass.Value * 25)
    pxx.Top = .Top
End With
End Sub

Private Sub WD_Click(Index As Integer)
If CBView.ListIndex = 0 And CBSelect.ListIndex <> -1 Then
    Dim HSG As Integer, Sql As String, lxt As String
TempInd = Index
'Get the height
    If WD(Index).BackColor = vbWhite Then addSched = True Else addSched = False
    HSG = WD(Index).Height / 240
    Load FrmSchedadd
    Dim SendWhat, i As Integer, x As String, Y As String, z As String
    If WD(Index).Caption = "." Then
        x = "": Y = "": z = ""
    Else
        SendWhat = Split(WD(Index).Caption, "|", , vbTextCompare)
        x = SendWhat(0): Y = SendWhat(1): z = SendWhat(2)
    End If
    If Index >= 28 Then
        MsgBox "You cannot Create Schedule here.", vbInformation, "ERROR"
        Exit Sub
    End If
    'Set RS_Scheds = New Recordset
    With RS_Scheds
        
        TIN.ListIndex = Index
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .ActiveConnection = Denver.MyCON
        Sql = "Select * From Schedules where day = 'W' and sn like '%" & CBSelect.Text & "%' " & _
            " and tin = " & TimeorDate(TIN.Text)
        .Open Sql, , adOpenDynamic, adLockOptimistic
        
        If .RecordCount = 0 Then
            lxt = CBSelect.Text
        Else
            lxt = .Fields!sn
        End If

    End With
    
    Get_Class lxt, Index, HSG + Index, "W", x, z, Y
    FrmSchedadd.Caption = CBSelect.Text
    FrmSchedadd.Show 1
    Dim yr As String, course As String
    yr = Right(CBSelect.Text, 1)
    course = Left(CBSelect.Text, Len(CBSelect.Text) - 1)
    CBSelect_Click
End If
End Sub
