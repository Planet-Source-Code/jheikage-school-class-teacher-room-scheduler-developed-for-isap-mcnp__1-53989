VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FRMHELP 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Help File"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMHELP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMHELP.frx":628A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CBCLOSE 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton CBSHOW 
      Caption         =   "Show Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   4440
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RBOX1 
      Height          =   4215
      Left            =   3000
      TabIndex        =   1
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7435
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"FRMHELP.frx":BA7C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView TVW1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   7646
      _Version        =   393217
      HideSelection   =   0   'False
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label LBLTOP 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Topic..."
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
      Height          =   735
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "FRMHELP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ADOCON As ADODB.Connection
Private RsCategories As ADODB.Recordset
Private RsSubCat As ADODB.Recordset
Private RsOpen As ADODB.Recordset
'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\ISAPMCNP\Patch1\Help\PRJ_HELP\HELP.mdb;Persist Security Info=False;Jet OLEDB:Database Password=vip
Private Sub CBCLOSE_Click()
End
End Sub

Private Sub CBSHOW_Click()
If TVW1.SelectedItem Is Nothing Then Exit Sub
If TVW1.SelectedItem.Tag = Null Then
    RBOX1.Text = "<<NOTHING>>"
Else
    RBOX1.Text = TVW1.SelectedItem.Tag
End If
Exit Sub
Set RsOpen = Nothing
Set RsOpen = New ADODB.Recordset
With RsOpen
    .ActiveConnection = ADOCON
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "SELECT CATEGORY FROM TAB_HELP GROUP BY CATEGORY ORDER BY CATEGORY"
End With

End Sub

Private Sub Form_Load()
ShowCursor True
Connect
        ScaleMode = vbPixels
        ScaleWidth = 256
        DrawWidth = 4
        For i = 0 To 255
        Line (x, 0)-(x, Height), RGB(0, 0, i), BF
        x = x + 1
        Next

End Sub

Sub Connect()
Dim files As String
Set ADOCON = Nothing
Set ADOCON = New ADODB.Connection
With ADOCON
    .CursorLocation = adUseClient
    files = App.Path & "\HELP.mdb"
    .Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;JET OLEDB:Database Password=vip;data source=" & files
End With
Load_Categories
End Sub

Sub Load_Categories()
Set RsCategories = Nothing
Set RsCategories = New ADODB.Recordset
With RsCategories
    .ActiveConnection = ADOCON
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "SELECT CATEGORY FROM TAB_HELP GROUP BY CATEGORY ORDER BY CATEGORY"
    Do Until .EOF
        TVW1.Nodes.Add , , , .Fields(0).Value, 1, 1
        TVW1.Nodes(TVW1.Nodes.Count).Tag = .Fields(0).Value
        Sub_Cat .Fields(0).Value, TVW1.Nodes.Count
        .MoveNext
    Loop
End With
Set RsCategories = Nothing
End Sub

Function Sub_Cat(Categories As String, RET As Integer)
Set RsSubCat = Nothing
Set RsSubCat = New ADODB.Recordset
With RsSubCat
    .ActiveConnection = ADOCON
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "SELECT * From TAB_HELP where Category='" & Categories & "' order by Nam"
    Do Until .EOF
 
        TVW1.Nodes.Add RET, tvwChild, , .Fields(1).Value, 1, 1
        If IsNull(.Fields(2).Value) Then
        TVW1.Nodes(TVW1.Nodes.Count).Tag = "There are no items to display here..."
        Else
        TVW1.Nodes(TVW1.Nodes.Count).Tag = .Fields(2).Value
        End If
        .MoveNext
    Loop
    .Close
End With
Set RsSubCat = Nothing
End Function

Private Sub Form_Resize()
        ScaleMode = vbPixels
        ScaleWidth = 256
        DrawWidth = 4
        For i = 0 To 255
        Line (x, 0)-(x, Height), RGB(0, 0, i), BF
        x = x + 1
        Next

End Sub

Private Sub TVW1_DblClick()
CBSHOW_Click
End Sub
