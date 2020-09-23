VERSION 5.00
Begin VB.Form FrmTeachersLoad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Teachers Load"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   3585
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LTlist 
      Height          =   2940
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   5565
   End
   Begin VB.CommandButton CBClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   465
      Left            =   4425
      TabIndex        =   0
      Top             =   3075
      Width           =   1215
   End
End
Attribute VB_Name = "FrmTeachersLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private myRset As New Recordset, Rsch As New Recordset
Private MyString As String
Private xp As Double

Private Sub CBClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
connectrs
End Sub

Function connectrs()
Dim myx As Double
With Rsch
If .State <> 0 Then .Close
    .ActiveConnection = Denver.MyCON
    .CursorLocation = adUseClient
    Call .Open("Select Teacher from Teachers group by teacher", , adOpenDynamic, adLockOptimistic)
    Do Until .EOF
        Call ConnectSet("Select * from Schedules where teacher = " & Quoted(.Fields(0).Value) & " order by SN,SC")
        myx = Try_Strings
        LTlist.AddItem .Fields(0).Value & " " & xp
        .MoveNext
    Loop
End With
End Function
Private Function ConnectSet(ByVal mx As String)
With myRset
    If .State <> 0 Then .Close
    .ActiveConnection = Denver.MyCON
    .CursorLocation = adUseClient
    Call .Open(mx, , adOpenDynamic, adLockOptimistic)
End With
End Function

Private Function Try_Strings() As Double

Dim SubAct As String
Dim Tunits As Double
With myRset
MyString = ""
Do Until .EOF
    If SubAct <> .Fields!sc Then
        Tunits = Tunits + .Fields!Units
        SubAct = .Fields!sc
         MyString = MyString & "|" & .Fields!sc & "," & .Fields!sn
    Else
        'store to mystr
        If finders(.Fields!sc & "," & .Fields!sn) = True Then
            Tunits = Tunits
        Else
            Tunits = Tunits + .Fields!Units
            MyString = MyString & "|" & .Fields!sc & "," & .Fields!sn
        End If
    End If
    .MoveNext
Loop
End With
try_String = Tunits
xp = Tunits
End Function

Function finders(What As String) As Boolean
Dim xp, finding, xfind As Boolean
xfind = True
Dim i As Integer, j As Integer
xp = Split(MyString, "|", , vbTextCompare)
finding = Split(What, ",", , vbTextCompare)
For i = LBound(xp) To UBound(xp)
    For j = LBound(finding) To UBound(finding)
        If xfind = False Then xfind = True: Exit For
        If InStr(1, xp(i), finding(j), vbTextCompare) Then
            xfind = True
            If j = UBound(finding) Then GoTo MySets
        Else
            xfind = False
        End If
    Next
Next
MySets:
finders = xfind
End Function
