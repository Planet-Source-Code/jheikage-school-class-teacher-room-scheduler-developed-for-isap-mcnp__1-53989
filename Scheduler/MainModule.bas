Attribute VB_Name = "MainModule"
'Public ADOCON As New ADODB.Connection
Public RS_PrjInfo As New ADODB.Recordset
Public RS_Subjects As New ADODB.Recordset
Public RS_Sections As New ADODB.Recordset
Public RS_Teachers As New ADODB.Recordset
Public RS_GroupSBS As New ADODB.Recordset
Public RS_Room As New ADODB.Recordset
Public RS_Scheds As New ADODB.Recordset
'Security
Public MyS As New SparkClass
Public Namx As String, PASX As String
'TRICK
Public Function ColorFullme(pb As PictureBox)
Dim i As Integer, r As Integer, g As Integer, b As Integer, x As Double
With pb
    .DrawWidth = 5
    .ScaleMode = vbPixels
    .ScaleWidth = 255
    'Drawlines
    .Cls
    For i = 0 To 255
        If r >= 79 Then
            r = 79
        Else
            r = r + 1
        End If
        g = r: b = r
        pb.Line (x, 0)-(x, pb.Height), RGB(r, g, b), BF
        x = x + 1.5
    Next
    .ScaleMode = vbTwips
End With
End Function
'END TRICK
Sub Main()
    MyS.ShowForm
    If MyS.LoginEnable(True) = True Then
    Namx = MyS.XNAM
    PASX = MyS.XPAS
    MaxOp
    ChDir App.Path
    MainForm.Show
    End If
End Sub

Public Function MainModule_Exception()
MsgBox Err.Description, vbCritical, "Error"
End Function

Public Function Connect(DataSource As String) As Integer
On Error GoTo ErrorConnect
    Dim ConStr As String

    If Denver.MyCON.State <> 0 Then MyCON.Close
    'Provider=MSDataShape.1;Extended Properties="Jet OLEDB:Database Password=vip";Persist Security Info=False;Mode=Read;Data Source=C:\Test.mdb;Data Provider=MICROSOFT.JET.OLEDB.4.0
    ConStr = "Provider=msDatashape.1;Persist Security Info=False;mode=share deny write|share deny read;Data Source=" & DataSource & ";Data Provider=Microsoft.jet.oledb.4.0;JET OLEDB:Database password=vip"
    With Denver.MyCON
        .CursorLocation = adUseClient
        .ConnectionString = ConStr
        Call .Open
    End With
    ' if successful return 1 else return 0
    Connect = 1
    MainForm.LBFile.Caption = "File Opened:" & Denver.MyCON.Properties("Data Source").Value
    MainForm.STAB.Visible = True
Exit Function
ErrorConnect:
    MainModule_Exception
    Close_Connection
    Connect = 0
End Function

Public Function Close_Connection()  'Close your current connection
With Denver.MyCON
    If .State <> 0 Then .Close
    MainForm.LBFile.Caption = "No Active Project"
    MainForm.STAB.Visible = False
    ClearInfo
End With
End Function

Public Function Quoted(Str As String) As String
Quoted = "'" & Str & "'"
End Function

Public Function TimeorDate(Str As String) As String
TimeorDate = "#" & Str & "#"
End Function

Public Function isNothing(Str As String) As Boolean
If Len(Trim(Str)) = 0 Then isNothing = True Else isNothing = False
End Function

Public Function LoadAllNeeds()
On Error GoTo ErrorLoadAll
    'load subjects
    ConnectSubject
    SubjectsToLV MainForm.LVSubjects
    ConnectSection
    SectionToLV MainForm.LVSection
    ConnectTeachers
    TeachersToLV MainForm.LVTeachers
    ConnectRooms
    RoomToLV MainForm.LVRooms
Exit Function
ErrorLoadAll:
    MainModule_Exception
    Close_Connection
End Function
'-------------------------------------------------------------------------------------
'Functions for the property window connection[this will check if it is a valid sproject

Public Function LoadInfo()
On Error GoTo ErrorLoadInfo
With RS_PrjInfo
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    Call .Open("PrjInfo", , adOpenDynamic, adLockOptimistic)
End With
Exit Function
ErrorLoadInfo:
    MainModule_Exception
    Close_Connection
End Function

Public Function LoadInfoToForm()
On Error GoTo ErrorLItoForm
With MainForm
    .TPrjName.Text = RS_PrjInfo("ProjectName").Value
    .TSy.Text = RS_PrjInfo("SY").Value
    .TSem.Text = RS_PrjInfo("Mode").Value
End With
Exit Function
ErrorLItoForm:
    MainModule_Exception
End Function

Public Function ClearInfo()
With MainForm
    .TPrjName.Text = ""
    .TSy.Text = ""
    .TSem.Text = ""
End With
End Function
