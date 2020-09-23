Attribute VB_Name = "MyMod"
 Public Mycon As New ADODB.Connection

Sub main()
    ChDir App.Path
    With DE
        If .CON.State <> 0 Then .CON.Close
        .CON.CursorLocation = adUseClient
        .CON.Open "Provider=MSDataShape;data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Reports\Default.mdb;Persist Security Info=False;Jet OLEDB:Database Password=vip"
    End With
    'FrmViewer.Show
    FrmMain.Show
End Sub

Private Function CreatedRemarks(ByVal section As String, ByVal Selected As String) As String

Dim i As Integer, Fract As String, mysplit, j As Integer

If section = Selected Then GoTo FractOut
mysplit = Split(section, ",", , vbTextCompare)
For j = LBound(mysplit) To UBound(mysplit)
    If mysplit(j) = Selected Then
        Fract = Fract
    Else
        Fract = Fract & mysplit(j) & ","
    End If
Next
Fract = Left(Fract, Len(Fract) - 1)
'If Len("," & Selected) = Len(Section) - Len(Fract) Then
'Fract = Left(Fract, Len(Fract))
'Else
'Fract = Fract & Right(Section, Len(Section) - (Len(Fract) + Len(Selected) - 1))
'End If

FractOut:
CreatedRemarks = Fract

End Function

Public Function TeacherScheds(Teacher As String)
On Error GoTo erbx
Dim FLetter As String, PREV_I As String, Prev_O As String, Sections As String, Subject As String
Dim ADDX As String, Desc As String, Comp As String, COL_SEC As String
Dim GoBack As String, i As Double, NewLetter As String
Dim rs1 As New Recordset, RS2 As New Recordset

With rs1
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    Call .Open("Select * From Schedules where Teacher = '" & _
    Teacher & "' order by Sc,tin asc", Mycon, adOpenDynamic, adLockOptimistic)
End With
With RS2
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    Call .Open("delete From ReportTable", DE.CON, adOpenDynamic, adLockOptimistic)
End With

'Get first record
If rs1.RecordCount = 0 Then Exit Function

    FLetter = Left(rs1.Fields("DAY").Value, 1)
        If FLetter = "T" Then
'Check next
        If Left(rs1.Fields("DAY").Value, 2) = "TH" Then
            FLetter = "TH"
        End If
        
        End If
    
'RS2.Recordset.AddNew
With rs1
    Do Until .EOF
NewLetter = Left(rs1.Fields("DAY").Value, 1)
        If NewLetter = "T" Then
'Check next
        If Left(rs1.Fields("DAY").Value, 2) = "TH" Then
            NewLetter = "TH"
        End If
        End If

    PREV_I = .Fields("TIn").Value
    Prev_O = .Fields("TOut").Value
    Subject = .Fields("Sc").Value
    Sections = .Fields("sn").Value
    
    If .Fields("Attribs").Value <> vbNullString Then
        Desc = .Fields("Descs").Value & "(" & .Fields("attribs").Value & ")"
    Else
        Desc = .Fields("Descs").Value
    End If
    
    COL_SEC = Teacher '.Fields("teacher").Value
    
        GoBack = "Select * from ReportTable where teacher = '" & COL_SEC & _
            "' and sc = '" & Subject & _
            "' and Schedule = '" & PREV_I & " - " & Prev_O & "' and SN = '" & _
            Sections & "'"
            
        If RS2.State <> 0 Then RS2.Close
        RS2.CursorLocation = adUseClient
        Call RS2.Open(GoBack, DE.CON, adOpenDynamic, adLockOptimistic)
        
        If RS2.RecordCount = 0 Then
        
            RS2.AddNew
            FLetter = NewLetter
            
        Else
        
            If Left(.Fields("Day").Value, 1) = "T" Then
                
                If Left(.Fields("DAY").Value, 2) = "TH" Then
                   ADDX = Left(.Fields("Day").Value, 2)
                Else
                    ADDX = Left(.Fields("Day").Value, 1)
                 End If
            Else
                   ADDX = Left(.Fields("DAY").Value, 1)
            End If
             FLetter = FLetter & ADDX
             
        End If
        
        RS2.Fields("teacher").Value = Teacher
        RS2.Fields("Sc").Value = Subject
        RS2.Fields("Descs").Value = Desc
        RS2.Fields("day").Value = FLetter
        RS2.Fields("Schedule").Value = PREV_I & " - " & Prev_O
        RS2.Fields("sn").Value = Sections
        RS2.Fields("Remarks").Value = .Fields("Room").Value
        RS2.Fields("Units").Value = .Fields("Units").Value
        RS2.Update
        
NextRecord:
        .MoveNext
Loop
End With

Set rs1 = Nothing
Set RS2 = Nothing
Exit Function
erbx:
    MsgBox Err.Description, vbCritical, "Error"
    If Mycon.State <> 0 Then Mycon.Close
    Set rs1 = Nothing
    Set RS2 = Nothing
End Function
Public Function SchedHandler(section As String)
On Error GoTo erbx
Dim FLetter As String, PREV_I As String, Prev_O As String, Teacher As String, Subject As String
Dim ADDX As String, Desc As String, Comp As String, COL_SEC As String
Dim GoBack As String, i As Double, NewLetter As String
Dim rs1 As New Recordset, RS2 As New Recordset
With rs1
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    Call .Open("Select * From Schedules where sn like '%" & _
    section & "%' order by Sc asc", Mycon, adOpenDynamic, adLockOptimistic)
End With
With RS2
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    Call .Open("delete From ReportTable", DE.CON, adOpenDynamic, adLockOptimistic)
End With

'Get first record
If rs1.RecordCount = 0 Then Exit Function

    FLetter = Left(rs1.Fields("DAY").Value, 1)
        If FLetter = "T" Then
'Check next
        If Left(rs1.Fields("DAY").Value, 2) = "TH" Then
            FLetter = "TH"
        End If
        
        End If
    
'RS2.Recordset.AddNew
With rs1
    Do Until .EOF
NewLetter = Left(rs1.Fields("DAY").Value, 1)
        If NewLetter = "T" Then
'Check next
        If Left(rs1.Fields("DAY").Value, 2) = "TH" Then
            NewLetter = "TH"
        End If
        End If

    PREV_I = .Fields("TIn").Value
    Prev_O = .Fields("TOut").Value
    Subject = .Fields("Sc").Value
    Teacher = .Fields("Teacher").Value
    
    If .Fields("Attribs").Value <> vbNullString Then
        Desc = .Fields("Descs").Value & "(" & .Fields("attribs").Value & ")"
    Else
        Desc = .Fields("Descs").Value
    End If
    
    COL_SEC = section '.Fields("SECTIONKO").Value
    
        GoBack = "Select * from ReportTable where Sn = '" & COL_SEC & _
            "' and sc = '" & Subject & _
            "' and Schedule = '" & PREV_I & " - " & Prev_O & "' and Teacher = '" & _
            Teacher & "'"
            
        If RS2.State <> 0 Then RS2.Close
        RS2.CursorLocation = adUseClient
        Call RS2.Open(GoBack, DE.CON, adOpenDynamic, adLockOptimistic)
        
        If RS2.RecordCount = 0 Then
        
            RS2.AddNew
            FLetter = NewLetter
            
        Else
        
            If Left(.Fields("Day").Value, 1) = "T" Then
                
                If Left(.Fields("DAY").Value, 2) = "TH" Then
                   ADDX = Left(.Fields("Day").Value, 2)
                Else
                    ADDX = Left(.Fields("Day").Value, 1)
                 End If
            Else
                   ADDX = Left(.Fields("DAY").Value, 1)
            End If
             FLetter = FLetter & ADDX
             
        End If
        
        RS2.Fields("Sn").Value = section
        RS2.Fields("Sc").Value = Subject
        RS2.Fields("Descs").Value = Desc
        RS2.Fields("day").Value = FLetter
        RS2.Fields("Schedule").Value = PREV_I & " - " & Prev_O
        RS2.Fields("TEACHER").Value = Teacher
        RS2.Fields("Remarks").Value = CreatedRemarks(.Fields("sn").Value, section)
        RS2.Fields("Units").Value = .Fields("Units").Value
        RS2.Update
        
NextRecord:
        .MoveNext
Loop
End With
Set rs1 = Nothing
Set RS2 = Nothing
Exit Function
erbx:
    MsgBox Err.Description, vbCritical, "Error"
    If Mycon.State <> 0 Then Mycon.Close
    Set rs1 = Nothing
    Set RS2 = Nothing
End Function

Public Function roomsched(room As String)
On Error GoTo erbx
Dim FLetter As String, PREV_I As String, Prev_O As String, Teacher As String, Subject As String
Dim ADDX As String, Desc As String, Comp As String, COL_SEC As String, section As String
Dim GoBack As String, i As Double, NewLetter As String
Dim rs1 As New Recordset, RS2 As New Recordset
With rs1
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    Call .Open("Select * From Schedules where room = '" & _
    room & "' order by Sc,tin asc", Mycon, adOpenDynamic, adLockOptimistic)
End With
With RS2
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    Call .Open("delete From ReportTable", DE.CON, adOpenDynamic, adLockOptimistic)
End With

'Get first record
If rs1.RecordCount = 0 Then Exit Function

    FLetter = Left(rs1.Fields("DAY").Value, 1)
        If FLetter = "T" Then
'Check next
        If Left(rs1.Fields("DAY").Value, 2) = "TH" Then
            FLetter = "TH"
        End If
        
        End If
    
'RS2.Recordset.AddNew
With rs1
    Do Until .EOF
NewLetter = Left(rs1.Fields("DAY").Value, 1)
        If NewLetter = "T" Then
'Check next
        If Left(rs1.Fields("DAY").Value, 2) = "TH" Then
            NewLetter = "TH"
        End If
        End If

    PREV_I = .Fields("TIn").Value
    Prev_O = .Fields("TOut").Value
    Subject = .Fields("Sc").Value
    Teacher = .Fields("Teacher").Value
    section = .Fields("Sn").Value
    
    If .Fields("Attribs").Value <> vbNullString Then
        Desc = .Fields("Descs").Value & "(" & .Fields("attribs").Value & ")"
    Else
        Desc = .Fields("Descs").Value
    End If
    
    COL_SEC = room '.Fields("SECTIONKO").Value
    
        GoBack = "Select * from ReportTable where remarks = '" & COL_SEC & _
            "' and sc = '" & Subject & _
            "' and Schedule = '" & PREV_I & " - " & Prev_O & "' and Teacher = '" & _
            Teacher & "'"
            
        If RS2.State <> 0 Then RS2.Close
        RS2.CursorLocation = adUseClient
        Call RS2.Open(GoBack, DE.CON, adOpenDynamic, adLockOptimistic)
        
        If RS2.RecordCount = 0 Then
        
            RS2.AddNew
            FLetter = NewLetter
            
        Else
        
            If Left(.Fields("Day").Value, 1) = "T" Then
                
                If Left(.Fields("DAY").Value, 2) = "TH" Then
                   ADDX = Left(.Fields("Day").Value, 2)
                Else
                    ADDX = Left(.Fields("Day").Value, 1)
                 End If
            Else
                   ADDX = Left(.Fields("DAY").Value, 1)
            End If
             FLetter = FLetter & ADDX
             
        End If
        
        RS2.Fields("Sn").Value = section
        RS2.Fields("Sc").Value = Subject
        RS2.Fields("Descs").Value = Desc
        RS2.Fields("day").Value = FLetter
        RS2.Fields("Schedule").Value = PREV_I & " - " & Prev_O
        RS2.Fields("TEACHER").Value = Teacher
        RS2.Fields("Remarks").Value = .Fields("Room").Value
        RS2.Fields("Units").Value = .Fields("Units").Value
        RS2.Update
        
NextRecord:
        .MoveNext
Loop
End With
Set rs1 = Nothing
Set RS2 = Nothing
Exit Function
erbx:
    MsgBox Err.Description, vbCritical, "Error"
    If Mycon.State <> 0 Then Mycon.Close
    Set rs1 = Nothing
    Set RS2 = Nothing
End Function


Function classLoad(LV As ListView)
LV.ListItems.Clear
Dim rs As New Recordset
With rs
    .CursorLocation = adUseClient
    .Open "Select * From ReportTable", DE.CON, adOpenDynamic, adLockOptimistic
    If .RecordCount <> 0 Then
        Do Until .EOF
            LV.ListItems.Add .AbsolutePosition, , .Fields("Day").Value
            LV.ListItems(.AbsolutePosition).SubItems(1) = .Fields("SC").Value
            LV.ListItems(.AbsolutePosition).SubItems(2) = .Fields("Descs").Value
            LV.ListItems(.AbsolutePosition).SubItems(3) = .Fields("Units").Value
            LV.ListItems(.AbsolutePosition).SubItems(4) = .Fields("Schedule").Value
            LV.ListItems(.AbsolutePosition).SubItems(5) = .Fields("Teacher").Value
            LV.ListItems(.AbsolutePosition).SubItems(6) = .Fields("Remarks").Value
            .MoveNext
        Loop
    End If
    .Close
End With
Set rs = Nothing
End Function

Function TeacherLoad(LV As ListView)
LV.ListItems.Clear
Dim rs As New Recordset
With rs
    .CursorLocation = adUseClient
    .Open "Select * From ReportTable", DE.CON, adOpenDynamic, adLockOptimistic
    If .RecordCount <> 0 Then
        Do Until .EOF
            LV.ListItems.Add .AbsolutePosition, , .Fields("Day").Value
            LV.ListItems(.AbsolutePosition).SubItems(1) = .Fields("SC").Value
            LV.ListItems(.AbsolutePosition).SubItems(2) = .Fields("Descs").Value
            LV.ListItems(.AbsolutePosition).SubItems(3) = .Fields("Units").Value
            LV.ListItems(.AbsolutePosition).SubItems(4) = .Fields("Schedule").Value
            LV.ListItems(.AbsolutePosition).SubItems(5) = .Fields("sn").Value
            LV.ListItems(.AbsolutePosition).SubItems(6) = .Fields("Remarks").Value
            .MoveNext
        Loop
    End If
    .Close
End With
Set rs = Nothing
End Function

Function RoomLoad(LV As ListView)
LV.ListItems.Clear
Dim rs As New Recordset
With rs
    .CursorLocation = adUseClient
    .Open "Select * From ReportTable", DE.CON, adOpenDynamic, adLockOptimistic
    If .RecordCount <> 0 Then
        Do Until .EOF
            LV.ListItems.Add .AbsolutePosition, , .Fields("Day").Value
            LV.ListItems(.AbsolutePosition).SubItems(1) = .Fields("SC").Value
            LV.ListItems(.AbsolutePosition).SubItems(2) = .Fields("Descs").Value
            LV.ListItems(.AbsolutePosition).SubItems(3) = .Fields("Units").Value
            LV.ListItems(.AbsolutePosition).SubItems(4) = .Fields("Schedule").Value
            LV.ListItems(.AbsolutePosition).SubItems(5) = .Fields("Teacher").Value
            LV.ListItems(.AbsolutePosition).SubItems(6) = .Fields("sn").Value
            .MoveNext
        Loop
    End If
    .Close
End With
Set rs = Nothing
End Function

