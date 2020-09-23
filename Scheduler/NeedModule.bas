Attribute VB_Name = "NeedModule"
Public AddnewNeed As Boolean
Public NeedActive As Integer
Public RepActive As Integer
'NeedActive/RepActive
'1 = subjects
'2 = course
'3 = sections
'4 = teachers
'5 = rooms
'End of NeedActive

'Codes for Subjects-------
Public Function ConnectSubject()
With RS_Subjects
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    Call .Open("Select * From subjects order by Course and Yr", , adOpenDynamic, adLockOptimistic)
End With
End Function

Public Function SubjectQuery(Sql As String)
With RS_Subjects
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    Call .Open(Sql, , adOpenDynamic, adLockOptimistic)
End With
End Function

Public Function SubjectsToLV(LV As ListView)
With RS_Subjects
    LV.ListItems.Clear
    If .RecordCount = 0 Then Exit Function
    Dim Itm As ListItem
    Do
        Set Itm = LV.ListItems.Add(, , .Fields!sc, 1, 1)
        Itm.SubItems(1) = .Fields!course
        Itm.SubItems(2) = .Fields!yr
        Itm.SubItems(3) = .Fields!Desc
        Itm.SubItems(4) = .Fields!Units
        .MoveNext
    Loop Until .EOF
End With
End Function

'end of codes for subjects

'Codes for Sections-------
Public Function ConnectSection()
With RS_Sections
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    Call .Open("Select * From Sections", , adOpenDynamic, adLockOptimistic)
End With
End Function

Public Function SectionQuery(Sql As String)
With RS_Sections
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    Call .Open(Sql, , adOpenDynamic, adLockOptimistic)
End With
End Function

Public Function SectionToLV(LV As ListView)
With RS_Sections
    LV.ListItems.Clear
    If .RecordCount = 0 Then Exit Function
    Dim Itm As ListItem
    Do
        Set Itm = LV.ListItems.Add(, , .Fields!Sectionname, 2, 2)
        .MoveNext
    Loop Until .EOF
End With
End Function

Public Function LoadSectionstoCombo(CB As ComboBox)
On Error GoTo ErrorLoadSTC
Dim Sql As String
With RS_GroupSBS
    CB.Clear
    If .State <> 0 Then .Close
    Sql = "Select Course from Subjects group by Course"
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    Call .Open(Sql, , adOpenDynamic, adLockOptimistic)
    If .RecordCount <> 0 Then
        Do
            CB.AddItem .Fields!course
            .MoveNext
        Loop Until .EOF
    End If
End With
Exit Function
ErrorLoadSTC:
    MainModule_Exception
End Function
'end of codes for sections

'codes for teachers-------
Public Function ConnectTeachers()
With RS_Teachers
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    Call .Open("Select * From Teachers order by teacher", , adOpenDynamic, adLockOptimistic)
End With
End Function

Public Function TeachersQuery(Sql As String)
With RS_Teachers
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    Call .Open(Sql, , adOpenDynamic, adLockOptimistic)
End With
End Function

Public Function TeachersToLV(LV As ListView)
With RS_Teachers
    LV.ListItems.Clear
    If .RecordCount = 0 Then Exit Function
    Dim Itm As ListItem
    Do
        Set Itm = LV.ListItems.Add(, , .Fields!Teacher, 3, 3)
        Itm.SubItems(1) = .Fields!Subject
        .MoveNext
    Loop Until .EOF
End With
End Function

'end of codes for teachers
'Room Codes-------
Public Function ConnectRooms()
With RS_Room
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    Call .Open("Select * From rooms", , adOpenDynamic, adLockOptimistic)
End With
End Function

Public Function RoomQuery(Sql As String)
With RS_Room
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    Call .Open(Sql, , adOpenDynamic, adLockOptimistic)
End With
End Function

Public Function RoomToLV(LV As ListView)
With RS_Room
    LV.ListItems.Clear
    If .RecordCount = 0 Then Exit Function
    Dim Itm As ListItem
    Do
        Set Itm = LV.ListItems.Add(, , .Fields!room, 4, 4)
        .MoveNext
    Loop Until .EOF
End With
End Function

'end of room codes

'Scheduling List Codes

Public Function SchedHandler(Section As String)
'On error resume next
Dim FLetter As String, PREV_I As String, Prev_O As String, Teacher As String, Subject As String
Dim ADDX As String, Desc As String, Comp As String, COL_SEC As String
Dim GoBack As String, i As Double, NewLetter As String
Dim rs1 As New Recordset, RS2 As New Recordset
With rs1
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    Call .Open("Select * From Schedules where sn like '%" & _
    Section & "%' order by Sc asc", Denver.MyCON, adOpenDynamic, adLockOptimistic)
End With
With RS2
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    Call .Open("delete From ReportTable", Denver.MyCON, adOpenDynamic, adLockOptimistic)
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
    
    COL_SEC = Section '.Fields("SECTIONKO").Value
    
        GoBack = "Select * from ReportTable where Sn = '" & COL_SEC & _
            "' and sc = '" & Subject & _
            "' and Schedule = '" & PREV_I & " - " & Prev_O & "' and Teacher = '" & _
            Teacher & "'"
            
        If RS2.State <> 0 Then RS2.Close
        RS2.CursorLocation = adUseClient
        Call RS2.Open(GoBack, Denver.MyCON, adOpenDynamic, adLockOptimistic)
        
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
        
        RS2.Fields("Sn").Value = Section
        RS2.Fields("Sc").Value = Subject
        RS2.Fields("Descs").Value = Desc
        RS2.Fields("day").Value = FLetter
        RS2.Fields("Schedule").Value = PREV_I & " - " & Prev_O
        RS2.Fields("TEACHER").Value = Teacher
        RS2.Fields("Remarks").Value = CreatedRemarks(.Fields("sn").Value, Section)
        RS2.Fields("Units").Value = .Fields("Units").Value
        RS2.Update
        
NextRecord:
        .MoveNext
Loop
End With
Set rs1 = Nothing
Set RS2 = Nothing
End Function

Public Function Count_Units(ByVal Section As String) As Double
    Dim Subject As String, tot_unit As Double
    Dim rs1 As New Recordset
    With rs1
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    Call .Open("Select * From ReportTable where sn = '" & Section & "'order by Sn", Denver.MyCON, adOpenDynamic, adLockOptimistic)
    End With
    tot_unit = 0
    Do Until rs1.EOF
        If rs1.Fields("Sc").Value <> Subject Then
            tot_unit = tot_unit + rs1.Fields("Units").Value
            Subject = rs1.Fields("Sc").Value
        End If
        rs1.MoveNext
    Loop
    Count_Units = tot_unit
Set rs1 = Nothing
End Function

Private Function CreatedRemarks(ByVal Section As String, ByVal Selected As String) As String

Dim i As Integer, Fract As String, mysplit, j As Integer

If Section = Selected Then GoTo FractOut
mysplit = Split(Section, ",", , vbTextCompare)
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
'end SLC


Sub MaxOp()
    Dim File As String, gig As String, enc As String, sync As String, maxs As Integer
    If Right(App.Path, 1) = "\" Then
    File = App.Path & "NIGOLFILE.FJV"
    Else
    File = App.Path & "\" & "NIGOLFILE.FJV"
    End If
    gig = "[USER:" & Namx & "]" & "[DATE:" & Date & "] [TIME:" & Time & "]"
    Dim i As Integer
    maxs = Len(gig)
    For i = 1 To maxs
        'Encript File
        sync = Left(gig, 1)
        gig = Right(gig, maxs - i)
        
        'MsgBox gig
        enc = enc & Chr(234) & sync
    Next
    frmuser.GetFileName File
    frmuser.Show 1
    
    Open File For Output As #1
        
        Write #1, enc
        
    Close #1
    
End Sub
