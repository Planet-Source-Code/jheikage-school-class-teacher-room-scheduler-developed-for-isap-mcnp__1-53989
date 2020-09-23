Attribute VB_Name = "SchedModule"
Public Conflict As Boolean
Public TempInd As Integer
Public addSched As Boolean
Public CurrentSection As String
Public colt As String
Public cmd_Usched As New Command

'Temporary files
Public TEMP_CLASS As String, TEMP_IN As String, TEMP_OUT As String, TEMP_ROOM As _
    String, TEMP_DAY As String, TEMP_TEACH As String, TEMP_YR As String, TEMP_SUB As _
    String, TEMP_UNIT As String, TEMP_SECTION As String
'Return Message
Public Return_Message As String

'Record sets
Public RS_CheckT As ADODB.Recordset
Public RS_CheckR As ADODB.Recordset

Public Function Create_BoxHeight(ByVal BoxNumber As Integer, ByVal EndTime As Integer, day As String, _
    Subject As String, TimeIn As String, TimeOut As String, room As String, Teacher As String)
    Dim Times As String, i As Integer, NewHeight As Single, RR As String
    Conflict = False
With MainForm
    NewHeight = (EndTime - BoxNumber) * 230
    Select Case day
    Case "M"
    
        If BoxNumber = EndTime Then
        .MN(BoxNumber).BackColor = vbWhite
        .MN(BoxNumber).Height = 495
        .MN(BoxNumber).ToolTipText = ""
        .MN(BoxNumber).Caption = ""
        .MN(BoxNumber).Height = 480
        .MN(BoxNumber).FontBold = False
        .MN(BoxNumber).Visible = False
        Else
        'Check for some more other active box
        
        For i = BoxNumber To EndTime
            If i <> BoxNumber Then
                
                If .MN(i).BackColor = vbGreen Then
                    If .MN(i).ToolTipText <> "" Then
                        'MsgBox MS(i).ToolTipText & vbNewLine & TimeOut
                        'If TimeValue(.MN(i).ToolTipText) < TimeValue(TimeOut) Then
                        MsgBox "Can't Create Schedule. Class Conflict Detected.", vbExclamation, "SCHEDULING"
                        Conflict = True
                        Exit Function
                        'End If
                    End If
                End If
            End If
        Next
        
        .MN(BoxNumber).Height = NewHeight
        .MN(BoxNumber).BackColor = vbGreen
        .MN(BoxNumber).Visible = True
        .MN(BoxNumber).Height = NewHeight
        .MN(BoxNumber).Caption = Subject & "|" & room & "|" & Teacher
        .MN(BoxNumber).ToolTipText = .MN(BoxNumber).Caption
        .MN(BoxNumber).FontBold = True
        End If
        
    Case "T"
        
        If BoxNumber = EndTime Then
        .Tu(BoxNumber).BackColor = vbWhite
        .Tu(BoxNumber).Height = 495
        .Tu(BoxNumber).ToolTipText = ""
        .Tu(BoxNumber).Caption = ""
        .Tu(BoxNumber).Height = 480
        .Tu(BoxNumber).FontBold = False
        .Tu(BoxNumber).Visible = False
        Else
        'Check for some more other active box
        For i = BoxNumber To EndTime
            If i <> BoxNumber Then
                
                If .Tu(i).BackColor = vbGreen Then
                    If .Tu(i).ToolTipText <> "" Then
                        'If TimeValue(.Tu(i).ToolTipText) < TimeValue(TimeOut) Then
                        MsgBox "Can't Create Schedule. Class Conflict Detected.", vbExclamation, "SCHEDULING"
                        Conflict = True
                        Exit Function
                        'End If
                    End If
                End If
            End If
        Next
        .Tu(BoxNumber).Height = NewHeight
        .Tu(BoxNumber).BackColor = vbGreen
        .Tu(BoxNumber).Visible = True
        .Tu(BoxNumber).Height = NewHeight
        .Tu(BoxNumber).Caption = Subject & "|" & room & "|" & Teacher
        .Tu(BoxNumber).ToolTipText = .Tu(BoxNumber).Caption
        .Tu(BoxNumber).FontBold = True
        End If
        
    Case "W"
    
    If BoxNumber = EndTime Then
        .WD(BoxNumber).BackColor = vbWhite
        .WD(BoxNumber).Height = 495
        .WD(BoxNumber).ToolTipText = ""
        .WD(BoxNumber).Caption = ""
        .WD(BoxNumber).Height = 480
        .WD(BoxNumber).FontBold = False
        .WD(BoxNumber).Visible = False
        Else
        'Check for some more other active box
        For i = BoxNumber To EndTime
            If i <> BoxNumber Then
                
                If .WD(i).BackColor = vbGreen Then
                    If .WD(i).ToolTipText <> "" Then
                        'MsgBox MS(i).ToolTipText & vbNewLine & TimeOut
                        'If TimeValue(.WD(i).ToolTipText) < TimeValue(TimeOut) Then
                        MsgBox "Can't Create Schedule. Class Conflict Detected.", vbExclamation, "SCHEDULING"
                        Conflict = True
                        Exit Function
                        'End If
                    End If
                End If
            End If
        Next
        .WD(BoxNumber).Height = NewHeight
        .WD(BoxNumber).BackColor = vbGreen
        .WD(BoxNumber).Visible = True
        .WD(BoxNumber).Height = NewHeight
        .WD(BoxNumber).Caption = Subject & "|" & room & "|" & Teacher
        .WD(BoxNumber).ToolTipText = .WD(BoxNumber).Caption
        .WD(BoxNumber).FontBold = True
        End If
    
    Case "TH"
    
    If BoxNumber = EndTime Then
        .TH(BoxNumber).BackColor = vbWhite
        .TH(BoxNumber).Height = 495
        .TH(BoxNumber).ToolTipText = ""
        .TH(BoxNumber).Caption = ""
        .TH(BoxNumber).Height = 480
        .TH(BoxNumber).FontBold = False
        .TH(BoxNumber).Visible = False
        Else
        'Check for some more other active box
        For i = BoxNumber To EndTime
            If i <> BoxNumber Then
                
                If .TH(i).BackColor = vbGreen Then
                    If .TH(i).ToolTipText <> "" Then
                        'MsgBox MS(i).ToolTipText & vbNewLine & TimeOut
                        'If TimeValue(.TH(i).ToolTipText) < TimeValue(TimeOut) Then
                        MsgBox "Can't Create Schedule. Class Conflict Detected.", vbExclamation, "SCHEDULING"
                        Conflict = True
                        Exit Function
                        'End If
                    End If
                End If
            End If
        Next
        .TH(BoxNumber).Height = NewHeight
        .TH(BoxNumber).BackColor = vbGreen
        .TH(BoxNumber).Visible = True
        .TH(BoxNumber).Height = NewHeight
        .TH(BoxNumber).Caption = Subject & "|" & room & "|" & Teacher
       .TH(BoxNumber).ToolTipText = .TH(BoxNumber).Caption
        .TH(BoxNumber).FontBold = True
        End If
        
    Case "F"
    
    If BoxNumber = EndTime Then
        .FR(BoxNumber).BackColor = vbWhite
        .FR(BoxNumber).Height = 495
        .FR(BoxNumber).ToolTipText = ""
        .FR(BoxNumber).Caption = ""
        .FR(BoxNumber).Height = 480
        .FR(BoxNumber).FontBold = False
        .FR(BoxNumber).Visible = False
        Else
        'Check for some more other active box
        For i = BoxNumber To EndTime
            If i <> BoxNumber Then
                
                If .FR(i).BackColor = vbGreen Then
                    If .FR(i).ToolTipText <> "" Then
                        'MsgBox MS(i).ToolTipText & vbNewLine & TimeOut
                        'If TimeValue(.FR(i).ToolTipText) < TimeValue(TimeOut) Then
                        MsgBox "Can't Create Schedule. Class Conflict Detected.", vbExclamation, "SCHEDULING"
                        Conflict = True
                        Exit Function
                        'End If
                    End If
                End If
            End If
        Next
        .FR(BoxNumber).Height = NewHeight
        .FR(BoxNumber).BackColor = vbGreen
        .FR(BoxNumber).Visible = True
        .FR(BoxNumber).Height = NewHeight
        .FR(BoxNumber).Caption = Subject & "|" & room & "|" & Teacher
        .FR(BoxNumber).ToolTipText = .FR(BoxNumber).Caption
        .FR(BoxNumber).FontBold = True
        End If
    
    Case "S"
    
    If BoxNumber = EndTime Then
        .SAT(BoxNumber).BackColor = vbWhite
        .SAT(BoxNumber).Height = 495
        .SAT(BoxNumber).ToolTipText = ""
        .SAT(BoxNumber).Caption = ""
        .SAT(BoxNumber).Height = 480
        .SAT(BoxNumber).FontBold = False
        .SAT(BoxNumber).Visible = False
        Else
        'Check for some more other active box
        For i = BoxNumber To EndTime
            If i <> BoxNumber Then
                
                If .SAT(i).BackColor = vbGreen Then
                    If .SAT(i).ToolTipText <> "" Then
                        'MsgBox MS(i).ToolTipText & vbNewLine & TimeOut
                        'If TimeValue(.SAT(i).ToolTipText) < TimeValue(TimeOut) Then
                        MsgBox "Can't Create Schedule. Class Conflict Detected.", vbExclamation, "SCHEDULING"
                        Conflict = True
                        Exit Function
                        'End If
                    End If
                End If
            End If
        Next
        .SAT(BoxNumber).Height = NewHeight
        .SAT(BoxNumber).BackColor = vbGreen
        .SAT(BoxNumber).Visible = True
        .SAT(BoxNumber).Height = NewHeight
        .SAT(BoxNumber).Caption = Subject & "|" & room & "|" & Teacher
        .SAT(BoxNumber).ToolTipText = .SAT(BoxNumber).Caption
        .SAT(BoxNumber).FontBold = True
        End If
    End Select
End With
End Function

'create the first impresion
Public Function Get_Class(ByVal Class As String, ByVal TimeIndex As Integer, LastIndex As Integer, day As String, Subject As String, Teacher As String, room As String)
With FrmSchedadd
    .LSection.Caption = Class
    .TIN.ListIndex = TimeIndex
    .TOUT.ListIndex = LastIndex
    .LDay.Caption = day
    .CBSubs.Text = Subject
    .CBTeacher.Text = Teacher
    .cbroom.Text = room
    TempInd = LastIndex
    GetTemporaryFiles
End With
End Function

Public Sub Load_TEMPS()
With FrmSchedadd
'call this if canceled
    .LSection.Caption = TEMP_CLASS
    .TIN.Text = TEMP_OUT
    .TOUT.Text = TEMP_YR
    .CBTeacher.Text = TEMP_TEACH
    .CBSubs.Text = TEMP_SUB
    .LUnits.Caption = TEMP_UNIT
    .cbroom.Text = TEMP_ROOM
End With
End Sub

Public Sub GetTemporaryFiles()
With FrmSchedadd
TEMP_IN = .TIN.Text
TEMP_ROOM = .cbroom.Text
TEMP_DAY = .LDay.Caption
TEMP_TEACH = .CBTeacher.Text
TEMP_SUB = .CBSubs.Text
TEMP_OUT = .TOUT.Text
TEMP_UNIT = .LUnits.Caption
TEMP_CLASS = .LSection.Caption
End With
End Sub

'Conflict Checking
Public Function CheckSection(rs As Recordset, CSect As String, Teach As String, Sbj As String, TIN As String, TOUT As String, day As String, room As String) As String
On Error GoTo exp
Dim Sql As String
'If addSched = False Then
    'update only
    Sql = "select * from Schedules where Day = " & Quoted(day) & " and teacher = " & Quoted(Teach) & " and Room = " & Quoted(room) & " and tin = " & _
        TimeorDate(TIN) & " and tout = " & TimeorDate(TOUT) & " and sc = " & Quoted(Sbj)
    With rs
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .ActiveConnection = Denver.MyCON
        Call .Open(Sql, , adOpenDynamic, adLockOptimistic)
        If .RecordCount = 0 Then
            CheckSection = CSect
            colt = CSect
        Else
            colt = .Fields!sn
        If InStr(1, .Fields!sn, CSect, vbTextCompare) = 0 Then
            If MsgBox("Do you wish to combine " & CSect & " with " & .Fields!sn & "?", vbYesNo + vbQuestion, "Combine") = vbYes Then
                CheckSection = .Fields!sn & "," & CSect
            Else
                CheckSection = ""
            End If
        Else
            CheckSection = .Fields!sn
        End If
        End If
    End With
'Else
'    CheckSection = CSect
'End If
Exit Function
exp:
MainModule_Exception
CheckSection = CSect
colt = CSect
End Function

Public Function CheckConflict(ByVal MyField As String, rs As Recordset, SCT As String, _
    Sbj As String, TIN As String, TOUT As String, day As String, Check As String) As Boolean
    'Check is name of room/teacher to check
On Error GoTo CheckConflictErr
Dim SQL1 As String, SQL2 As String, Msg As String
With rs
    'Check phase 1
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    SQL1 = "select * From Schedules where day = " & Quoted(day) & _
        " and " & MyField & " = " & Quoted(Check) & " and Tin < " & _
        TimeorDate(TIN) & " and tout > " & TimeorDate(TIN)
    Call .Open(SQL1, , adOpenDynamic, adLockOptimistic)
    If .RecordCount <> 0 Then
    If comparetemp(rs) = False Then
        Return_Message = "Phase1: Conflict in Time in with " & _
            .Fields!sn & vbNewLine _
            & .Fields!TIN & " to " & .Fields!TOUT
        CheckConflict = False
        Exit Function
    End If
    End If
    'Check phase 2
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    SQL1 = "select * From Schedules where day = " & Quoted(day) & _
        " and " & MyField & " = " & Quoted(Check) & " and Tin = " & _
        TimeorDate(TIN) & " and tout <> " & TimeorDate(TOUT)
    Call .Open(SQL1, , adOpenDynamic, adLockOptimistic)
    If .RecordCount <> 0 Then
    If comparetemp(rs) = False Then
        Return_Message = "Phase 2: Conflict detected with " & _
            .Fields!sn & vbNewLine & .Fields!TIN & " to " & .Fields!TOUT
        CheckConflict = False
        Exit Function
    End If
    End If
    'Check phase 3
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    SQL1 = "select * From Schedules where day = " & Quoted(day) & _
        " and " & MyField & " = " & Quoted(Check) & " and Tin > " & _
        TimeorDate(TIN) & " and tout <= " & TimeorDate(TOUT)
    Call .Open(SQL1, , adOpenDynamic, adLockOptimistic)
    If .RecordCount <> 0 Then
    If comparetemp(rs) = False Then
        Return_Message = "Phase 3: Conflict in Time in with " & _
            .Fields!sn & vbNewLine _
            & .Fields!TIN & " to " & .Fields!TOUT
        CheckConflict = False
        Exit Function
    End If
    End If
    'Check phase 4
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    SQL1 = "select * From Schedules where day = " & Quoted(day) & _
        " and " & MyField & " = " & Quoted(Check) & " and Tin > " & _
         TimeorDate(TOUT) & " and tout < " & TimeorDate(TOUT)
    Call .Open(SQL1, , adOpenDynamic, adLockOptimistic)
    If .RecordCount <> 0 Then
    If comparetemp(rs) = False Then
        Return_Message = "Phase 4: Conflict detected with " & _
            .Fields!sn & vbNewLine _
            & .Fields!TIN & " to " & .Fields!TOUT
        CheckConflict = False
        Exit Function
    End If
    End If
    'check phase 5
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    SQL1 = "select * From Schedules where day = " & Quoted(day) & _
        " and " & MyField & " = " & Quoted(Check) & " and Tout > " & _
        TimeorDate(TOUT) & " and tin < " & TimeorDate(TIN)
    Call .Open(SQL1, , adOpenDynamic, adLockOptimistic)
    If .RecordCount <> 0 Then
    If comparetemp(rs) = False Then
        Return_Message = "Phase 5: Conflict detected with " & _
            .Fields!sn & vbNewLine _
            & .Fields!TIN & " to " & .Fields!TOUT
        CheckConflict = False
        Exit Function
    End If
    End If
'phase 6
If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    SQL1 = "select * From Schedules where day = " & Quoted(day) & _
        " and " & MyField & " = " & Quoted(Check) & " and Tin > " & _
        TimeorDate(TIN) & " and tout = " & TimeorDate(TOUT)
    Call .Open(SQL1, , adOpenDynamic, adLockOptimistic)
    If .RecordCount <> 0 Then
    If comparetemp(rs) = False Then
        Return_Message = "Phase 6: Conflict in Time with " & _
            .Fields!sn & vbNewLine _
            & .Fields!TIN & " to " & .Fields!TOUT
        CheckConflict = False
        Exit Function
    End If
    End If
'phase 7
If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .ActiveConnection = Denver.MyCON
    SQL1 = "Select * From Schedules where day = " & Quoted(day) & _
        " and " & MyField & " = " & Quoted(Check) & " and Tin = " & _
        TimeorDate(TIN) & " and tout = " & TimeorDate(TOUT)
    .Open SQL1, , adOpenDynamic, adLockOptimistic
    If .RecordCount <> 0 Then
    'If comparetemp(rs) = False Then
    '    Return_Message = "Conflict in Time with " & _
    '        .Fields!sn & vbNewLine _
    '        & .Fields!TIN & " to " & .Fields!TOUT
    '    CheckConflict = False
    '    Exit Function
    'End If
    Select Case UCase(MyField)
    Case "ROOM"
        If Trim(.Fields!Teacher) <> Trim(FrmSchedadd.CBTeacher.Text) Then
        Return_Message = "Phase 7: Conflict detected with " & _
            .Fields!Teacher & vbNewLine & _
            .Fields!TIN & " to " & .Fields!TOUT
            CheckConflict = False
            Exit Function
        End If
    Case "TEACHER"
        If Trim(.Fields!room) <> Trim(FrmSchedadd.cbroom.Text) Then
        Return_Message = "Conflict detected with " & _
            .Fields!room & vbNewLine & _
            .Fields!TIN & " to " & .Fields!TOUT
            CheckConflict = False
            Exit Function
        End If
    End Select
    End If
End With
CheckConflict = True

'False error true no error
Exit Function
CheckConflictErr:
    MainModule_Exception
    CheckConflict = False
End Function

Public Function comparetemp(ByVal rs As Recordset) As Boolean
Dim mycomp As Boolean
mycomp = True
With rs
    If InStr(1, .Fields!sn, SchedModule.TEMP_CLASS, vbTextCompare) = 0 Then
        mycomp = False
    End If
    If SchedModule.TEMP_DAY <> .Fields!day Then
        mycomp = False
    End If
    If TimeValue(SchedModule.TEMP_IN) <> TimeValue(.Fields!TIN) Then
        mycomp = False
    End If
    If TimeValue(SchedModule.TEMP_OUT) <> TimeValue(.Fields!TOUT) Then
        mycomp = False
    End If
    If SchedModule.TEMP_TEACH <> .Fields!Teacher Then
        mycomp = False
    End If
    If SchedModule.TEMP_ROOM <> .Fields!room Then
        mycomp = False
    End If
    If SchedModule.TEMP_SUB <> .Fields!sc Then
        mycomp = False
    End If
End With
comparetemp = mycomp
End Function
