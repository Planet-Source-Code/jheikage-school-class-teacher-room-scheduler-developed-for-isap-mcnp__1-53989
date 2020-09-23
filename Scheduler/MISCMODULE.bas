Attribute VB_Name = "MISCMODULE"
Public rs_MyLoadings As New Recordset
Public rs_Loader As New Recordset
Public MyActive_File As String
Public Function whatLoad(CB As ComboBox, mx As String)
With rs_MyLoadings
    If .State <> 0 Then .Close
    .ActiveConnection = Denver.MyCON
    .CursorLocation = adUseClient
    If UCase(mx) = "TEACHERS" Then mx = "select teacher from teachers group by teacher"
    Call .Open(mx, , adOpenDynamic, adLockOptimistic)
    CB.Clear
    If .RecordCount <> 0 Then
    Do
        CB.AddItem .Fields(0).Value
        .MoveNext
    Loop Until .EOF
    End If
End With
End Function

Public Function Loaders(day As String, SCT As String, cb1 As ComboBox, cb2 As ComboBox)
    With rs_Loader
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .ActiveConnection = Denver.MyCON
        Dim Sql As String, i As Integer
        For i = 0 To 28
        Select Case LCase(day)
            Case "m"
                MainForm.MN(i).BackColor = vbWhite
                MainForm.MN(i).Height = 240
                MainForm.MN(i).Caption = "."
                MainForm.MN(i).Visible = True
            Case "t"
                MainForm.Tu(i).BackColor = vbWhite
                MainForm.Tu(i).Height = 240
                MainForm.Tu(i).Caption = "."
                MainForm.Tu(i).Visible = True
            Case "w"
                MainForm.WD(i).BackColor = vbWhite
                MainForm.WD(i).Height = 240
                MainForm.WD(i).Caption = "."
                MainForm.WD(i).Visible = True
            Case "th"
                MainForm.TH(i).BackColor = vbWhite
                MainForm.TH(i).Height = 240
                MainForm.TH(i).Caption = "."
                MainForm.TH(i).Visible = True
            Case "f"
                MainForm.FR(i).BackColor = vbWhite
                MainForm.FR(i).Height = 240
                MainForm.FR(i).Caption = "."
                MainForm.FR(i).Visible = True
            Case "s"
                MainForm.SAT(i).BackColor = vbWhite
                MainForm.SAT(i).Height = 240
                MainForm.SAT(i).Caption = "."
                MainForm.SAT(i).Visible = True
        End Select
        Next
        Sql = "Select * From Schedules where Day = " & Quoted(day) & " and Sn like '%" & SCT & "%'"
        .Open Sql
        If .RecordCount <> 0 Then
            Do
             
                For i = 0 To cb1.ListCount - 1
                    cb1.ListIndex = i
                    If TimeValue(cb1.Text) = .Fields!TIN Then
                        cb1.ListIndex = i
                        Exit For
                    End If
                Next
                
                For i = 0 To cb2.ListCount - 1
                    cb2.ListIndex = i
                    If TimeValue(cb2.Text) = .Fields!TOUT Then
                        cb2.ListIndex = i
                        Exit For
                    End If
                Next
                Create_BoxHeight cb1.ListIndex, cb2.ListIndex, day, .Fields!sc, cb1.Text, cb2.Text, .Fields!room, .Fields!Teacher
                .MoveNext
            Loop Until .EOF
        End If
    End With
End Function

Public Function OtherLoader(Sql As String, day As String, cb1 As ComboBox, cb2 As ComboBox)
With rs_Loader
If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .ActiveConnection = Denver.MyCON
        Dim i As Integer
        For i = 0 To 28
        Select Case LCase(day)
            Case "m"
                MainForm.MN(i).BackColor = vbWhite
                MainForm.MN(i).Height = 240
                MainForm.MN(i).Caption = "."
                MainForm.MN(i).Visible = True
            Case "t"
                MainForm.Tu(i).BackColor = vbWhite
                MainForm.Tu(i).Height = 240
                MainForm.Tu(i).Caption = "."
                MainForm.Tu(i).Visible = True
            Case "w"
                MainForm.WD(i).BackColor = vbWhite
                MainForm.WD(i).Height = 240
                MainForm.WD(i).Caption = "."
                MainForm.WD(i).Visible = True
            Case "th"
                MainForm.TH(i).BackColor = vbWhite
                MainForm.TH(i).Height = 240
                MainForm.TH(i).Caption = "."
                MainForm.TH(i).Visible = True
            Case "f"
                MainForm.FR(i).BackColor = vbWhite
                MainForm.FR(i).Height = 240
                MainForm.FR(i).Caption = "."
                MainForm.FR(i).Visible = True
            Case "s"
                MainForm.SAT(i).BackColor = vbWhite
                MainForm.SAT(i).Height = 240
                MainForm.SAT(i).Caption = "."
                MainForm.SAT(i).Visible = True
        End Select
        Next
        .Open Sql, , adOpenDynamic, adLockOptimistic
        If .RecordCount <> 0 Then
            Do
             
                For i = 0 To cb1.ListCount - 1
                    cb1.ListIndex = i
                    If TimeValue(cb1.Text) = .Fields!TIN Then
                        cb1.ListIndex = i
                        Exit For
                    End If
                Next
                
                For i = 0 To cb2.ListCount - 1
                    cb2.ListIndex = i
                    If TimeValue(cb2.Text) = .Fields!TOUT Then
                        cb2.ListIndex = i
                        Exit For
                    End If
                Next
                Create_BoxHeight cb1.ListIndex, cb2.ListIndex, day, .Fields!sc, cb1.Text, cb2.Text, .Fields!room, .Fields!sn
                .MoveNext
            Loop Until .EOF
        End If
End With
End Function

Public Function loadSubjectUnits(Lbox As ListBox, SCT As String, yr As String, cb1 As ComboBox, cb2 As ComboBox, mode As Integer)
On Error GoTo LSU:
Dim MyTotHour As Integer
Dim MyCom As String
Dim MyUnt As Double, Meetings As Double, i As Integer
Dim MyCRT As New Recordset
With rs_Loader
    Lbox.Clear
    If .State <> 0 Then .Close
    .ActiveConnection = Denver.MyCON
    .CursorLocation = adUseClient
    .Open "Select * From Subjects where course = " & Quoted(SCT) & " and Yr = " & Quoted(yr)
    If .RecordCount <> 0 Then
        Do
            If MyCRT.State <> 0 Then MyCRT.Close
            MyCRT.ActiveConnection = Denver.MyCON
            MyCRT.CursorLocation = adUseClient
            MyCRT.Open "Select * FRom Schedules where sn like" & Quoted("%" & SCT & "%") & " and Sc = " & Quoted(.Fields!sc), , adOpenDynamic, adLockOptimistic
            Select Case mode
            Case 0
                'regular
                Meetings = ((.Fields!Units - 3) * 3) + 3
            Case 1
                'summer
                Meetings = (((.Fields!Units - 3) * 3) + 3) * 3
            End Select
            MyUnt = 0
            If MyCRT.RecordCount <> 0 Then
                Do
                    For i = 0 To 28
                        cb1.ListIndex = i
                        If TimeValue(cb1.Text) = TimeValue(MyCRT.Fields!TIN) Then GoTo X1
                    Next
X1:
                    For i = 0 To 28
                        cb2.ListIndex = i
                        If TimeValue(cb2.Text) = TimeValue(MyCRT.Fields!TOUT) Then GoTo x2
                    Next
x2:
                    'get their difference
                    MyUnt = MyUnt + (cb2.ListIndex - cb1.ListIndex) / 2
                    MyCRT.MoveNext
                Loop Until MyCRT.EOF
                If MyUnt > Meetings Then
                    MyCom = "Excess Time! " & .Fields!sc & "-Total Meeting/s:" & MyUnt & " hours. Needs " & Meetings & " hours."
                Else
                    MyCom = .Fields!sc & "-Total Meeting/s:" & MyUnt & " hours. Needs " & Meetings & " hours."
                End If
            Else
                MyCom = .Fields!sc & "Needs " & Meetings & " hours."
            End If
            Lbox.AddItem MyCom
            .MoveNext
        Loop Until .EOF
    End If
End With
Exit Function
LSU:
MainModule_Exception

End Function
