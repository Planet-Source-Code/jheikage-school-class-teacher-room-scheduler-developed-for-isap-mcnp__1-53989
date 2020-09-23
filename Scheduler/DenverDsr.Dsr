VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} Denver 
   ClientHeight    =   8700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10695
   _ExtentX        =   18865
   _ExtentY        =   15346
   FolderFlags     =   1
   TypeLibGuid     =   "{810EB81F-3C8F-456E-8DA3-82EA6ACB8DA0}"
   TypeInfoGuid    =   "{7496F53F-23DA-4E11-B28A-59C30280B2F3}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "MyCON"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   $"DenverDsr.dsx":0000
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   12
   BeginProperty Recordset1 
      CommandName     =   "Subjects"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * From subjects"
      ActiveConnectionName=   "MyCON"
      CommandType     =   1
      GroupingName    =   "Subjects_Grouping"
      RelateToParent  =   -1  'True
      ParentCommandName=   "SubHead"
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "SC"
         Caption         =   "SC"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   202
         Name            =   "Course"
         Caption         =   "Course"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   202
         Name            =   "Yr"
         Caption         =   "Yr"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "Desc"
         Caption         =   "Desc"
      EndProperty
      BeginProperty Field5 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "Units"
         Caption         =   "Units"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   2
      BeginProperty Relation1 
         ParentField     =   "Course"
         ChildField      =   "Course"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "Yr"
         ChildField      =   "Yr"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "SubHead"
      CommDispId      =   1010
      RsDispId        =   1033
      CommandText     =   "SELECT Course, Yr FROM Subjects GROUP BY Course, yr ORDER BY course,yr"
      ActiveConnectionName=   "MyCON"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   202
         Name            =   "Course"
         Caption         =   "Course"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   202
         Name            =   "Yr"
         Caption         =   "Yr"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "SubSections"
      CommDispId      =   1034
      RsDispId        =   1044
      CommandText     =   "Sections"
      ActiveConnectionName=   "MyCON"
      CommandType     =   2
      dbObjectType    =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "SectionName"
         Caption         =   "SectionName"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "Teachers"
      CommDispId      =   1045
      RsDispId        =   1054
      CommandText     =   "Select Teacher From Teachers group by Teacher"
      ActiveConnectionName=   "MyCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   0
         Size            =   70
         Scale           =   0
         Type            =   202
         Name            =   "Teacher"
         Caption         =   "Teacher"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset5 
      CommandName     =   "SubTeachers"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * From Teachers"
      ActiveConnectionName=   "MyCON"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "Teachers"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   70
         Scale           =   0
         Type            =   202
         Name            =   "Teacher"
         Caption         =   "Teacher"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   202
         Name            =   "Subject"
         Caption         =   "Subject"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "Teacher"
         ChildField      =   "Teacher"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset6 
      CommandName     =   "CmdRooms"
      CommDispId      =   1055
      RsDispId        =   1058
      CommandText     =   "Select * from Rooms"
      ActiveConnectionName=   "MyCON"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Room"
         Caption         =   "Room"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset7 
      CommandName     =   "HEADSt"
      CommDispId      =   1059
      RsDispId        =   1065
      CommandText     =   "Select * From SORTER order by Numbers"
      ActiveConnectionName=   "MyCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Numbers"
         Caption         =   "Numbers"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "sorter"
         Caption         =   "sorter"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset8 
      CommandName     =   "SubSt"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * From Schedules"
      ActiveConnectionName=   "MyCON"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "HEADSt"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   202
         Name            =   "Day"
         Caption         =   "Day"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "SN"
         Caption         =   "SN"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "SC"
         Caption         =   "SC"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "Descs"
         Caption         =   "Descs"
      EndProperty
      BeginProperty Field5 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "Units"
         Caption         =   "Units"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "Attribs"
         Caption         =   "Attribs"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "TIN"
         Caption         =   "TIN"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "TOUT"
         Caption         =   "TOUT"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   70
         Scale           =   0
         Type            =   202
         Name            =   "Teacher"
         Caption         =   "Teacher"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Room"
         Caption         =   "Room"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "sorter"
         ChildField      =   "Day"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset9 
      CommandName     =   "HEAD_ROOMS"
      CommDispId      =   1066
      RsDispId        =   1077
      CommandText     =   "Select * From Sorter order by numbers"
      ActiveConnectionName=   "MyCON"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Numbers"
         Caption         =   "Numbers"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "sorter"
         Caption         =   "sorter"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset10 
      CommandName     =   "ROOM_CONNECT"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * From Schedules"
      ActiveConnectionName=   "MyCON"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "HEAD_ROOMS"
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   202
         Name            =   "Day"
         Caption         =   "Day"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "SN"
         Caption         =   "SN"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "SC"
         Caption         =   "SC"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "Descs"
         Caption         =   "Descs"
      EndProperty
      BeginProperty Field5 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "Units"
         Caption         =   "Units"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "Attribs"
         Caption         =   "Attribs"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "TIN"
         Caption         =   "TIN"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "TOUT"
         Caption         =   "TOUT"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   70
         Scale           =   0
         Type            =   202
         Name            =   "Teacher"
         Caption         =   "Teacher"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Room"
         Caption         =   "Room"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "sorter"
         ChildField      =   "Day"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset11 
      CommandName     =   "Sched_List_Grouping"
      CommDispId      =   1078
      RsDispId        =   1088
      CommandText     =   $"DenverDsr.dsx":008F
      ActiveConnectionName=   "MyCON"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   202
         Name            =   "Sc"
         Caption         =   "Sc"
      EndProperty
      BeginProperty Field2 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "Units"
         Caption         =   "Units"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Descs"
         Caption         =   "Descs"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset12 
      CommandName     =   "Sched_List"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * From REPORTTABLE"
      ActiveConnectionName=   "MyCON"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "Sched_List_Grouping"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "SN"
         Caption         =   "SN"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   202
         Name            =   "SC"
         Caption         =   "SC"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "DESCS"
         Caption         =   "DESCS"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   202
         Name            =   "DAY"
         Caption         =   "DAY"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "SCHEDULE"
         Caption         =   "SCHEDULE"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   202
         Name            =   "TEACHER"
         Caption         =   "TEACHER"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "REMARKS"
         Caption         =   "REMARKS"
      EndProperty
      BeginProperty Field8 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "UNITS"
         Caption         =   "UNITS"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   3
      BeginProperty Relation1 
         ParentField     =   "Units"
         ChildField      =   "UNITS"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "Descs"
         ChildField      =   "DESCS"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation3 
         ParentField     =   "Sc"
         ChildField      =   "SC"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "Denver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
