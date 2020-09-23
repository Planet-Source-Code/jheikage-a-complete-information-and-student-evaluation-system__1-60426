VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmReports 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Reports"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6930
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
   ScaleHeight     =   4485
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Rbox 
      Height          =   855
      Left            =   5520
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"FrmReports.frx":0000
   End
   Begin TabDlg.SSTab MyTabs 
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   2566
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483644
      OLEDropMode     =   1
      TabCaption(0)   =   "Summary Reports (Course)"
      TabPicture(0)   =   "FrmReports.frx":0077
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TxtYr"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TxtCourse"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ChkAllYear"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "BtnCreateCourse"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Summary Reports (Department)"
      TabPicture(1)   =   "FrmReports.frx":0093
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "BtnCreateDep"
      Tab(1).Control(1)=   "Opt2"
      Tab(1).Control(2)=   "Opt1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Lists"
      TabPicture(2)   =   "FrmReports.frx":00AF
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtCourse1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "TxtYr1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "BtnCreateList"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "ListOpt1(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "ListOpt1(1)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "ListOpt1(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "ListOpt1(3)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.OptionButton ListOpt1 
         Caption         =   "Report All"
         Height          =   255
         Index           =   3
         Left            =   -72600
         TabIndex        =   28
         ToolTipText     =   "Selects all students"
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton ListOpt1 
         Caption         =   "Report Select Year Level"
         Height          =   255
         Index           =   2
         Left            =   -72600
         TabIndex        =   27
         ToolTipText     =   "Selects all students within the selected year level regardless of their courses"
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton ListOpt1 
         Caption         =   "Report Selected Course"
         Height          =   255
         Index           =   1
         Left            =   -72600
         TabIndex        =   26
         ToolTipText     =   "Reports all students within the selected course regardless of their year level"
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton ListOpt1 
         Caption         =   "Report Selected"
         Height          =   255
         Index           =   0
         Left            =   -72600
         TabIndex        =   25
         ToolTipText     =   "Selects All students within the selected Course and Year Level"
         Top             =   120
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton BtnCreateList 
         Caption         =   "Create"
         Height          =   495
         Left            =   -69360
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox TxtYr1 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmReports.frx":00CB
         Left            =   -73200
         List            =   "FrmReports.frx":00DE
         TabIndex        =   23
         Text            =   "1"
         Top             =   120
         Width           =   615
      End
      Begin VB.ComboBox TxtCourse1 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -74880
         TabIndex        =   22
         Text            =   "Course"
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton BtnCreateDep 
         Caption         =   "Create"
         Height          =   495
         Left            =   -69360
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton BtnCreateCourse 
         Caption         =   "Create"
         Height          =   495
         Left            =   5640
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Opt2 
         Appearance      =   0  'Flat
         Caption         =   "By Year Level"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74760
         TabIndex        =   13
         Top             =   480
         Width           =   2655
      End
      Begin VB.OptionButton Opt1 
         Appearance      =   0  'Flat
         Caption         =   "Entire Department (CAS)"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74760
         TabIndex        =   12
         Top             =   120
         Width           =   2655
      End
      Begin VB.CheckBox ChkAllYear 
         Appearance      =   0  'Flat
         Caption         =   "Report Entire Course"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   120
         Width           =   2175
      End
      Begin VB.ComboBox TxtCourse 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Text            =   "Course"
         Top             =   120
         Width           =   2415
      End
      Begin VB.ComboBox TxtYr 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmReports.frx":00F1
         Left            =   2640
         List            =   "FrmReports.frx":0104
         TabIndex        =   10
         Text            =   "1"
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Include the Following to the Report"
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   6735
      Begin VB.CheckBox ChkIncome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Income (Gross)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox ChkAge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Age"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox ChkBhouse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Boarding House"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox ChkFoccup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Fathers Occupation"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox ChkMOccup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Mothers Occupation"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox ChkCS 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Civil Status"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox ChkSex 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Sex"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.ComboBox TxtSy 
      Appearance      =   0  'Flat
      Height          =   360
      ItemData        =   "FrmReports.frx":0117
      Left            =   1320
      List            =   "FrmReports.frx":0148
      TabIndex        =   0
      Text            =   "9999-9999"
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox TxtSem 
      Appearance      =   0  'Flat
      Height          =   360
      ItemData        =   "FrmReports.frx":01F1
      Left            =   3600
      List            =   "FrmReports.frx":01FE
      TabIndex        =   1
      Text            =   "Sem"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Summary Report Wizard."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   3780
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Semester:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   2640
      TabIndex        =   19
      Top             =   480
      Width           =   900
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "School Year:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   1125
   End
End
Attribute VB_Name = "FrmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   MIS Report Module
'Description    :   Generates important reports such as count of students per school
'                   year, semester, course and by year level.
'Date           :   February 26, 2005
'Programmer     :   iehjsuckers
'Comment        :   ?
'************************************************************************************
Private Sub BtnCreateCourse_Click()
    CreateSummary_For_Courses txtSY.Text, txtSem.Text, txtCourse.Text
End Sub

Private Sub BtnCreateDep_Click()
If Opt1.Value = True Then Me.CreateSummary_For_CAS txtSY.Text, txtSem.Text
If Opt2.Value = True Then Me.CreateSummary_For_CAS_Per_YEAR txtSY.Text, txtSem.Text
End Sub

Private Sub btnCreateList_Click()
'select option
Dim msg As String, comment As String
If ListOpt1(0).Value = True Then
msg = "Select IDNUM,LNAM,FNAM,MNAM from subjectsenrolled where sy = '" & txtSY.Text & "' and sem = '" & txtSem.Text & "' and course = '" & TxtCourse1.Text & "' and yr = '" & TxtYr1.Text & "' group by IDNUM,MNAM, LNAM, FNAM"
comment = "Course: " & TxtCourse1.Text & " Year Level: " & TxtYr1.Text
End If
If ListOpt1(1).Value = True Then
msg = "Select IDNUM,LNAM,FNAM,MNAM from subjectsenrolled where sy = '" & txtSY.Text & "' and sem = '" & txtSem.Text & "' and course = '" & TxtCourse1.Text & "' group by IDNUM,MNAM, LNAM, FNAM"
comment = "Course: " & TxtCourse1.Text
End If
If ListOpt1(2).Value = True Then
msg = "Select IDNUM,LNAM,FNAM,MNAM from subjectsenrolled where sy = '" & txtSY.Text & "' and sem = '" & txtSem.Text & "' and yr = '" & TxtYr1.Text & "' group by IDNUM,MNAM, LNAM, FNAM"
comment = " Year Level: " & TxtYr1.Text
End If
If ListOpt1(3).Value = True Then
msg = "Select IDNUM,LNAM,FNAM,MNAM from subjectsenrolled where sy = '" & txtSY.Text & "' and sem = '" & txtSem.Text & "' group by IDNUM,MNAM, LNAM, FNAM"
comment = "All Students for the specified school year and semester"
End If
With DE.rsAllStudes
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .Open msg, DE.Con, adOpenDynamic, adLockOptimistic
End With
With DrepStudList
    .Sections("PageHeader").Controls("LblSY").Caption = txtSY.Text
    .Sections("PageHeader").Controls("LblSem").Caption = txtSem.Text
    .Sections("PageHeader").Controls("LblComment").Caption = comment
    .Show 1
End With
End Sub

Private Sub Form_Load()
getcourses
End Sub

'user defined functions
Function getcourses()
    txtCourse.Clear
    TxtCourse1.Clear
    With DE.rsCourses
        If .State <> 0 Then .Close
        .Open "Select * from courses"
        'clear the combobox
        txtCourse.Clear
        If .RecordCount > 0 Then
            Do Until .EOF
                txtCourse.AddItem .Fields(0).Value
                TxtCourse1.AddItem .Fields(0).Value
                .MoveNext
            Loop
        End If
    End With
End Function
'***********************************************************
'***********Ang mga functions na ito ay para sa paggawa ng**
'***********summary report sa bawat kurso ng CAS************
'***********************************************************
Function CreateSummary_For_Courses(SY As String, SEM As String, Course As String)
'It will output an html formatted file: filename is Report.html
'try deletingthe file RePort.html
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim MyReport As String  'the report
ChDir App.Path 'Force path
If Right(App.Path, 1) = "\" Then
    MyReport = App.Path & "report.html"
Else
    MyReport = App.Path & "\report.html"
End If
If fso.fileexists(MyReport) Then Kill MyReport
Dim mymes As String
mymes = mymes & "<html><header><title>Report for " & Course & "</title></header>"
mymes = mymes & "<style>"
mymes = mymes & ".Heads{Font:Bold 14pt Arial;Color:Black};"
mymes = mymes & ".Myx{Font:Bold 12pt Arial;Color:Black};"
mymes = mymes & "</style>"
mymes = mymes & "<body>"
mymes = mymes & "<span style='Font:Bold 16pt Arial;width:100%;text-Align:center'>SUMMARY REPORT FOR " & UCase(Course) & "<BR>FOR SCHOOL YEAR " & SY & " " & SEM & " SEMESTER" & "</span>"

Dim rs As New Recordset
With rs
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .CursorType = adOpenDynamic
    
Dim Sql As String
'check if you need to get age, also, if the check box for all year level
'******* ******** *********
'*     * *        *
'******* *  ***** *********
'*     * *      * *
'*     * ******** *********
If ChkAge.Value = vbChecked Then
Sql = "SELECT COUNT(AGE) AS Total, AGE " & _
    "FROM NAMES INNER JOIN (SELECT SubjectsEnrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem, yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and course = '" & Course & "' "
    If ChkAllYear.Value <> vbChecked Then    'report specified year level
        Sql = Sql & " and inx.YR = '" & txtYr.Text & "'"
        mymes = mymes & "<span style='Font:Bold 13pt Arial;width:100%;text-Align:center'>Year Level:" & txtYr.Text & "</span>"
        'write headers here
    End If
    Sql = Sql & " GROUP BY AGE"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>AGE</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("AGE").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If

'SEX HERE
If ChkSex.Value = vbChecked Then
Sql = "SELECT COUNT(SEX) AS Total, SEX " & _
    "FROM NAMES INNER JOIN (SELECT SubjectsEnrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem, yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and course = '" & Course & "' "
    If ChkAllYear.Value <> vbChecked Then    'report specified year level
        Sql = Sql & " and inx.YR = '" & txtYr.Text & "'"
    End If
    Sql = Sql & " GROUP BY SEX"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>SEX</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("SEX").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    
'Civil Status here
If ChkCS.Value = vbChecked Then
Sql = "SELECT COUNT(CS) AS Total, CS " & _
    "FROM NAMES INNER JOIN (SELECT SubjectsEnrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem, yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and course = '" & Course & "' "
    If ChkAllYear.Value <> vbChecked Then    'report specified year level
        Sql = Sql & " and inx.YR = '" & txtYr.Text & "'"
    End If
    Sql = Sql & " GROUP BY CS"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>CIVIL STATUS</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("CS").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    
If ChkIncome.Value = vbChecked Then
Sql = "SELECT COUNT(INCOME) AS Total, INCOME " & _
    "FROM NAMES INNER JOIN (SELECT SubjectsEnrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem, yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and course = '" & Course & "' "
    If ChkAllYear.Value <> vbChecked Then    'report specified year level
        Sql = Sql & " and inx.YR = '" & txtYr.Text & "'"
    End If
    Sql = Sql & " GROUP BY INCOME"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>INCOME</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("Income").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    
If ChkFoccup.Value = vbChecked Then
Sql = "SELECT COUNT(Foccup) AS Total, Foccup " & _
    "FROM NAMES INNER JOIN (SELECT SubjectsEnrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem, yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and course = '" & Course & "' "
    If ChkAllYear.Value <> vbChecked Then    'report specified year level
        Sql = Sql & " and inx.YR = '" & txtYr.Text & "'"
        'write headers here
        mymes = mymes & "<span style='Font:Bold 13pt Arial;width:100%;text-Align:center'>Year Level:" & txtYr.Text & "</span>"
    End If
    Sql = Sql & " GROUP BY Foccup"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>FATHER'S OCCUPATION</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("Foccup").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    
If ChkMOccup.Value = vbChecked Then
Sql = "SELECT COUNT(MOCCUP) AS Total, MOCCUP " & _
    "FROM NAMES INNER JOIN (SELECT SubjectsEnrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem, yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and course = '" & Course & "' "
    If ChkAllYear.Value <> vbChecked Then    'report specified year level
        Sql = Sql & " and inx.YR = '" & txtYr.Text & "'"
    End If
    Sql = Sql & " GROUP BY MOCCUP"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>MOTHER'S OCCUPATION</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("Moccup").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    
If ChkBhouse.Value = vbChecked Then
Sql = "SELECT COUNT(BOARDING) AS Total, BOARDING " & _
    "FROM NAMES INNER JOIN (SELECT SubjectsEnrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem, yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and course = '" & Course & "' "
    If ChkAllYear.Value <> vbChecked Then    'report specified year level
        Sql = Sql & " and inx.YR = '" & txtYr.Text & "'"
    End If
    Sql = Sql & " GROUP BY Boarding"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>BOARDING HOUSE</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("Boarding").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    
    'close the rs
End With
Set rs = Nothing
    mymes = mymes & "</body></html>"

'Shell "iexplore " & MyReport
Rbox.Text = ""
Rbox.Text = mymes
Rbox.SaveFile MyReport, rtfText
Shell "report.bat " & MyReport
Exit Function
Erb:
    MsgBox Err.Description, vbCritical, "Error"
    Set rs = Nothing
End Function

Function CreateSummary_For_CAS_Per_YEAR(SY As String, SEM As String)
'It will output an html formatted file: filename is Report.html
'try deletingthe file RePort.html
Dim yr As Integer
Set fso = CreateObject("Scripting.FileSystemObject")

Dim MyReport As String  'the report
ChDir App.Path 'Force path
If Right(App.Path, 1) = "\" Then
    MyReport = App.Path & "report.html"
Else
    MyReport = App.Path & "\report.html"
End If
If fso.fileexists(MyReport) Then Kill MyReport
Dim mymes As String
mymes = mymes & "<html><header><title>Report for College of Arts and Sciences</title></header>"
mymes = mymes & "<style>"
mymes = mymes & ".Heads{Font:Bold 14pt Arial;Color:Black};"
mymes = mymes & ".Myx{Font:Bold 12pt Arial;Color:Black};"
mymes = mymes & "</style>"
mymes = mymes & "<body>"
For yr = 1 To 4
mymes = mymes & "<span style='Font:Bold 16pt Arial;width:100%;text-Align:center'>SUMMARY REPORT FOR COLLEGE OF ARTS AND SCIENCES<BR>FOR SCHOOL YEAR " & SY & " " & SEM & " SEMESTER" & "</span>"
mymes = mymes & "<span style='Font:Bold 16pt Arial;width:100%;text-Align:center'>YEAR LEVEL:" & yr & "</span>"
Dim rs As New Recordset
With rs
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .CursorType = adOpenDynamic
    
Dim Sql As String
'check if you need to get age, also, if the check box for all year level
'******* ******** *********
'*     * *        *
'******* *  ***** *********
'*     * *      * *
'*     * ******** *********
If ChkAge.Value = vbChecked Then
Sql = "SELECT COUNT(AGE) AS Total, AGE " & _
    "FROM NAMES INNER JOIN (SELECT subjectsenrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem,yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and inx.yr = '" & yr & "' GROUP BY AGE"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>AGE</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("AGE").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If

'SEX HERE
If ChkSex.Value = vbChecked Then
Sql = "SELECT COUNT(SEX) AS Total, SEX " & _
    "FROM NAMES INNER JOIN (SELECT  subjectsenrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem,yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and inx.yr = '" & yr & "' GROUP BY SEX"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>SEX</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("SEX").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    
'Civil Status here
If ChkCS.Value = vbChecked Then
Sql = "SELECT COUNT(CS) AS Total, CS " & _
    "FROM NAMES INNER JOIN (SELECT subjectsenrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem,yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and inx.yr = '" & yr & "' GROUP BY CS"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>CIVIL STATUS</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("CS").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    
If ChkIncome.Value = vbChecked Then
Sql = "SELECT COUNT(INCOME) AS Total, INCOME " & _
    "FROM NAMES INNER JOIN (SELECT subjectsenrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem,yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and inx.yr = '" & yr & "' GROUP BY INCOME"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>INCOME</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("Income").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    
If ChkFoccup.Value = vbChecked Then
Sql = "SELECT COUNT(Foccup) AS Total, Foccup " & _
    "FROM NAMES INNER JOIN (SELECT subjectsenrolled.lnam, subjectsenrolled.yr," & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem,yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and inx.yr = '" & yr & "' GROUP BY Foccup"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>FATHER'S OCCUPATION</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("Foccup").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    
If ChkMOccup.Value = vbChecked Then
Sql = "SELECT COUNT(MOCCUP) AS Total, MOCCUP " & _
    "FROM NAMES INNER JOIN (SELECT subjectsenrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem,yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and inx.yr = '" & yr & "' GROUP BY MOCCUP"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>MOTHER'S OCCUPATION</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("Moccup").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    
If ChkBhouse.Value = vbChecked Then
Sql = "SELECT COUNT(BOARDING) AS Total, BOARDING " & _
    "FROM NAMES INNER JOIN (SELECT subjectsenrolled.yr, subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem,yr) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' and inx.yr = '" & yr & "' GROUP BY BOARDING"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>BOARDING HOUSE</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("Boarding").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table><BR><BR>"
    .Close
End If
    
    'close the rs

End With
Next
Set rs = Nothing
    mymes = mymes & "</body></html>"

'Shell "iexplore " & MyReport
Rbox.Text = ""
Rbox.Text = mymes
Rbox.SaveFile MyReport, rtfText
Shell "report.bat " & MyReport
Exit Function
Erb:
    MsgBox Err.Description, vbCritical, "Error"
    Set rs = Nothing

End Function

Function CreateSummary_For_CAS(SY As String, SEM As String)
'It will output an html formatted file: filename is Report.html
'try deletingthe file RePort.html
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim MyReport As String  'the report
ChDir App.Path 'Force path
If Right(App.Path, 1) = "\" Then
    MyReport = App.Path & "report.html"
Else
    MyReport = App.Path & "\report.html"
End If
If fso.fileexists(MyReport) Then Kill MyReport
Dim mymes As String
mymes = mymes & "<html><header><title>Report for College of Arts and Sciences</title></header>"
mymes = mymes & "<style>"
mymes = mymes & ".Heads{Font:Bold 14pt Arial;Color:Black};"
mymes = mymes & ".Myx{Font:Bold 12pt Arial;Color:Black};"
mymes = mymes & "</style>"
mymes = mymes & "<body>"
mymes = mymes & "<span style='Font:Bold 16pt Arial;width:100%;text-Align:center'>SUMMARY REPORT FOR COLLEGE OF ARTS AND SCIENCES<BR>FOR SCHOOL YEAR " & SY & " " & SEM & " SEMESTER" & "</span>"

Dim rs As New Recordset
With rs
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .CursorType = adOpenDynamic
    
Dim Sql As String
'check if you need to get age, also, if the check box for all year level
'******* ******** *********
'*     * *        *
'******* *  ***** *********
'*     * *      * *
'*     * ******** *********
If ChkAge.Value = vbChecked Then
Sql = "SELECT COUNT(AGE) AS Total, AGE " & _
    "FROM NAMES INNER JOIN (SELECT subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' GROUP BY AGE"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>AGE</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("AGE").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
'SEX HERE
If ChkSex.Value = vbChecked Then
Sql = "SELECT COUNT(SEX) AS Total, SEX " & _
    "FROM NAMES INNER JOIN (SELECT  subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' GROUP BY SEX"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>SEX</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("SEX").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
'Civil Status here
If ChkCS.Value = vbChecked Then
Sql = "SELECT COUNT(CS) AS Total, CS " & _
    "FROM NAMES INNER JOIN (SELECT subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' GROUP BY CS"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>CIVIL STATUS</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("CS").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If

If ChkIncome.Value = vbChecked Then
Sql = "SELECT COUNT(INCOME) AS Total, INCOME " & _
    "FROM NAMES INNER JOIN (SELECT subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' GROUP BY INCOME"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>INCOME</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("Income").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If

If ChkFoccup.Value = vbChecked Then
Sql = "SELECT COUNT(Foccup) AS Total, Foccup " & _
    "FROM NAMES INNER JOIN (SELECT subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' GROUP BY Foccup"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>FATHER'S OCCUPATION</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("Foccup").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
If ChkMOccup.Value = vbChecked Then
Sql = "SELECT COUNT(MOCCUP) AS Total, MOCCUP " & _
    "FROM NAMES INNER JOIN (SELECT subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' GROUP BY MOCCUP"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>MOTHER'S OCCUPATION</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("Moccup").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    
If ChkBhouse.Value = vbChecked Then
Sql = "SELECT COUNT(BOARDING) AS Total, BOARDING " & _
    "FROM NAMES INNER JOIN (SELECT subjectsenrolled.lnam, " & _
    "subjectsenrolled.fnam, subjectsenrolled.mnam,subjectsenrolled.SY , subjectsenrolled.SEM " & _
    "From subjectsenrolled GROUP BY lnam, fnam, mnam, sy, sem) AS inx ON " & _
    "inx.lnam = names.lnam AND inx.fnam = names.fnam AND inx.Mnam = names.Mnam" & _
    " where Sy = '" & SY & "' and SEM = '" & SEM & "' GROUP BY BOARDING"
    .Open Sql, DE.Con
    'create a table to present
    mymes = mymes & "<BR><BR><BR><table align = 'center' cellpadding='0' cellspacing = '0' border='1' style='Font:12pt Arial' width = '80%'>"
    mymes = mymes & "<tr width='100%'><td width='50%' align='center'><B>BOARDING HOUSE</b><td><td width='50%' align='center'><B>TOTAL</b><td></tr>"
    Do Until .EOF
        'write now
       If .Fields("total").Value > 0 Then mymes = mymes & "<tr ><td align='center'>" & .Fields("Boarding").Value & "<td><td align='center'>" & .Fields("Total").Value & "<td></tr>"
        .MoveNext
    Loop
    mymes = mymes & "</table>"
    .Close
End If
    'close the rs
End With
Set rs = Nothing
    mymes = mymes & "</body></html>"

'Shell "iexplore " & MyReport
Rbox.Text = ""
Rbox.Text = mymes
Rbox.SaveFile MyReport, rtfText
Shell "report.bat " & MyReport
Exit Function
Erb:
    MsgBox Err.Description, vbCritical, "Error"
    Set rs = Nothing

End Function
'***********************************************************
'end of user defined functions


