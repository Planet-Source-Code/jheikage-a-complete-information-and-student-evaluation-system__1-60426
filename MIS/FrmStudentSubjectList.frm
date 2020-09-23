VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStudentSubjectList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Subject List"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmStudentSubjectList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnRetrive 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Retrive"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   1935
   End
   Begin VB.ListBox LstSY 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   8400
      TabIndex        =   16
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Course"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   6855
      Begin VB.ComboBox TxtYr 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmStudentSubjectList.frx":030A
         Left            =   2760
         List            =   "FrmStudentSubjectList.frx":031D
         TabIndex        =   3
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox TxtCourse 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   240
         TabIndex        =   2
         Text            =   "Course"
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox ChkAll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Show All"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3480
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton BtnLoad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Load Subjects"
      Height          =   975
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton BtnPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Print"
      Height          =   735
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton BtnDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Delete"
      Height          =   735
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton BtnEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Edit"
      Height          =   735
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton BtnNew 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&New"
      Height          =   735
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin MSComctlLib.ListView LvSubjects 
      Height          =   3375
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Year"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Units"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Grade"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Remarks"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "School Year Information"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   8055
      Begin VB.ComboBox TxtSy 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmStudentSubjectList.frx":0330
         Left            =   1440
         List            =   "FrmStudentSubjectList.frx":0361
         TabIndex        =   0
         Text            =   "9999-9999"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox TxtSem 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmStudentSubjectList.frx":040A
         Left            =   3720
         List            =   "FrmStudentSubjectList.frx":0417
         TabIndex        =   1
         Text            =   "Sem"
         Top             =   360
         Width           =   975
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
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1125
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
         Left            =   2760
         TabIndex        =   12
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Records Found at"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8400
      TabIndex        =   17
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Menu SubjectMove 
      Caption         =   "Move Subjects"
      Visible         =   0   'False
      Begin VB.Menu MnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu MnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu break1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMove 
         Caption         =   "&Move This Subject to Another SY/SEM"
      End
   End
End
Attribute VB_Name = "FrmStudentSubjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   College Record Module
'Description    :   Displays records of students pertaining their subjects taken
'                   but you can only view 1 school year at a given semester and course
'                   at a time. A summary of all the school year in which the
'                   active student is active or enrolled is seen in the retreive box.
'Date           :   January 20, 2005
'Programmer     :   iehjsuckers
'Comment        :   ?
'************************************************************************************
Public Lnam As String, Mnam As String, Fnam As String, Idnum As String
Public isadd As Boolean
Public Course As String, yr As String

Private Sub BtnDelete_Click()
If Me.LvSubjects.SelectedItem Is Nothing Then Exit Sub
'Delete here
    FrmConfirm.Show 1
    If IsTrue = False Then
        MsgBox "You do not have the permission to delete files.", vbCritical, "Error"
        Exit Sub
    End If

If MsgBox("do you really want to delete this record?", vbYesNo, "Confirm") = vbYes Then
    Dim msg As String
Dim spl
    spl = Split(Me.Caption, "|", , vbTextCompare)
    Idnum = spl(1)
    Lnam = spl(2)
    Fnam = spl(3)
    Mnam = spl(4)
    msg = "Delete from SubjectsEnrolled where SY = '" & TxtSy.Text & "' and sem = '" & TxtSem.Text & "' and course='" & TxtCourse.Text & "' and yr = '" & TxtYr.Text & "' and sc = '" & LvSubjects.SelectedItem.SubItems(1) & "' and lnam = '" & Lnam & "' and mnam = '" & Mnam & "' and fnam = '" & Fnam & "' and idnum = '" & Idnum & "'"
    SetSelectedStudSub msg
    'reload
    BtnLoad_Click
End If
End Sub

Private Sub BtnEdit_Click()
If Me.LvSubjects.SelectedItem Is Nothing Then Exit Sub
Me.isadd = False
Dim spl
spl = Split(Me.Caption, "|", , vbTextCompare)
Idnum = Trim(spl(1))
Lnam = Trim(spl(2))
Fnam = Trim(spl(3))
Mnam = Trim(spl(4))
'Set the sy and sem
Dim msg As String
msg = "Select * from SubjectsEnrolled where SY='" & TxtSy.Text & "' and SEM = '" & TxtSem.Text & _
    "' and Course = '" & TxtCourse.Text & "' and lnam = '" & Lnam & "' and fnam = '" & Fnam & "' and mnam = '" & Mnam & "' and sc = '" & LvSubjects.SelectedItem.SubItems(1) & "'"
SetSelectedStudSub msg
FrmColRecAdd.Caption = Me.Caption
FrmColRecAdd.Show 1
msg = "Select * from SubjectsEnrolled where SY='" & TxtSy.Text & "' and SEM = '" & TxtSem.Text & _
    "' and Course = '" & TxtCourse.Text & "' and lnam = '" & Lnam & "' and fnam = '" & Fnam & "' and mnam = '" & Mnam & "'"
    If ChkAll.Value <> vbChecked Then
        msg = msg & " and yr = '" & TxtYr.Text & "'"
    End If
SetSelectedStudSub msg
loadSLtoLv LvSubjects
End Sub

Private Sub BtnLoad_Click()
'view records
Dim spl
spl = Split(Me.Caption, "|", , vbTextCompare)
Idnum = spl(1)
Lnam = spl(2)
Fnam = spl(3)
Mnam = spl(4)
'Set the sy and sem
Dim msg As String
msg = "Select * from SubjectsEnrolled where SY='" & TxtSy.Text & "' and SEM = '" & TxtSem.Text & _
    "' and Course = '" & TxtCourse.Text & "' and lnam = '" & Lnam & "' and fnam = '" & Fnam & "' and mnam = '" & Mnam & "'"
    If ChkAll.Value <> vbChecked Then
        msg = msg & " and yr = '" & TxtYr.Text & "'"
    End If
SetSelectedStudSub msg
loadSLtoLv LvSubjects
End Sub

Private Sub BtnNew_Click()
'Try spliting the caption
Dim spl
Me.isadd = True
spl = Split(Me.Caption, "|", , vbTextCompare)
Idnum = spl(1)
Lnam = spl(2)
Fnam = spl(3)
Mnam = spl(4)
Course = TxtCourse.Text
yr = TxtYr.Text
FrmColRecAdd.Show 1
Dim msg As String
msg = "Select * from SubjectsEnrolled where SY='" & TxtSy.Text & "' and SEM = '" & TxtSem.Text & _
    "' and Course = '" & TxtCourse.Text & "' and lnam = '" & Lnam & "' and fnam = '" & Fnam & "' and mnam = '" & Mnam & "'"
    If ChkAll.Value <> vbChecked Then
        msg = msg & " and yr = '" & TxtYr.Text & "'"
    End If
SetSelectedStudSub msg
loadSLtoLv LvSubjects
End Sub

Private Sub BtnPrint_Click()
'updated on march 17, 2005
Dim myMx As String
Dim spl
    spl = Split(Me.Caption, "|", , vbTextCompare)
    Idnum = spl(1)
    Lnam = spl(2)
    Fnam = spl(3)
    Mnam = spl(4)
myMx = "SHAPE {Select SY,SEM,COURSE,YR from subjectsenrolled WHERE lnam='" & Lnam & "' and fnam ='" & Fnam & "' and mnam = '" & Mnam & "' group by SY,SEM,COURSE,YR}  AS PrintSubjects APPEND ({Select * from SubjectsEnrolled WHERE lnam='" & Lnam & "' and fnam ='" & Fnam & "' and mnam = '" & Mnam & "'}  AS TheSubs RELATE 'SY' TO 'SY','SEM' TO 'SEM','COURSE' TO 'Course','YR' TO 'Yr') AS TheSubs"
With DE.rsPrintSubjects
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .Open myMx
End With
With DrepSubs
    .Sections("PageHeader").Controls("LblName").Caption = UCase(Lnam) & ", " & UCase(Fnam) & " " & UCase(Mnam)
    .Sections("PageHeader").Controls("LblIDNUM").Caption = Idnum
    DrepSubs.Show 1
End With
End Sub

Private Sub BtnRetrive_Click()
GroupSYSem
End Sub

Private Sub Form_Load()
GetCourse
End Sub

'User Functions
Public Function GroupSYSem()
    On Error GoTo bx
    Dim rs As New Recordset
    Dim spl
    spl = Split(Me.Caption, "|", , vbTextCompare)
    Idnum = spl(1)
    Lnam = spl(2)
    Fnam = spl(3)
    Mnam = spl(4)
    
    'clear data
    LstSY.Clear
    With rs
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open "Select SY,SEM,course,yr from SubjectsEnrolled enrolled where lnam = '" & Lnam & "' and fnam = '" & Fnam & "' and mnam = '" & Mnam & "' Group by SY,SEM,course,yr ", DE.Con
        If .RecordCount > 0 Then
            'do here
            Do Until .EOF
                LstSY.AddItem .Fields("SY").Value & " - " & .Fields("SEM").Value & " " & .Fields("Course").Value & " " & .Fields("yr").Value
                .MoveNext
            Loop
        End If
    End With
    Set rs = Nothing
    Exit Function
bx:
    MsgBox Err.Description, vbCritical, "Error"
    Set rs = Nothing
End Function
Private Sub GetCourse()
    With DE.rsCourses
        If .State <> 0 Then .Close
        .Open "Select * from courses"
        'clear the combobox
        TxtCourse.Clear
        If .RecordCount > 0 Then
            Do Until .EOF
                TxtCourse.AddItem .Fields(0).Value
                .MoveNext
            Loop
        End If
    End With
End Sub


Private Sub LstSY_Click()
LstSY.ToolTipText = LstSY.Text
End Sub

'End of user functions

Private Sub LvSubjects_DblClick()
    If LvSubjects.SelectedItem Is Nothing Then Exit Sub
    BtnEdit_Click
End Sub

Private Sub LvSubjects_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If LvSubjects.SelectedItem Is Nothing Then Exit Sub
    If Button = 2 Then
        Me.PopupMenu SubjectMove
    End If
End Sub

Private Sub MnuDelete_Click()
    BtnDelete_Click
End Sub

Private Sub MnuEdit_Click()
    BtnEdit_Click
End Sub

Private Sub MnuMove_Click()
    'Try moving
    If Me.LvSubjects.SelectedItem Is Nothing Then Exit Sub
'Delete here
    FrmConfirm.Show 1
    If IsTrue = False Then
        MsgBox "You do not have the permission to move the record. Please contact your administrator.", vbCritical, "Error"
        Exit Sub
    End If
    With FrmMoving
        .LblStudent.Caption = Me.Caption
        .LblSC.Caption = LvSubjects.SelectedItem.SubItems(1)
        .LblDesc.Caption = LvSubjects.SelectedItem.SubItems(2)
        .LblUnits.Caption = LvSubjects.SelectedItem.SubItems(3)
        .LblGrade.Caption = LvSubjects.SelectedItem.SubItems(4)
        .LblRemarks.Caption = LvSubjects.SelectedItem.SubItems(5)
        .LblSY.Caption = TxtSy.Text
        .LblSem.Caption = TxtSem.Text
        .Show 1
    End With
    BtnLoad_Click
End Sub

Private Sub MnuNew_Click()
    BtnNew_Click
End Sub

Private Sub MnuPrint_Click()
    BtnPrint_Click
End Sub
