VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCurs 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Curriculum"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCurs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6360
      Width           =   1215
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
      Left            =   6240
      TabIndex        =   14
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton BtnRetrive 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Retrive"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton BtnShow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Load Records"
      Height          =   855
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton BtnDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton BtnAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Add"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   1215
   End
   Begin MSComctlLib.ListView LvCurSub 
      Height          =   3975
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "YR"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject Code"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Units"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Prerequisites"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Course"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   6855
      Begin VB.CheckBox ChkAll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Show All Year Level"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3480
         TabIndex        =   12
         Top             =   360
         Width           =   2175
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
      Begin VB.ComboBox TxtYr 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmCurs.frx":030A
         Left            =   2760
         List            =   "FrmCurs.frx":031D
         TabIndex        =   3
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "School Year Information"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   8055
      Begin VB.ComboBox TxtSem 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmCurs.frx":0330
         Left            =   3720
         List            =   "FrmCurs.frx":033D
         TabIndex        =   1
         Text            =   "Sem"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox TxtSy 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmCurs.frx":0350
         Left            =   1440
         List            =   "FrmCurs.frx":0381
         TabIndex        =   0
         Text            =   "9999-9999"
         Top             =   360
         Width           =   1215
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
         Index           =   1
         Left            =   2760
         TabIndex        =   10
         Top             =   360
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
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Records Found at"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "FrmCurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   Curriculum Module
'Description    :   Use for navigation, adding, deleting, editing of subjects for
'                   a given curriculum. You can only display 1 curriculum year at
'                   a time at a specific Semester and Course.
'Date           :   January 25, 2005
'Programmer     :   iehjsuckers
'comment        :   ?
'********************************************************************************
Public SY As String, SEM As String, Course As String, yr As String, SC As String

Public isadd As Boolean
Private Sub BtnAdd_Click()
    SY = txtSY.Text
    SEM = txtSem.Text
    Course = txtCourse.Text
    yr = txtYr.Text
    isadd = True
    Load FrmAddCurSub
    FrmAddCurSub.Show 1, MainForm
    BtnShow_Click
End Sub

Private Sub BtnDelete_Click()
If LvCurSub.SelectedItem Is Nothing Then Exit Sub
'Delete here
    FrmConfirm.Show 1
    If IsTrue = False Then
        MsgBox "You do not have the permission to delete files.", vbCritical, "Error"
        Exit Sub
    End If

If MsgBox("do you really want to delete this record?", vbYesNo, "Confirm") = vbYes Then
    Dim msg As String
    msg = "Delete from Curriculum where SY = '" & txtSY.Text & "' and sem = '" & txtSem.Text & "' and course='" & txtCourse.Text & "' and yr = '" & txtYr.Text & "' and sc = '" & LvCurSub.SelectedItem.SubItems(1) & "'"
    SetSelectedCur msg
    BtnShow_Click
End If
End Sub

Private Sub BtnPrint_Click()
' SHAPE {Select SY,SEM,COURSE,YR FROM CURRICULUM group by sy,sem,course,YR}  AS CurrRep APPEND ({Select * from curriculum}  AS BodyCur RELATE 'SY' TO 'SY','SEM' TO 'Sem','COURSE' TO 'Course') AS BodyCur
Dim myx As String
myx = " SHAPE {Select SY,SEM,COURSE,YR FROM CURRICULUM where course ='" & txtCourse.Text & "' and sy = '" & txtSY.Text & "' group by sy,sem,course,YR}  AS CurrRep APPEND ({Select * from curriculum where course ='" & txtCourse.Text & "' and sy = '" & txtSY.Text & "'}  AS BodyCur RELATE 'SY' TO 'SY','SEM' TO 'Sem','COURSE' TO 'Course') AS BodyCur"
With DE.rsCurrRep
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .Open myx, DE.Con, adOpenDynamic, adLockOptimistic
End With
With DREPCur
    .Sections("PageHeader").Controls("LblSY").Caption = txtSY.Text
    .Sections("PageHeader").Controls("LBLCourse").Caption = txtCourse.Text
    .Show 1
End With
End Sub

Private Sub BtnRetrive_Click()
GroupSYSem
End Sub

Private Sub BtnShow_Click()
Dim msg As String
msg = "Select * from Curriculum where SY='" & txtSY.Text & "' and SEM = '" & txtSem.Text & _
    "' and Course = '" & txtCourse.Text & "'"
    If ChkAll.Value <> vbChecked Then
        msg = msg & " and yr = '" & txtYr.Text & "'"
    End If
SetSelectedCur msg
LoadCurtoLV Me.LvCurSub
GroupSYSem
End Sub

Private Sub Form_Load()
GetCourse
End Sub

Private Sub LstSY_Click()
LstSY.ToolTipText = LstSY.Text
End Sub

Private Sub LvCurSub_DblClick()
If LvCurSub.SelectedItem Is Nothing Then Exit Sub
isadd = False
'show the form
    SY = txtSY.Text
    SEM = txtSem.Text
    Course = txtCourse.Text
    yr = txtYr.Text
    SC = LvCurSub.SelectedItem.Text
    isadd = False
    Load FrmAddCurSub
    FrmAddCurSub.rebinds
    FrmAddCurSub.Show 1
BtnShow_Click
End Sub

'User functions
Private Sub GetCourse()
    With DE.rsCourses
        If .State <> 0 Then .Close
        .Open "Select * from courses"
        'clear the combobox
        txtCourse.Clear
        If .RecordCount > 0 Then
            Do Until .EOF
                txtCourse.AddItem .Fields(0).Value
                .MoveNext
            Loop
        End If
    End With
End Sub

Public Function GroupSYSem()
    On Error GoTo bx
    Dim rs As New Recordset
    'clear data
    LstSY.Clear
    With rs
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open "Select SY,SEM,course,yr from Curriculum Group by SY,SEM,course,yr ", DE.Con
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

'end of user functions
