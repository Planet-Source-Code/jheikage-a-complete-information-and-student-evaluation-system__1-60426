VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSubOffers 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subjects Offered"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSubOffers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "School Year Information"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   8055
      Begin VB.ComboBox TxtSy 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmSubOffers.frx":030A
         Left            =   1440
         List            =   "FrmSubOffers.frx":033B
         TabIndex        =   12
         Text            =   "9999-9999"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox TxtSem 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmSubOffers.frx":03E4
         Left            =   3720
         List            =   "FrmSubOffers.frx":03F1
         TabIndex        =   11
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
         Index           =   0
         Left            =   240
         TabIndex        =   14
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
         Index           =   1
         Left            =   2760
         TabIndex        =   13
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Course"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   6855
      Begin VB.ComboBox TxtYr 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmSubOffers.frx":0404
         Left            =   2760
         List            =   "FrmSubOffers.frx":0417
         TabIndex        =   9
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox TxtCourse 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   240
         TabIndex        =   8
         Text            =   "Course"
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox ChkAll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Show All Year Level"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton BtnShow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Load Records"
      Height          =   855
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton BtnRetrive 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Retrive"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
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
      Left            =   6240
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton BtnPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton BtnDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton BtnAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add Subject"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1455
   End
   Begin MSComctlLib.ListView LvCurSub 
      Height          =   3975
      Left            =   120
      TabIndex        =   16
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
Attribute VB_Name = "FrmSubOffers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NEW! You will notice that the table for curriculum and subjects_offered are the same
'The only difference is that a curriculum is used for evaluation, what the subject
'offering table does is that this is the subjects that are available for a certain school year
'The logic is, in a curriculum year, not all the subjects are offered. some are only offered
' and these subjects must appear and must be added to the subjects_offered table.
'Module Name    :   Subject Offered List Module (College Record Add on)
'Description    :   This module is for subjects offered. This will base on the curriculum being entered
'                   You will have to choose which subjects to offer from a certain school year and semester
'                   This is primarily used for enrollment purpose. The system determines what
'                   subjects are included and are not.
'Date           :   March 2, 2005
'Programmer     :   iehjsuckers
'Comment        :   Better idea? send me
'************************************************************************************
Public SY As String, SEM As String, Course As String, yr As String, SC As String

Public isadd As Boolean
Private Sub BtnAdd_Click()
    SY = txtSY.Text
    SEM = txtSem.Text
    Course = txtCourse.Text
    yr = txtYr.Text
    isadd = True
    Load FrmAddSubOff
    FrmAddSubOff.Show 1, MainForm
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
    msg = "Delete from subjects_Offered where SY = '" & txtSY.Text & "' and sem = '" & txtSem.Text & "' and course='" & txtCourse.Text & "' and yr = '" & txtYr.Text & "' and sc = '" & LvCurSub.SelectedItem.SubItems(1) & "'"
    SetSelectedCur msg
    BtnShow_Click
End If
End Sub

Private Sub BtnPrint_Click()
' SHAPE {SELECT SY,SEM,COURSE,YR from Subjects_Offered group by SY,SEM,COURSE,YR }  AS SubOffRep APPEND ({Select * From Subjects_Offered}  AS SubOffBody RELATE 'SY' TO 'SY','SEM' TO 'Sem','COURSE' TO 'Course','YR' TO 'Yr') AS SubOffBody
Dim myx As String
myx = "  SHAPE {SELECT SY,SEM,COURSE,YR from Subjects_Offered where sy = '" & txtSY.Text & "' and course = '" & txtCourse.Text & "' group by SY,SEM,COURSE,YR }  AS SubOffRep APPEND ({Select * From Subjects_Offered where sy = '" & txtSY.Text & "' and course = '" & txtCourse.Text & "'}  AS SubOffBody RELATE 'SY' TO 'SY','SEM' TO 'Sem','COURSE' TO 'Course','YR' TO 'Yr') AS SubOffBody"
With DE.rsSubOffRep
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .Open myx, DE.Con, adOpenDynamic, adLockOptimistic
End With
With DREPSE
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
msg = "Select * from subjects_Offered where SY='" & txtSY.Text & "' and SEM = '" & txtSem.Text & _
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
    Load FrmAddSubOff
    FrmAddSubOff.rebinds
    FrmAddSubOff.Show 1
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
        .Open "Select SY,SEM,course,yr from Subjects_Offered Group by SY,SEM,course,yr ", DE.Con
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


