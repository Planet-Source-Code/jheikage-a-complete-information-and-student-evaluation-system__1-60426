VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSubList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Subject List"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSubList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnLoad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Load Subjects"
      Height          =   735
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin MSComctlLib.ListView LVLists 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2355
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "DESCRIPTION"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "UNITS"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton BtnOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   975
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
      Width           =   4935
      Begin VB.ComboBox TxtSy 
         Appearance      =   0  'Flat
         DataSource      =   "DE"
         Height          =   360
         ItemData        =   "FrmSubList.frx":030A
         Left            =   1440
         List            =   "FrmSubList.frx":033B
         TabIndex        =   0
         Text            =   "9999-9999"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox TxtSem 
         Appearance      =   0  'Flat
         DataSource      =   "DE"
         Height          =   360
         ItemData        =   "FrmSubList.frx":03E4
         Left            =   3720
         List            =   "FrmSubList.frx":03F1
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Course Info"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   3735
      Begin VB.CheckBox ChkAll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Show all Year Level"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox TxtYr 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmSubList.frx":0404
         Left            =   1560
         List            =   "FrmSubList.frx":0417
         TabIndex        =   3
         Text            =   "1"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox TxtCourse 
         Appearance      =   0  'Flat
         DataMember      =   "SubjectsEnrolled"
         DataSource      =   "DE"
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Text            =   "Course"
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmSubList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   Subject List Module (College Record Add on)
'Description    :   This is a module to make adding of records for subjects easier.
'                   This will only activate if you click the elipse button (...) in
'                   College Record Adding Addon Module. You just have to select the
'                   school year, semester, and course and year that you want to select.
'                   NEW! just added functionality to comply with the subjects offering!
'                   if the integer Who_Called = 1 , the one calling this function is
'                   the Form for College record adding.
'Date           :   February 27, 2005
'Programmer     :   iehjsuckers
'Comment        :   Better idea? send me
'************************************************************************************
Public SC As String, Description As String, Units As Double
Public who_Called As Integer    '1 for subject offering and 2 for college record
Private Sub BtnLoad_Click()
Dim msg As String
Select Case who_Called
Case 1
    msg = "Select * from curriculum where SY='" & txtSY.Text & "' and SEM = '" & txtSem.Text & _
        "' and Course = '" & txtCourse.Text & "'"
        If ChkAll.Value <> vbChecked Then
            msg = msg & " and yr = '" & txtYr.Text & "'"
        End If
    SetSelectedCur msg
    loadSLToListLv Me.LVLists
Case 2
    msg = "Select * from Subjects_Offered where SY='" & txtSY.Text & "' and SEM = '" & txtSem.Text & _
        "' and Course = '" & txtCourse.Text & "'"
        If ChkAll.Value <> vbChecked Then
            msg = msg & " and yr = '" & txtYr.Text & "'"
        End If
    SetSelectedSubOff msg
    loadSOToListLv Me.LVLists
End Select


End Sub

Private Sub BtnOk_Click()
    If LVLists.SelectedItem Is Nothing Then
        MsgBox "No Subject Selected.", vbInformation, "Error"
    Else
        'set the values\
        Select Case who_Called
        Case 2
            With FrmColRecAdd
                .txtSC.Text = LVLists.SelectedItem.Text
                .txtDescription.Text = LVLists.SelectedItem.SubItems(1)
                .txtUnits.Text = LVLists.SelectedItem.SubItems(2)
                Unload Me
            End With
        Case 1
            With FrmAddSubOff
                .txtSC.Text = LVLists.SelectedItem.Text
                .txtDescription.Text = LVLists.SelectedItem.SubItems(1)
                .txtunts.Text = LVLists.SelectedItem.SubItems(2)
                Unload Me
            End With
        End Select
    End If
End Sub

Private Sub Form_Load()
GetCourse
End Sub

'User Functions
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


