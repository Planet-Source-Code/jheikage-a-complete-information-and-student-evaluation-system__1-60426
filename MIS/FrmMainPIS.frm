VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMainPIS 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PIS"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMainPIS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imx 
      Left            =   360
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMainPIS.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMainPIS.frx":075C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton BtnEvaluate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "E&valuate Student"
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Evaluate the Active Student"
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton BtnRecords 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&College Records"
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "View the subjects taken and enrolled by the selected student."
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton BtnPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Print"
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Print the record of the selected student"
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton BtnDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Delete"
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Delete Information about the selected student"
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton BtnEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Edit"
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Edit the Profile of the selected student"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton BtnNew 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&New"
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Enter a new student Personal Information"
      Top             =   2400
      Width           =   1935
   End
   Begin MSComctlLib.ListView LVStudents 
      Height          =   4215
      Left            =   120
      TabIndex        =   41
      Top             =   2400
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imx"
      SmallIcons      =   "imx"
      ColHdrIcons     =   "imx"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IDNO"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Last Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "First Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Middle Name"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Display Box"
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.CheckBox chklike 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Like"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6840
         TabIndex        =   45
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton BtnAdvQ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "..."
         Height          =   375
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton BtnGo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "GO"
         Height          =   375
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TxtQuery 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7680
         TabIndex        =   29
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox CbxField 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmMainPIS.frx":2390
         Left            =   5640
         List            =   "FrmMainPIS.frx":23A3
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Ñ"
         Height          =   375
         Index           =   26
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Z"
         Height          =   375
         Index           =   25
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Y"
         Height          =   375
         Index           =   24
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "X"
         Height          =   375
         Index           =   23
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "W"
         Height          =   375
         Index           =   22
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "V"
         Height          =   375
         Index           =   21
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "U"
         Height          =   375
         Index           =   20
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "T"
         Height          =   375
         Index           =   19
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "S"
         Height          =   375
         Index           =   18
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "R"
         Height          =   375
         Index           =   17
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Q"
         Height          =   375
         Index           =   16
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "P"
         Height          =   375
         Index           =   15
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "O"
         Height          =   375
         Index           =   14
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "N"
         Height          =   375
         Index           =   13
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "M"
         Height          =   375
         Index           =   12
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "L"
         Height          =   375
         Index           =   11
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "K"
         Height          =   375
         Index           =   10
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "J"
         Height          =   375
         Index           =   9
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "I"
         Height          =   375
         Index           =   8
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "H"
         Height          =   375
         Index           =   7
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "G"
         Height          =   375
         Index           =   6
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "F"
         Height          =   375
         Index           =   5
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "E"
         Height          =   375
         Index           =   4
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "D"
         Height          =   375
         Index           =   3
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "C"
         Height          =   375
         Index           =   2
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "B"
         Height          =   375
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Btns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "A"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Advance Query: If you know sql syntax, you can click here."
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4680
         TabIndex        =   44
         Top             =   1440
         Width           =   5145
      End
      Begin VB.Label LblDB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Count of PIS Records in your database."
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4680
         TabIndex        =   43
         Top             =   1080
         Width           =   3435
      End
      Begin VB.Label LblQinfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "About Your query."
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   42
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Column:"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4800
         TabIndex        =   40
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Or you can type the query here hit go."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         TabIndex        =   39
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select the starting letter of the last name you want to display."
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   4380
      End
   End
   Begin VB.Menu RMenu 
      Caption         =   "Hidden"
      Visible         =   0   'False
      Begin VB.Menu HNew 
         Caption         =   "&New Student"
      End
      Begin VB.Menu HEdit 
         Caption         =   "&Edit Profile"
      End
      Begin VB.Menu HDelete 
         Caption         =   "&Delete Profile"
      End
      Begin VB.Menu HPrint 
         Caption         =   "&Print Profile"
      End
      Begin VB.Menu HView 
         Caption         =   "&View College Records"
      End
      Begin VB.Menu HEval 
         Caption         =   "E&valuate Student"
      End
   End
End
Attribute VB_Name = "FrmMainPIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   SPI, EVALUATION, COLLEGE RECORD Main Module
'Description    :   This is the heart of the entire system. All Student records
'                   are seen here. You can access the evaluation, SPI, and college
'                   records. Adding and udpating of student personal information is
'                   done here.
'Date           :   January 4, 2005
'Programmer     :   iehjsuckers
'Last Update    :   January 30, 2005
'Comments       :   For further development
'**********************************************************************************
Public isadd As Boolean
Private lastclick As Integer
Private Sub BtnAdvQ_Click()
    MsgBox "Not yet implemented.", vbInformation, "Error"
End Sub

Private Sub BtnDelete_Click()
    If LVStudents.SelectedItem Is Nothing Then Exit Sub
    FrmConfirm.Show 1
    If IsTrue = False Then
        MsgBox "You do not have the permission to delete files.", vbCritical, "Error"
        Exit Sub
    End If
    Dim x As Integer
    x = MsgBox("Do you want to Delete " & LVStudents.SelectedItem.SubItems(1) & ", " & LVStudents.SelectedItem.SubItems(2) & " " & LVStudents.SelectedItem.SubItems(3) & " and all its College records stored in your database?", vbYesNo + vbQuestion, "Confirmation")
    If x = vbYes Then
        SetRs "Delete From names where lnam = '" & LVStudents.SelectedItem.SubItems(1) & "' and fnam = '" & LVStudents.SelectedItem.SubItems(2) & "' and mnam = '" & LVStudents.SelectedItem.SubItems(3) & "'" 'this will be added idnum = '" & LVStudents.SelectedItem.Text & "' and
        'delete also college records
        SetSelectedStudSub "Delete From subjectsenrolled where lnam = '" & LVStudents.SelectedItem.SubItems(1) & "' and fnam = '" & LVStudents.SelectedItem.SubItems(2) & "' and mnam = '" & LVStudents.SelectedItem.SubItems(3) & "'"
    End If
    If lastclick = -20 Then BtnGo_Click Else Btns_Click lastclick
End Sub

Private Sub BtnEdit_Click()
    isadd = False
    SetRs "Select * From names where lnam = '" & FrmMainPIS.LVStudents.SelectedItem.SubItems(1) & _
            "' and fnam = '" & FrmMainPIS.LVStudents.SelectedItem.SubItems(2) & "' and mnam = '" & _
             FrmMainPIS.LVStudents.SelectedItem.SubItems(3) & "'" 'this will be added lst"' and idnum = '" & FrmMainPIS.LVStudents.SelectedItem.Text & "'"
    'Load FrmPIS
    FrmPIS.Show 1
    If lastclick = -20 Then BtnGo_Click Else Btns_Click lastclick
    DE.rsPIS.Close
End Sub

Private Sub BtnEvaluate_Click()
    If LVStudents.SelectedItem Is Nothing Then Exit Sub
    
    With FrmEvaluate
        If LVStudents.SelectedItem.ToolTipText = "" Then
            MsgBox "This student do not contain any record on what curriculum to follow. Please select from the list in the evaluation Form.", vbInformation, "Information"
        Else
            .txtSY.Text = LVStudents.SelectedItem.ToolTipText
        End If
        If LVStudents.SelectedItem.Tag = "" Then
            MsgBox "This student do not contain any record about his/her course. Please select from the list in the evaluation Form.", vbInformation, "Information"
        Else
            .txtCourse.Text = LVStudents.SelectedItem.Tag
        End If
        .LblName.Caption = "|" & Trim(LVStudents.SelectedItem.Text) & "|" & Trim(LVStudents.SelectedItem.SubItems(1)) & "|" & Trim(LVStudents.SelectedItem.SubItems(2)) & "|" & Trim(LVStudents.SelectedItem.SubItems(3)) & "|"
        .Show 1
    End With
End Sub

Private Sub BtnGo_Click()
    If CbxField.Text = "" Then Exit Sub 'ensure selection
    Dim msg As String
    msg = "Select * From names where " & CbxField.Text
    If chklike.Value = vbChecked Then   'use like
        msg = msg & " like '" & Me.TxtQuery.Text & "'"
    Else
        msg = msg & " = '" & Me.TxtQuery.Text & "'"
    End If
    LoadLV Me.LVStudents, msg
    
    lastclick = -20
End Sub

Private Sub BtnNew_Click()
    isadd = True
    Load FrmPIS
    FrmPIS.Show 1
    If lastclick = -20 Then BtnGo_Click Else Btns_Click lastclick
End Sub

Private Sub BtnPrint_Click()
If LVStudents.SelectedItem Is Nothing Then Exit Sub
    With DE.rsPIS
        'select the selected one
        If .State <> 0 Then .Close
        .Open "Select * from names where lnam = '" & LVStudents.SelectedItem.SubItems(1) & "' and fnam = '" & LVStudents.SelectedItem.SubItems(2) & "' and mnam = '" & LVStudents.SelectedItem.SubItems(3) & "' and idnum = '" & LVStudents.SelectedItem.Text & "'"
    End With
    'show the report here
    DREPPIS.Show 1
    DE.rsPIS.Close
End Sub

Private Sub BtnRecords_Click()
    If LVStudents.SelectedItem Is Nothing Then Exit Sub
    With FrmStudentSubjectList
        .Caption = "|" & Trim(LVStudents.SelectedItem.Text) & "|" & Trim(LVStudents.SelectedItem.SubItems(1)) & "|" & Trim(LVStudents.SelectedItem.SubItems(2)) & "|" & Trim(LVStudents.SelectedItem.SubItems(3)) & "|"
        .Show 1
    End With
End Sub

Private Sub Btns_Click(Index As Integer)
    LoadLV Me.LVStudents, "Select * From names where lnam like '" & Btns(Index).Caption & "%'"
    lastclick = Index
End Sub

Private Sub Form_Load()
    Btns_Click 0
    SetRs "Select * from names"
    'set captions
    Me.LblDB.Caption = "Total Record Count: " & DE.rsPIS.RecordCount
    Me.LblQinfo.Caption = "Total Query Result: " & LVStudents.ListItems.Count
    DE.rsPIS.Close
End Sub

Private Sub HDelete_Click()
BtnDelete_Click
End Sub

Private Sub HEdit_Click()
    BtnEdit_Click
End Sub

Private Sub HEval_Click()
    BtnEvaluate_Click
End Sub

Private Sub HNew_Click()
    BtnNew_Click
End Sub

Private Sub HPrint_Click()
    BtnPrint_Click
End Sub

Private Sub HView_Click()
    BtnRecords_Click
End Sub

Private Sub LVStudents_Click()
    'this code is udpdated on march 17,2005
    If LVStudents.SelectedItem Is Nothing Then Exit Sub
    Dim mycur As String, mycorse As String, msg As String
    mycur = LVStudents.SelectedItem.ToolTipText
    mycorse = LVStudents.SelectedItem.Tag
    If mycur = "" Then
        msg = "No curriculum to follow "
    Else
        msg = "Curriculum is " & mycur
    End If
    If mycorse = "" Then
        msg = msg & ", No course stored."
    Else
        msg = msg & ", Course is " & mycorse
    End If
    LVStudents.ToolTipText = msg
End Sub

Private Sub LVStudents_DblClick()
    If LVStudents.SelectedItem Is Nothing Then Exit Sub
    BtnEdit_Click
End Sub

Private Sub LVStudents_Mouseup(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        'show right click
        'select the file
        
        Me.PopupMenu RMenu
    End If
End Sub

Private Sub TxtQuery_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BtnGo_Click
End Sub
