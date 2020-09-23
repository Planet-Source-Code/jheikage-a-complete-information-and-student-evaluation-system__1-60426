VERSION 5.00
Begin VB.Form FrmColRecAdd 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Name Here"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5145
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmColRecAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton BtnOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Okay"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Subject Infromation"
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton BtnSel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "..."
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox TxtRemarks 
         Appearance      =   0  'Flat
         DataField       =   "Remarks"
         Height          =   360
         ItemData        =   "FrmColRecAdd.frx":030A
         Left            =   3720
         List            =   "FrmColRecAdd.frx":031A
         TabIndex        =   4
         Text            =   "NONE"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtGrade 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2400
         TabIndex        =   3
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtUnits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "Units"
         Height          =   360
         Left            =   1305
         TabIndex        =   2
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         DataField       =   "Description"
         Height          =   360
         Left            =   1305
         TabIndex        =   1
         Top             =   855
         Width           =   3375
      End
      Begin VB.TextBox txtSC 
         Appearance      =   0  'Flat
         DataField       =   "SC"
         Height          =   360
         Left            =   1305
         TabIndex        =   0
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks:"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   2760
         TabIndex        =   26
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Grade:"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   1770
         TabIndex        =   25
         Top             =   1305
         Width           =   585
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Units:"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   765
         TabIndex        =   24
         Top             =   1290
         Width           =   510
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   240
         TabIndex        =   23
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Code:"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   45
         TabIndex        =   22
         Top             =   525
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "School Year Information"
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.ComboBox TxtSem 
         DataSource      =   "DE"
         Height          =   360
         ItemData        =   "FrmColRecAdd.frx":033C
         Left            =   3720
         List            =   "FrmColRecAdd.frx":0349
         TabIndex        =   20
         Text            =   "Sem"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox TxtSy 
         DataSource      =   "DE"
         Height          =   360
         ItemData        =   "FrmColRecAdd.frx":035C
         Left            =   1440
         List            =   "FrmColRecAdd.frx":038D
         TabIndex        =   19
         Text            =   "9999-9999"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Semester:"
         Height          =   240
         Index           =   5
         Left            =   2760
         TabIndex        =   18
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "School Year:"
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.TextBox txtmnam 
      DataField       =   "mnam"
      Height          =   360
      Left            =   2400
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtfnam 
      DataField       =   "fnam"
      Height          =   360
      Left            =   2400
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtlnam 
      DataField       =   "lnam"
      Height          =   360
      Left            =   2400
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtIdnum 
      DataField       =   "Idnum"
      Height          =   360
      Left            =   1200
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame3 
      Caption         =   "Course Information"
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   1080
      Visible         =   0   'False
      Width           =   4935
      Begin VB.ComboBox TxtCourse 
         DataMember      =   "SubjectsEnrolled"
         DataSource      =   "DE"
         Height          =   360
         Left            =   120
         TabIndex        =   29
         Text            =   "Course"
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox TxtYr 
         Height          =   360
         ItemData        =   "FrmColRecAdd.frx":0436
         Left            =   2640
         List            =   "FrmColRecAdd.frx":0449
         TabIndex        =   28
         Text            =   "1"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Middle Name:"
      Height          =   240
      Index           =   3
      Left            =   2400
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   240
      Index           =   2
      Left            =   2400
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Last Name:"
      Height          =   240
      Index           =   1
      Left            =   2400
      TabIndex        =   10
      Top             =   375
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ID Number:"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FrmColRecAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   College Record Adding Addon
'Description    :   This is the form for adding new college records(Subjects for a
'                   given student) for a specified school year and semester.
'                   this will only activate if you click add or edit from the main
'                   form of College Record Adding Module.
'Date           :   February 5, 2005
'Programmer     :   iehjsuckers
'Comment        :   ?
'*********************************************************************************
Private Sub BtnCancel_Click()
If FrmStudentSubjectList.isadd = True Then DE.rsSubjectsEnrolled.CancelBatch adAffectAllChapters
Unload Me
End Sub

Private Sub BtnOk_Click()
'updated on march 18, 2005
On Error GoTo Erb
'check if it is in the subjects_offered.
If CheckExistance = False Then Exit Sub
If CheckPrerequisites = False Then Exit Sub
setvaluesColRec
DE.rsSubjectsEnrolled.Update
MsgBox "Record Saved.", vbInformation, "Message"
Unload Me
Exit Sub
Erb:
MsgBox Err.Description, vbCritical, "Error"
DE.rsCurricula.Cancel
End Sub

Private Sub BtnSel_Click()
    FrmSubList.who_Called = 2
    FrmSubList.Show 1
End Sub

Private Sub Form_Load()
Select Case FrmStudentSubjectList.isadd
Case True
    'addnew
    'setthe values here
    If DE.rsSubjectsEnrolled.State = 0 Then DE.rsSubjectsEnrolled.Open
    DE.rsSubjectsEnrolled.AddNew
    With FrmStudentSubjectList
        txtfnam.Text = .Fnam
        txtIdnum.Text = .Idnum
        txtlnam.Text = .Lnam
        txtmnam.Text = .Mnam
        TxtSy.Text = .TxtSy.Text
        TxtSem.Text = .TxtSem.Text
        Caption = .Lnam & ", " & .Fnam & " " & .Mnam
        TxtCourse.Text = .Course
        TxtYr.Text = .yr
    End With
Case False
    'update get values
    getValuesColRec
End Select
End Sub

Function getValuesColRec()
txtIdnum.Text = DE.rsSubjectsEnrolled.Fields("idnum").Value
txtlnam.Text = DE.rsSubjectsEnrolled.Fields("Lnam").Value
txtmnam.Text = DE.rsSubjectsEnrolled.Fields("mnam").Value
txtfnam.Text = DE.rsSubjectsEnrolled.Fields("fnam").Value
TxtSy.Text = DE.rsSubjectsEnrolled.Fields("sy").Value
TxtSem.Text = DE.rsSubjectsEnrolled.Fields("sem").Value
TxtCourse.Text = DE.rsSubjectsEnrolled.Fields("course").Value
TxtYr.Text = DE.rsSubjectsEnrolled.Fields("yr").Value
txtSC.Text = DE.rsSubjectsEnrolled.Fields("sc").Value
txtDescription.Text = DE.rsSubjectsEnrolled.Fields("description").Value
txtUnits.Text = DE.rsSubjectsEnrolled.Fields("units").Value
txtGrade.Text = DE.rsSubjectsEnrolled.Fields("grade").Value
TxtRemarks.Text = DE.rsSubjectsEnrolled.Fields("remarks").Value
End Function

Function setvaluesColRec()
DE.rsSubjectsEnrolled.Fields("IDNUM").Value = txtIdnum.Text
DE.rsSubjectsEnrolled.Fields("Lnam").Value = txtlnam.Text
DE.rsSubjectsEnrolled.Fields("mnam").Value = txtmnam.Text
DE.rsSubjectsEnrolled.Fields("fnam").Value = txtfnam.Text
DE.rsSubjectsEnrolled.Fields("sy").Value = TxtSy.Text
DE.rsSubjectsEnrolled.Fields("sem").Value = TxtSem.Text
DE.rsSubjectsEnrolled.Fields("course").Value = TxtCourse.Text
DE.rsSubjectsEnrolled.Fields("yr").Value = TxtYr.Text
DE.rsSubjectsEnrolled.Fields("sc").Value = txtSC.Text
DE.rsSubjectsEnrolled.Fields("description").Value = txtDescription.Text
If IsNumeric(txtUnits.Text) Then
DE.rsSubjectsEnrolled.Fields("units").Value = txtUnits.Text
Else
    MsgBox "Invalid Input.", vbCritical, "Error"
    txtUnits.SetFocus
    SendKeys "{HOME}+{END}"
End If
If IsNumeric(txtGrade.Text) Then
DE.rsSubjectsEnrolled.Fields("grade").Value = txtGrade.Text
Else
DE.rsSubjectsEnrolled.Fields("grade").Value = 0
End If
DE.rsSubjectsEnrolled.Fields("remarks").Value = TxtRemarks.Text
End Function

Private Sub txtUnits_LostFocus()
    If IsNumeric(txtUnits.Text) = False Then
        MsgBox "Invalid Input.", vbCritical, "Error"
        txtUnits.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

'user functions
Private Function CheckExistance() As Boolean    'check existance of the subject being added
Dim msg As String
Dim rs As New Recordset
On Error GoTo Erb
msg = "Select * from Subjects_Offered where SY='" & TxtSy.Text & "' and SEM = '" & TxtSem.Text & _
    "' and Course = '" & TxtCourse.Text & "'"
With rs
    .CursorLocation = adUseClient
    .Open msg, DE.Con, adOpenDynamic, adLockOptimistic
    If .RecordCount = 0 Then
        MsgBox "This subject does not appear to be offered with the specified school year and semester. Please contact your administrator.", vbInformation, "Subject Not Offered"
        CheckExistance = False
    Else
        CheckExistance = True
    End If
    .Close
End With
Set rs = Nothing
Exit Function
Erb:
MsgBox Err.Description, vbCritical, "Error"
Set rs = Nothing
CheckExistance = False
End Function

Private Function CheckPrerequisites() As Boolean    'check for prerequisites of a subject before you add
Dim msg As String, prere As String, Splt
Dim rs As New Recordset
'On Error GoTo erb
msg = "Select * from Subjects_Offered where SY='" & TxtSy.Text & "' and SEM = '" & TxtSem.Text & _
    "' and Course = '" & TxtCourse.Text & "'"
With rs
    .CursorLocation = adUseClient
    .Open msg, DE.Con, adOpenDynamic, adLockOptimistic
    If .RecordCount = 0 Then
        MsgBox "This subject does not appear to be offered with the specified school year and semester. Please contact your administrator.", vbInformation, "Subject Not Offered"
    Else
        'check for prerequisites here
        Splt = Split(.Fields("prerequisites").Value, ", ", , vbTextCompare)
        Dim mymess As String, i As Integer
        Dim rs2 As New Recordset
        For i = LBound(Splt) To UBound(Splt)
            rs2.CursorLocation = adUseClient
            mymess = "Select * From subjectsEnrolled where lnam='" & txtlnam.Text & "' and mnam = '" & txtmnam.Text & "' and fnam = '" & txtfnam.Text & "' and sc = '" & Splt(i) & "'"
            rs2.Open mymess, DE.Con, adOpenDynamic, adLockOptimistic
            If rs2.RecordCount = 0 Then
                MsgBox "This student does not appear to take some of the prerequisite of the subject being added. Please evluate first the student.", vbInformation, "Prerequisite not taken"
                CheckPrerequisites = False
                Exit Function
            Else
                'if there is a record check if he failed or not
                If UCase(rs2.Fields("Remarks").Value) <> "PASSED" Then
                    MsgBox "This student appears to have not passed some of the prerequisites on this subject. Please evaluate first the student.", vbInformation, "Prerequisite not Passed"
                    CheckPrerequisites = False
                    Exit Function
                End If
            End If
            rs2.Close
        Next
    End If
End With
'no problem, can add
CheckPrerequisites = True
Exit Function
Erb:
MsgBox Err.Description, vbCritical, "Error"
Set rs = Nothing
CheckPrerequisites = False
End Function
'end of user functions
