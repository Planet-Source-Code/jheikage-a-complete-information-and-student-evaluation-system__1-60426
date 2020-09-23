VERSION 5.00
Begin VB.Form FrmAddSubOff 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Subject Offered"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5955
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
   Icon            =   "FrmAddSubOff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSC 
      Appearance      =   0  'Flat
      DataField       =   "SC"
      Height          =   360
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hidden"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtSY 
         DataSource      =   "DE"
         Height          =   360
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   1650
      End
      Begin VB.TextBox txtSem 
         DataSource      =   "DE"
         Height          =   360
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtCourse 
         DataSource      =   "DE"
         Height          =   360
         Left            =   4080
         TabIndex        =   13
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txtYr 
         DataSource      =   "DE"
         Height          =   360
         Left            =   5040
         TabIndex        =   12
         Top             =   240
         Width           =   180
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SY:"
         Height          =   240
         Index           =   0
         Left            =   465
         TabIndex        =   19
         Top             =   285
         Width           =   330
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sem:"
         Height          =   240
         Index           =   1
         Left            =   2520
         TabIndex        =   18
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Course:"
         Height          =   240
         Index           =   2
         Left            =   3720
         TabIndex        =   17
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Yr:"
         Height          =   240
         Index           =   3
         Left            =   4800
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CommandButton btnelipse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Select from available Curriculum"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      DataField       =   "Description"
      Height          =   360
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   4335
   End
   Begin VB.TextBox txtunts 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "unts"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtPrerequisites 
      Appearance      =   0  'Flat
      DataField       =   "Prerequisites"
      Height          =   360
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   4380
   End
   Begin VB.CommandButton BtnOk 
      BackColor       =   &H80000004&
      Caption         =   "&OK"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      BackColor       =   &H80000004&
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Code:"
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   165
      Width           =   1230
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "units:"
      Height          =   240
      Index           =   6
      Left            =   915
      TabIndex        =   7
      Top             =   1170
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prerequisites:"
      Height          =   240
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1200
   End
End
Attribute VB_Name = "FrmAddSubOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   Subjects Offered Adding Add On
'Description    :   Adds subjects offered for a specified school year and semester.
'                   This is the adding and editing form. This will only activate
'                   after you click add in the main form of Subjects Offered Module.
'
'Date           :   February 1, 2005
'Programmer     :   iehjsuckers
'Last Update    :   February 20, 2005
'Comments       :   ?
'********************************************************************************
Private Sub BtnCancel_Click()
If FrmSubOffers.isadd = True Then DE.rsSubjectOffered.CancelBatch adAffectAllChapters
SetSelectedSubOff "Select * From subjects_Offered where sy = '" & FrmCurs.SY & "' and sem = '" & FrmCurs.SEM & "' and Course= '" & FrmCurs.Course & "' and yr = '" & FrmCurs.yr & "'"
Unload Me
End Sub

Private Sub btnelipse_Click()
FrmSubList.who_Called = 1
FrmSubList.Show 1
End Sub

Private Sub BtnOk_Click()
On Error GoTo Erb
setValues
DE.rsSubjectOffered.Update
MsgBox "Record Saved.", vbInformation, "Message"
Unload Me
Exit Sub
Erb:
MsgBox Err.Description, vbCritical, "Error"
DE.rsSubjectOffered.Cancel
'Refresh
'With FrmCurs
'    SetSelectedCur "Select * from Curriculum where sy = '" & .SY & "' sem = '" & .SEM & "' and course = '" & .Course & "' and yr = '" & .Yr & "'"
'End With
End Sub

Private Sub Form_Load()
Select Case FrmSubOffers.isadd
Case True
    'add new
    If DE.rsSubjectOffered.State = 0 Then DE.rsSubjectOffered.Open
    DE.rsSubjectOffered.AddNew
    'set values here
    With FrmSubOffers
        txtSem.Text = .SEM
        txtSY.Text = .SY
        txtYr.Text = .yr
        txtCourse.Text = .Course
    End With

Case False
    'edit lang po
    SetSelectedSubOff "Select * From subjects_Offered where sy = '" & FrmSubOffers.SY & "' and sem = '" & FrmSubOffers.SEM & "' and Course= '" & FrmSubOffers.Course & "' and yr = '" & FrmSubOffers.yr & "' and sc = '" & FrmSubOffers.LvCurSub.SelectedItem.SubItems(1) & "'"
    getValues
End Select
End Sub

'User Functions
Function getValues()
On Error GoTo Erb
    txtDescription.Text = DE.rsSubjectOffered("Description").Value
    txtPrerequisites.Text = DE.rsSubjectOffered("Prerequisites").Value
    txtunts.Text = DE.rsSubjectOffered("Unts").Value
    txtSC.Text = DE.rsSubjectOffered("SC").Value
Exit Function
Erb:
MsgBox Err.Description, vbCritical, "Error"
Exit Function
End Function

Function setValues()
On Error GoTo Erb
    DE.rsSubjectOffered("Description").Value = Trim(txtDescription.Text)
    DE.rsSubjectOffered("Prerequisites").Value = Trim(txtPrerequisites.Text)
    DE.rsSubjectOffered("Unts").Value = Trim(txtunts.Text)
    DE.rsSubjectOffered("SC").Value = Trim(txtSC.Text)
    DE.rsSubjectOffered("Sy").Value = Trim(txtSY.Text)
    DE.rsSubjectOffered("Sem").Value = Trim(txtSem.Text)
    DE.rsSubjectOffered("Course").Value = Trim(txtCourse.Text)
    DE.rsSubjectOffered("Yr").Value = Trim(txtYr.Text)
Exit Function
Erb:
MsgBox Err.Description, vbCritical, "Error"
Exit Function
End Function
Function rebinds()
    Set txtSem.DataSource = DE
    txtSem.DataMember = "subjectoffered"
    txtSem.DataField = "Sem"
    Set txtunts.DataSource = DE
    txtunts.DataMember = "subjectoffered"
    txtunts.DataField = "unts"
    Set txtCourse.DataSource = DE
    txtCourse.DataMember = "subjectoffered"
    txtCourse.DataField = "Course"
    Set txtYr.DataSource = DE
    txtYr.DataMember = "subjectoffered"
    txtYr.DataField = "yr"
    Set txtSC.DataSource = DE
    txtSC.DataMember = "subjectoffered"
    txtSC.DataField = "SC"
    Set txtSY.DataSource = DE
    txtSY.DataMember = "subjectoffered"
    txtSY.DataField = "Sy"
    Set txtPrerequisites.DataSource = DE
    txtPrerequisites.DataMember = "subjectoffered"
    txtPrerequisites.DataField = "Prerequisites"
    Set txtDescription.DataSource = DE
    txtDescription.DataMember = "subjectoffered"
    txtDescription.DataField = "Description"
End Function
'End of User Functions


Private Sub txtunts_LostFocus()
    If IsNumeric(txtunts.Text) = False Then
        MsgBox "Invalid input.", vbCritical, "Error"
        txtunts.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

