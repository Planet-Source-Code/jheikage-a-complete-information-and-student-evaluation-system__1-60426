VERSION 5.00
Begin VB.Form FrmAddCurSub 
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Subject"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5850
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
   Icon            =   "FrmAddCurSub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnCancel 
      BackColor       =   &H80000004&
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton BtnOk 
      BackColor       =   &H80000004&
      Caption         =   "&OK"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hidden"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtYr 
         DataSource      =   "DE"
         Height          =   360
         Left            =   5040
         TabIndex        =   18
         Top             =   240
         Width           =   180
      End
      Begin VB.TextBox txtCourse 
         DataSource      =   "DE"
         Height          =   360
         Left            =   4080
         TabIndex        =   16
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txtSem 
         DataSource      =   "DE"
         Height          =   360
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtSY 
         DataSource      =   "DE"
         Height          =   360
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   1650
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Yr:"
         Height          =   240
         Index           =   3
         Left            =   4800
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Course:"
         Height          =   240
         Index           =   2
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sem:"
         Height          =   240
         Index           =   1
         Left            =   2520
         TabIndex        =   13
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SY:"
         Height          =   240
         Index           =   0
         Left            =   465
         TabIndex        =   11
         Top             =   285
         Width           =   330
      End
   End
   Begin VB.TextBox txtPrerequisites 
      Appearance      =   0  'Flat
      DataField       =   "Prerequisites"
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   4305
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
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      DataField       =   "Description"
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.TextBox txtSC 
      Appearance      =   0  'Flat
      DataField       =   "SC"
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prerequisites:"
      Height          =   240
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "units:"
      Height          =   240
      Index           =   6
      Left            =   915
      TabIndex        =   8
      Top             =   930
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Code:"
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   165
      Width           =   1230
   End
End
Attribute VB_Name = "FrmAddCurSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   Curriculum Subject Adding Add On
'Description    :   Adds subjects for a specified curriculum year and semester.
'                   This is the adding and editing form. This will only activate
'                   after you click add in the main form of Curriculum Module.
'
'Date           :   February 1, 2005
'Programmer     :   iehjsuckers
'Last Update    :   February 20, 2005
'Comments       :   ?
'********************************************************************************
Private Sub BtnCancel_Click()
If FrmCurs.isadd = True Then DE.rsCurricula.CancelBatch adAffectAllChapters
SetSelectedCur "Select * From Curriculum where sy = '" & FrmCurs.SY & "' and sem = '" & FrmCurs.SEM & "' and Course= '" & FrmCurs.Course & "' and yr = '" & FrmCurs.yr & "'"
Unload Me
End Sub

Private Sub BtnOk_Click()
On Error GoTo Erb
setValues
DE.rsCurricula.Update
MsgBox "Record Saved.", vbInformation, "Message"
Unload Me
Exit Sub
Erb:
MsgBox Err.Description, vbCritical, "Error"
DE.rsCurricula.Cancel
'Refresh
'With FrmCurs
'    SetSelectedCur "Select * from Curriculum where sy = '" & .SY & "' sem = '" & .SEM & "' and course = '" & .Course & "' and yr = '" & .Yr & "'"
'End With
End Sub

Private Sub Form_Load()
Select Case FrmCurs.isadd
Case True
    'add new
    If DE.rsCurricula.State = 0 Then DE.rsCurricula.Open
    DE.rsCurricula.AddNew
    'set values here
    With FrmCurs
        TxtSem.Text = .SEM
        TxtSy.Text = .SY
        TxtYr.Text = .yr
        TxtCourse.Text = .Course
    End With

Case False
    'edit lang po
    SetSelectedCur "Select * From Curriculum where sy = '" & FrmCurs.SY & "' and sem = '" & FrmCurs.SEM & "' and Course= '" & FrmCurs.Course & "' and yr = '" & FrmCurs.yr & "' and sc = '" & FrmCurs.LvCurSub.SelectedItem.SubItems(1) & "'"
    getValues
End Select
End Sub

'User Functions
Function getValues()
On Error GoTo Erb
    txtDescription.Text = DE.rsCurricula("Description").Value
    txtPrerequisites.Text = DE.rsCurricula("Prerequisites").Value
    txtunts.Text = DE.rsCurricula("Unts").Value
    txtSC.Text = DE.rsCurricula("SC").Value
Exit Function
Erb:
MsgBox Err.Description, vbCritical, "Error"
Exit Function
End Function

Function setValues()
On Error GoTo Erb
    DE.rsCurricula("Description").Value = Trim(txtDescription.Text)
    DE.rsCurricula("Prerequisites").Value = Trim(txtPrerequisites.Text)
    DE.rsCurricula("Unts").Value = Trim(txtunts.Text)
    DE.rsCurricula("SC").Value = Trim(txtSC.Text)
    DE.rsCurricula("Sy").Value = Trim(TxtSy.Text)
    DE.rsCurricula("Sem").Value = Trim(TxtSem.Text)
    DE.rsCurricula("Course").Value = Trim(TxtCourse.Text)
    DE.rsCurricula("Yr").Value = Trim(TxtYr.Text)
Exit Function
Erb:
MsgBox Err.Description, vbCritical, "Error"
Exit Function
End Function
Function rebinds()
    Set TxtSem.DataSource = DE
    TxtSem.DataMember = "Curricula"
    TxtSem.DataField = "Sem"
    Set txtunts.DataSource = DE
    txtunts.DataMember = "Curricula"
    txtunts.DataField = "unts"
    Set TxtCourse.DataSource = DE
    TxtCourse.DataMember = "Curricula"
    TxtCourse.DataField = "Course"
    Set TxtYr.DataSource = DE
    TxtYr.DataMember = "Curricula"
    TxtYr.DataField = "yr"
    Set txtSC.DataSource = DE
    txtSC.DataMember = "Curricula"
    txtSC.DataField = "SC"
    Set TxtSy.DataSource = DE
    TxtSy.DataMember = "Curricula"
    TxtSy.DataField = "Sy"
    Set txtPrerequisites.DataSource = DE
    txtPrerequisites.DataMember = "Curricula"
    txtPrerequisites.DataField = "Prerequisites"
    Set txtDescription.DataSource = DE
    txtDescription.DataMember = "Curricula"
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
