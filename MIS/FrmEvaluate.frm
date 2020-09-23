VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEvaluate 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evaluate Student"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   14430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmEvaluate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   14430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnEval 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Evaluate Subjects"
      Height          =   495
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9000
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ILIS 
      Left            =   240
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEvaluate.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEvaluate.frx":0764
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEvaluate.frx":0BBE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "School Year Information"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7320
      TabIndex        =   10
      Top             =   840
      Width           =   6975
      Begin VB.ComboBox TxtCourse 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2640
         TabIndex        =   3
         Text            =   "Course"
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox TxtSy 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmEvaluate.frx":1018
         Left            =   1320
         List            =   "FrmEvaluate.frx":1049
         TabIndex        =   2
         Text            =   "9999-9999"
         Top             =   240
         Width           =   1215
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
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.CommandButton BtnLoadcur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "L&oad Subjects"
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton BtnLoadSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Load Subjects"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9000
      Width           =   1815
   End
   Begin MSComctlLib.ListView LVSubjectsTaken 
      Height          =   6735
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   11880
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ILIS"
      SmallIcons      =   "ILIS"
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "School Year"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Semester"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Course"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Yr"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Subject"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Description"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Units"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Grade"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView LVCur 
      Height          =   6735
      Left            =   7320
      TabIndex        =   8
      Top             =   1680
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   11880
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ILIS"
      SmallIcons      =   "ILIS"
      ColHdrIcons     =   "ILIS"
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Semester"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Subject Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Units"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Prerequisites"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select subjects above to evaluate (use the CTRL or SHIFT KEY for multiple selection."
      Height          =   495
      Left            =   7320
      TabIndex        =   13
      Top             =   8400
      Width           =   6975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Curriculum Information:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7320
      TabIndex        =   9
      Top             =   600
      Width           =   2490
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subjects Taken"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1650
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   14280
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select A Student"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   7125
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Student to Evaluate:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3330
   End
End
Attribute VB_Name = "FrmEvaluate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   Evaluation Module
'Description    :   Use to evaluate a selected student for a specifice curriculum year.
'                   User must highlight/select subjects from the curriculum display box
'                   and click the evaluate button.
'Date           :   February 28, 2005
'Programmer     :   iehjsuckers
'Comment        :   ?
'*************************************************************************************
Public Student As String
'Public Stud_SY As String, Stud_course As String

Private Sub BtnEval_Click()
If LVCur.SelectedItem Is Nothing Then
    MsgBox "There are no selected subjects to evaluate. Please Select from the list above.", vbInformation, "Evaluation"
    Exit Sub
End If
'try clicking the button load
BtnLoadSub_Click
'continue here
Dim i As Integer, msg As String

For i = 1 To LVCur.ListItems.Count

        If LVCur.ListItems(i).Selected = True Then
            Evaluate LVCur.ListItems(i).SubItems(2), LVCur.ListItems(i).SubItems(5), FrmEvaluationRep.Txt1
        End If
Next
'partition the message
FrmEvaluationRep.Caption = "Evalution Result for " & LblName.Caption
FrmEvaluationRep.Show 1
End Sub

Private Sub BtnLoadcur_Click()
'Try loading the Subjects
LoadCurriculum
End Sub

Private Sub BtnLoadSub_Click()
Dim spl
spl = Split(LblName.Caption, "|", , vbTextCompare)
Idnum = spl(1)
Lnam = spl(2)
Fnam = spl(3)
Mnam = spl(4)
'Set the sy and sem
Dim msg As String
msg = "Select * from SubjectsEnrolled where lnam = '" & Lnam & "' and fnam = '" & Fnam & "' and mnam = '" & Mnam & "' Order by SY,SEM,COURSE,YR"
SetSelectedStudSub msg
LoadALLtoLV LVSubjectsTaken
End Sub

Private Function LoadCurriculum()
LVCur.ListItems.Clear
On Error GoTo Erb
Dim rs As New Recordset
With rs
    Dim msg As String
    msg = "Select * from Curriculum where SY = '" & txtSY.Text & "' and course = '" & txtCourse.Text & "' order by Yr, sem asc"
    .Open msg, DE.Con, adOpenDynamic, adLockOptimistic
    If .RecordCount <> 0 Then
        Do Until .EOF
            LVCur.ListItems.Add .AbsolutePosition, , .Fields("Yr").Value, 3, 3
            LVCur.ListItems(.AbsolutePosition).SubItems(1) = .Fields("Sem").Value
            LVCur.ListItems(.AbsolutePosition).SubItems(2) = .Fields("SC").Value
            LVCur.ListItems(.AbsolutePosition).SubItems(3) = .Fields("Description").Value
            LVCur.ListItems(.AbsolutePosition).SubItems(4) = .Fields("Unts").Value
            LVCur.ListItems(.AbsolutePosition).SubItems(5) = .Fields("Prerequisites").Value
            .MoveNext
        Loop
    End If
End With
Set rs = Nothing
Exit Function
Erb:
MsgBox "Error: " & Err.Description, vbCritical, "Error"
Set rs = Nothing
LVCur.ListItems.Clear
End Function

Private Function LoadALLtoLV(lv As ListView)
lv.ListItems.Clear
With DE.rsSubjectsEnrolled
    If .RecordCount > 0 Then
        Do Until .EOF
            If UCase(Trim(.Fields("Remarks").Value)) <> UCase(Trim("Passed")) Then
                lv.ListItems.Add .AbsolutePosition, , .Fields("SY").Value, 1, 1
            Else
                lv.ListItems.Add .AbsolutePosition, , .Fields("SY").Value, 3, 3
            End If
            lv.ListItems(.AbsolutePosition).SubItems(1) = .Fields("SEM").Value
            lv.ListItems(.AbsolutePosition).SubItems(2) = .Fields("Course").Value
            lv.ListItems(.AbsolutePosition).SubItems(3) = .Fields("YR").Value
            lv.ListItems(.AbsolutePosition).SubItems(4) = .Fields("SC").Value
            lv.ListItems(.AbsolutePosition).SubItems(5) = .Fields("Description").Value
            lv.ListItems(.AbsolutePosition).SubItems(6) = .Fields("Units").Value
            lv.ListItems(.AbsolutePosition).SubItems(7) = .Fields("Grade").Value
            lv.ListItems(.AbsolutePosition).SubItems(8) = .Fields("Remarks").Value
            .MoveNext
        Loop
    End If
End With

End Function

Private Sub Form_Load()
GetCourse
'set if there is a sy and course
End Sub

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

Private Function Evaluate(Subject As String, prere As String, msgx As ListView) As Boolean
Dim Splt
Dim msg As ListItem
'check if the subject is taken before
Dim fnx As ListItem
Set fnx = LVSubjectsTaken.FindItem(Subject, lvwSubItem)

If Not fnx Is Nothing Then
    'check if student passed
    Set msg = msgx.ListItems.Add(, , , 3, 3)
    If Trim(UCase(fnx.SubItems(8))) = "PASSED" Then
        'continue
        msg.Text = Subject
        msg.SubItems(3) = "This has been taken"
    Else
        'create report here
        msg.Text = Subject
        msg.SubItems(3) = "Taken before but not passed. Can still take"
    End If
    Evaluate = True
    Exit Function
End If

If Trim(prere) = "" Then    'you can take subject
    Set msg = msgx.ListItems.Add(, , , 3, 3)
    msg.Text = Subject
    msg.SubItems(3) = "Can Enroll"
    Exit Function
End If
Splt = Split(prere, ", ", , vbTextCompare)
Dim last As Integer
Dim i As Integer
Dim fn As ListItem
For i = LBound(Splt) To UBound(Splt)
    Set fn = LVSubjectsTaken.FindItem(Splt(i), lvwSubItem)
    If fn Is Nothing Then
        Set msg = msgx.ListItems.Add(, , , 1, 1)
        msg.Text = Subject
        msg.SubItems(1) = Splt(i)
        msg.SubItems(2) = "Not Taken"
        msg.SubItems(3) = "Can't Enroll"
        
    Else
        If StrConv(Trim(fn.SubItems(8)), vbProperCase) <> "Passed" Then
            Set msg = msgx.ListItems.Add(, , , 1, 1)
        Else
            Set msg = msgx.ListItems.Add(, , , 3, 3)
        End If
        msg.Text = Subject
        msg.SubItems(1) = Splt(i)
        msg.SubItems(2) = "Taken"
        msg.SubItems(3) = fn.SubItems(8)
        If fn.SubItems(8) <> "Passed" Then
            msg.SubItems(4) = "Not Done"
        Else
            msg.SubItems(4) = "Ok"
        End If
    End If
Next
Evaluate = True
End Function
