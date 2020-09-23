VERSION 5.00
Begin VB.Form FrmMoving 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5415
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnCancel 
      BackColor       =   &H80000004&
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton BtnMove 
      BackColor       =   &H80000004&
      Caption         =   "&Move"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Source School Year and Semester"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   5175
      Begin VB.Label LblSem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semester"
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   840
      End
      Begin VB.Label LblSY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School Year"
         Height          =   240
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Subject to Move"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   5175
      Begin VB.Label LblRemarks 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   240
         Left            =   3840
         TabIndex        =   15
         Top             =   240
         Width           =   780
      End
      Begin VB.Label LblGrade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grade"
         Height          =   240
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   525
      End
      Begin VB.Label LblUnits 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         Height          =   240
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   450
      End
      Begin VB.Label LblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.Label LblSC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SC"
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   270
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Destination School Year and Semester"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   5175
      Begin VB.ComboBox TxtSem 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmMoving.frx":0000
         Left            =   3720
         List            =   "FrmMoving.frx":000D
         TabIndex        =   2
         Text            =   "Sem"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox TxtSy 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmMoving.frx":0020
         Left            =   1440
         List            =   "FrmMoving.frx":0051
         TabIndex        =   1
         Text            =   "9999-9999"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Semester:"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   2760
         TabIndex        =   6
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "School Year:"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT SUBJECT MOVING WIZARD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5235
   End
   Begin VB.Label LblStudent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1245
   End
End
Attribute VB_Name = "FrmMoving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   Subject Moving (College Record Add on Module)
'Description    :   Moves a subject of certain student from 1 schoolyear and semester
'               :   to another schoolyear and semester. Active only if a subject is
'                   selected and on College Record Module.
'Date           :   February 26, 2005
'Programmer     :   iehjsuckers
'Comment        :   This will allow users to move subjects if there are mistakes in
'                   adding/updating records instead of deleting them (deleting is
'                    secured so with this one!).
'*************************************************************************************
Private Sub BtnCancel_Click()
Unload Me
End Sub

Private Sub BtnMove_Click()
If CheckExistance = True Then
    CompleteMove
    Unload Me
Else
    MsgBox "Can not complete move.", vbCritical, "Error"
End If
End Sub

'Function For moving
Function CompleteMove()
Dim spl
    spl = Split(LblStudent.Caption, "|", , vbTextCompare)
    Idnum = spl(1)
    Lnam = spl(2)
    Fnam = spl(3)
    Mnam = spl(4)
    
On Error GoTo Erb
    Dim rs As New Recordset
    With rs
        .CursorLocation = adUseClient
        Dim msg As String
        msg = "update SubjectsEnrolled set sy = '" & TxtSy.Text & "', sem = '" & TxtSem.Text & "' " & _
            " where lnam = '" & Lnam & _
            "' and fnam = '" & Fnam & "' and mnam = '" & Mnam & "' and sy = '" & _
            LblSY.Caption & "' and sem = '" & LblSem.Caption & "' and sc = '" & _
            LblSC.Caption & "'"
        .Open msg, DE.Con, adOpenDynamic, adLockOptimistic
        MsgBox "Move Complete.", vbInformation, "Move"
    End With
    Set rs = Nothing
    Exit Function
Erb:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
    Unload Me
End Function

'function check if the record exists if not, return nothing
Function CheckExistance() As Boolean
Dim spl
    spl = Split(LblStudent.Caption, "|", , vbTextCompare)
    Idnum = spl(1)
    Lnam = spl(2)
    Fnam = spl(3)
    Mnam = spl(4)
    
On Error GoTo Erb
    Dim rs As New Recordset
    With rs
        .CursorLocation = adUseClient
        Dim msg As String
        msg = "Select * from SubjectsEnrolled where lnam = '" & Lnam & _
            "' and fnam = '" & Fnam & "' and mnam = '" & Mnam & "' and sy = '" & _
            LblSY.Caption & "' and sem = '" & LblSem.Caption & "' and sc = '" & _
            LblSC.Caption & "'"
        .Open msg, DE.Con, adOpenDynamic, adLockOptimistic
        If .RecordCount = 0 Then
            MsgBox "The record you are trying to move does not exist. Please check your Query.", vbInformation, "Error"
            CheckExistance = False
        Else
            'Continue the move
            CheckExistance = True
        End If
    End With
    Set rs = Nothing
    Exit Function
Erb:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
    CheckExistance = False
End Function
