VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCourses 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Courses"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCourses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnAddnew 
      BackColor       =   &H80000004&
      Caption         =   "&AddNew"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "FrmCourses.frx":030A
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5318
      _Version        =   393216
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "Courses"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Course"
         Caption         =   "Course"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Department"
         Caption         =   "Department"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Description"
         Caption         =   "Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4860.284
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   Course Adding Module
'Description    :   Adds courses in your database. (important in selecting subjects,
'                   evaluation, and report generation).
'Date           :   January 15, 2005
'Programmer     :   iehjsuckers
'Comment        :   ?
'***********************************************************************************
Private Sub BtnAddnew_Click()
On Error GoTo Erb
DE.rsCourses.AddNew
Exit Sub
Erb:
    DE.rsCourses.CancelBatch adAffectAllChapters
    DE.rsCourses.Close
    DE.rsCourses.Open
    DG.ReBind
    DG.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo Erb
If DE.rsCourses.State <> 0 Then DE.rsCourses.Close
DE.rsCourses.Open
DG.DataMember = "Courses"
Set DG.DataSource = DE
DG.ReBind
DG.Refresh
DG.Refresh
Exit Sub
Erb:
    DE.rsCourses.CancelBatch adAffectAllChapters
If DE.rsCourses.State <> 0 Then DE.rsCourses.Close
DE.rsCourses.Open
End Sub

Private Sub Form_Resize()
DG.Width = ScaleWidth
DG.Height = ScaleHeight - 500
BtnAddnew.Top = ScaleHeight - BtnAddnew.Height
End Sub

