VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Student Management Information System HX Technologies"
   ClientHeight    =   8850
   ClientLeft      =   90
   ClientTop       =   870
   ClientWidth     =   11850
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
   ForeColor       =   &H00404040&
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar ssbr 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   8475
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Picture         =   "MainForm.frx":030A
            Text            =   "MIS"
            TextSave        =   "MIS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3889
            Picture         =   "MainForm.frx":073E
            Text            =   "Server Information:"
            TextSave        =   "Server Information:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   4101
            Picture         =   "MainForm.frx":0A58
            Text            =   "Logged User:"
            TextSave        =   "Logged User:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   10239
            Picture         =   "MainForm.frx":10CC
            Text            =   "HX® Technologies:e-mail: iehjsucker@yahoo.com"
            TextSave        =   "HX® Technologies:e-mail: iehjsucker@yahoo.com"
         EndProperty
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
   End
   Begin MIS.chameleonButton BtnPIS 
      Height          =   1800
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3175
      BOLD            =   0   'False
      ITA             =   0   'False
      INAME           =   "MS Sans Serif"
      SIZE            =   0
      UND             =   0   'False
      BTYPE           =   3
      TX              =   "STUDENT PIS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   4
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MIS.chameleonButton BtnCur 
      Height          =   1800
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3175
      BOLD            =   0   'False
      ITA             =   0   'False
      INAME           =   "MS Sans Serif"
      SIZE            =   0
      UND             =   0   'False
      BTYPE           =   3
      TX              =   "CURRICULA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   4
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MIS.chameleonButton BtnCourse 
      Height          =   1800
      Left            =   3960
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3175
      BOLD            =   0   'False
      ITA             =   0   'False
      INAME           =   "MS Sans Serif"
      SIZE            =   0
      UND             =   0   'False
      BTYPE           =   3
      TX              =   "COURSES"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   4
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MIS.chameleonButton BtnReports 
      Height          =   1800
      Left            =   5880
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3175
      BOLD            =   0   'False
      ITA             =   0   'False
      INAME           =   "MS Sans Serif"
      SIZE            =   0
      UND             =   0   'False
      BTYPE           =   3
      TX              =   "REPORTS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   4
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MIS.chameleonButton BtnINFO 
      Height          =   1800
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3175
      BOLD            =   0   'False
      ITA             =   0   'False
      INAME           =   "MS Sans Serif"
      SIZE            =   0
      UND             =   0   'False
      BTYPE           =   3
      TX              =   "SYSTEM INFO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   4
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MIS.chameleonButton BtnQUit 
      Height          =   735
      Left            =   9360
      TabIndex        =   8
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      BOLD            =   0   'False
      ITA             =   0   'False
      INAME           =   "MS Sans Serif"
      SIZE            =   0
      UND             =   0   'False
      BTYPE           =   3
      TX              =   "QUIT"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   4
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MIS.chameleonButton btnEvaluation 
      Height          =   1800
      Left            =   7800
      TabIndex        =   11
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3175
      BOLD            =   0   'False
      ITA             =   0   'False
      INAME           =   "MS Sans Serif"
      SIZE            =   0
      UND             =   0   'False
      BTYPE           =   3
      TX              =   "EVALUATE STUDENTS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   4
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MIS.chameleonButton BtnSubOffered 
      Height          =   1800
      Left            =   9720
      TabIndex        =   14
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3175
      BOLD            =   0   'False
      ITA             =   0   'False
      INAME           =   "MS Sans Serif"
      SIZE            =   0
      UND             =   0   'False
      BTYPE           =   3
      TX              =   "Subjects Offered"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   4
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblHx 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HX Tech® Building Outstanding Technology"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   600
      TabIndex        =   13
      Top             =   7800
      Width           =   4170
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MISCELLANEOUS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   3120
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "APPLICATION MODULES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   120
      Picture         =   "MainForm.frx":163F
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label l3 
      BackColor       =   &H00404080&
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label l2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MANAGEMENT INFORMATION SYSTEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8175
   End
   Begin VB.Label L1 
      BackColor       =   &H00FF0000&
      Caption         =   "School"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   5055
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   Main Form
'Description    :   Provides access for the entire system. Active only if you
'                   have logged in the system.
'Date           :   January 3, 2005
'Programmer     :   iehjsuckers
'Comment        :   Further Designs!!
'***********************************************************************************
Private Sub BtnCourse_Click()
FrmCourses.Show 1
End Sub

Private Sub BtnCur_Click()
FrmCurs.Show 1
End Sub

Private Sub btnEvaluation_Click()
MsgBox "The Evaluation Module is Transfered at Student PI Module." & vbNewLine & "Click the Button Evaluate Student in SPI Module." & vbNewLine & "Thank you.", vbInformation, "Module Transfered"
End Sub

Private Sub BtnINFO_Click()
FrmAbout.Show 1
'MsgBox "Hindi pa po tapos ang Sistema namin. Paki antay na lang po.", vbInformation, "Mensahe"
End Sub

Private Sub BtnPIS_Click()
FrmMainPIS.Show 1
End Sub

Private Sub BtnQuit_Click()
End
End Sub

Private Sub BtnReports_Click()
    'MsgBox "Not Implemented yet Gwapo.", vbInformation, "Message"
    FrmReports.Show 1
End Sub

Private Sub BtnSubOffered_Click()
    FrmSubOffers.Show 1, Me
End Sub

Private Sub Form_Load()
Tricks
End Sub

Private Sub Form_Resize()
    L1.Width = Me.ScaleWidth
    l2.Width = Me.ScaleWidth
    l3.Width = Me.ScaleWidth
    BtnQUit.Left = Me.ScaleWidth - BtnQUit.Width - 10
    BtnQUit.Top = Me.ScaleHeight - BtnQUit.Height - 30
    Image1.Top = Me.ScaleHeight - Image1.Height - 50
    LblHx.Left = Image1.Left
    LblHx.Top = Image1.Top + Image1.Height
End Sub

Private Sub Tricks()
Dim mov As Integer
Me.ScaleMode = vbPixels
Me.DrawWidth = 2
Me.DrawMode = vbCopyPen
Dim i As Integer
For i = 0 To 255
    Me.Line (0, mov)-(Me.ScaleWidth + 1000, mov), RGB(0, 0, i), B
    mov = mov + 2
Next
Dim j As Integer
For j = 255 To 0 Step -1
    Me.Line (0, mov)-(Me.ScaleWidth + 1000, mov), RGB(0, 0, j), B
    mov = mov + 2
Next
End Sub
