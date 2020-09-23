VERSION 5.00
Begin VB.Form FrmPIS 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personal Information"
   ClientHeight    =   10095
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11400
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
   Icon            =   "FrmPIS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox TxtSy 
      Appearance      =   0  'Flat
      DataField       =   "curriculum"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      ItemData        =   "FrmPIS.frx":030A
      Left            =   4440
      List            =   "FrmPIS.frx":033B
      TabIndex        =   81
      Text            =   "9999-9999"
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox TxtIncome 
      DataField       =   "Income"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      ItemData        =   "FrmPIS.frx":03E4
      Left            =   8760
      List            =   "FrmPIS.frx":03FD
      TabIndex        =   27
      Text            =   "A"
      Top             =   6000
      Width           =   2535
   End
   Begin VB.ComboBox TxtCollege 
      Appearance      =   0  'Flat
      DataField       =   "College"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      ItemData        =   "FrmPIS.frx":0416
      Left            =   1080
      List            =   "FrmPIS.frx":0423
      TabIndex        =   13
      Text            =   "College"
      Top             =   3000
      Width           =   4455
   End
   Begin VB.ComboBox TxtCourse 
      Appearance      =   0  'Flat
      DataField       =   "Course"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   6360
      TabIndex        =   14
      Text            =   "Course"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ComboBox TxtCS 
      Appearance      =   0  'Flat
      DataField       =   "CS"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      ItemData        =   "FrmPIS.frx":0444
      Left            =   7080
      List            =   "FrmPIS.frx":0451
      TabIndex        =   11
      Text            =   "S"
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox TxtSex 
      Appearance      =   0  'Flat
      DataField       =   "Sex"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "F"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      ItemData        =   "FrmPIS.frx":045E
      Left            =   5160
      List            =   "FrmPIS.frx":0468
      TabIndex        =   10
      Text            =   "Sex"
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton BtnCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton BtnOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "&Okay"
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   9480
      Width           =   1215
   End
   Begin VB.ComboBox TxtOccupation 
      Appearance      =   0  'Flat
      DataField       =   "Occupation"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      ItemData        =   "FrmPIS.frx":0472
      Left            =   8160
      List            =   "FrmPIS.frx":049D
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   4080
      Width           =   3135
   End
   Begin VB.ComboBox TxtBoarding 
      Appearance      =   0  'Flat
      DataField       =   "boarding"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      ItemData        =   "FrmPIS.frx":0568
      Left            =   1560
      List            =   "FrmPIS.frx":0575
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   5040
      Width           =   3015
   End
   Begin VB.ComboBox TxtMoccup 
      Appearance      =   0  'Flat
      DataField       =   "moccup"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      ItemData        =   "FrmPIS.frx":05A0
      Left            =   6000
      List            =   "FrmPIS.frx":05CB
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   5520
      Width           =   2655
   End
   Begin VB.ComboBox TxtFoccup 
      Appearance      =   0  'Flat
      DataField       =   "foccup"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      ItemData        =   "FrmPIS.frx":0696
      Left            =   6000
      List            =   "FrmPIS.frx":06DC
      TabIndex        =   26
      Text            =   "Combo1"
      Top             =   6000
      Width           =   2655
   End
   Begin VB.ComboBox txtrel 
      Appearance      =   0  'Flat
      DataField       =   "rel"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      ItemData        =   "FrmPIS.frx":0862
      Left            =   10320
      List            =   "FrmPIS.frx":088D
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton BPrev 
      Appearance      =   0  'Flat
      Caption         =   "<"
      Height          =   495
      Left            =   8880
      TabIndex        =   74
      Top             =   9480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Bnext 
      Appearance      =   0  'Flat
      Caption         =   ">"
      Height          =   495
      Left            =   9720
      TabIndex        =   73
      Top             =   9480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Bbof 
      Appearance      =   0  'Flat
      Caption         =   "<<"
      Height          =   495
      Left            =   8040
      TabIndex        =   72
      Top             =   9480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Beof 
      Appearance      =   0  'Flat
      Caption         =   ">>"
      Height          =   495
      Left            =   10320
      TabIndex        =   71
      Top             =   9480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Bdelete 
      Appearance      =   0  'Flat
      Caption         =   "Delete"
      Height          =   495
      Left            =   6120
      TabIndex        =   70
      Top             =   9480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Bsave 
      Appearance      =   0  'Flat
      Caption         =   "Save"
      Height          =   495
      Left            =   4680
      TabIndex        =   69
      Top             =   9480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Badd 
      Appearance      =   0  'Flat
      Caption         =   "Add"
      Height          =   495
      Left            =   3240
      TabIndex        =   68
      Top             =   9480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCschladd 
      Appearance      =   0  'Flat
      DataField       =   "Cschladd"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1560
      TabIndex        =   34
      Top             =   8880
      Width           =   9735
   End
   Begin VB.TextBox txtyrgrad 
      Appearance      =   0  'Flat
      DataField       =   "yrgrad"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   5640
      TabIndex        =   33
      Top             =   8400
      Width           =   900
   End
   Begin VB.TextBox txtcoll 
      Appearance      =   0  'Flat
      DataField       =   "coll"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1080
      TabIndex        =   32
      Top             =   8400
      Width           =   2970
   End
   Begin VB.TextBox txtSschladd 
      Appearance      =   0  'Flat
      DataField       =   "Sschladd"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1560
      TabIndex        =   31
      Top             =   7920
      Width           =   9735
   End
   Begin VB.TextBox txtyrgrduated 
      Appearance      =   0  'Flat
      DataField       =   "yrgrduated"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   5640
      TabIndex        =   30
      Top             =   7440
      Width           =   900
   End
   Begin VB.TextBox txtscondary 
      Appearance      =   0  'Flat
      DataField       =   "scondary"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1080
      TabIndex        =   29
      Top             =   7485
      Width           =   3015
   End
   Begin VB.TextBox txtnmelanlord 
      Appearance      =   0  'Flat
      DataField       =   "nmelanlord"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   6840
      TabIndex        =   22
      Top             =   5025
      Width           =   3375
   End
   Begin VB.TextBox txtaddr 
      Appearance      =   0  'Flat
      DataField       =   "addr"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1080
      TabIndex        =   28
      Top             =   6480
      Width           =   10215
   End
   Begin VB.TextBox txtnmefther 
      Appearance      =   0  'Flat
      DataField       =   "nmefther"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1080
      TabIndex        =   25
      Top             =   6000
      Width           =   3495
   End
   Begin VB.TextBox txtnmemther 
      Appearance      =   0  'Flat
      DataField       =   "nmemther"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1080
      TabIndex        =   23
      Top             =   5535
      Width           =   3495
   End
   Begin VB.TextBox txtRaddress 
      Appearance      =   0  'Flat
      DataField       =   "Raddress"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1080
      TabIndex        =   20
      Top             =   4560
      Width           =   10215
   End
   Begin VB.TextBox txtrlntship 
      Appearance      =   0  'Flat
      DataField       =   "rlntship"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   5280
      TabIndex        =   18
      Top             =   4095
      Width           =   1650
   End
   Begin VB.TextBox txtGuardian 
      Appearance      =   0  'Flat
      DataField       =   "Guardian"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1080
      TabIndex        =   17
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox txtHAddress 
      Appearance      =   0  'Flat
      DataField       =   "HAddress"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1485
      TabIndex        =   7
      Top             =   1725
      Width           =   9765
   End
   Begin VB.TextBox txtNationality 
      Appearance      =   0  'Flat
      DataField       =   "Nationality"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   7785
      TabIndex        =   5
      Top             =   1080
      Width           =   1650
   End
   Begin VB.TextBox txtMinor 
      Appearance      =   0  'Flat
      DataField       =   "Minor"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   10140
      TabIndex        =   16
      Top             =   3000
      Width           =   1110
   End
   Begin VB.TextBox txtPlacebirth 
      Appearance      =   0  'Flat
      DataField       =   "Placebirth"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1050
      TabIndex        =   12
      Top             =   2550
      Width           =   4800
   End
   Begin VB.TextBox txtBday 
      Appearance      =   0  'Flat
      DataField       =   "Bday"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1065
      TabIndex        =   8
      Top             =   2160
      Width           =   2505
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      DataField       =   "Age"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   4170
      TabIndex        =   9
      Top             =   2160
      Width           =   570
   End
   Begin VB.TextBox txtMajor 
      Appearance      =   0  'Flat
      DataField       =   "Major"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   8370
      TabIndex        =   15
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtmnam 
      Appearance      =   0  'Flat
      DataField       =   "mnam"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   4560
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtfnam 
      Appearance      =   0  'Flat
      DataField       =   "fnam"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtlnam 
      Appearance      =   0  'Flat
      DataField       =   "lnam"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtidnum 
      Appearance      =   0  'Flat
      DataField       =   "idnum"
      DataMember      =   "PIS"
      DataSource      =   "DE"
      Height          =   360
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Curriculum to follow:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   2640
      TabIndex        =   80
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   " SCHOOLS ATTENDED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   79
      Top             =   6960
      Width           =   11175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   " FAMILY BACKGROUND"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   78
      Top             =   3600
      Width           =   11175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   " BASIC INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   77
      Top             =   120
      Width           =   11175
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Family Income"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   8760
      TabIndex        =   76
      Top             =   5640
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Civil Status:"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6000
      TabIndex        =   75
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "School Address"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   35
      Left            =   135
      TabIndex        =   67
      Top             =   8880
      Width           =   1380
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Year Graduated:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   34
      Left            =   4200
      TabIndex        =   66
      Top             =   8400
      Width           =   1425
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "College"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   33
      Left            =   360
      TabIndex        =   65
      Top             =   8460
      Width           =   645
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "School Address"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   32
      Left            =   120
      TabIndex        =   64
      Top             =   7920
      Width           =   1380
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Year Graduated:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   31
      Left            =   4200
      TabIndex        =   63
      Top             =   7485
      Width           =   1425
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Secondary"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   30
      Left            =   120
      TabIndex        =   62
      Top             =   7530
      Width           =   930
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name Lanlord/Landlady"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   29
      Left            =   4680
      TabIndex        =   61
      Top             =   5070
      Width           =   2040
   End
   Begin VB.Label lblFieldLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Boarding House"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   28
      Left            =   135
      TabIndex        =   60
      Top             =   5010
      Width           =   1380
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   27
      Left            =   9480
      TabIndex        =   59
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   26
      Left            =   240
      TabIndex        =   58
      Top             =   6480
      Width           =   720
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   25
      Left            =   4920
      TabIndex        =   57
      Top             =   6000
      Width           =   990
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Father"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   24
      Left            =   360
      TabIndex        =   56
      Top             =   6120
      Width           =   555
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   23
      Left            =   4920
      TabIndex        =   55
      Top             =   5520
      Width           =   990
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mother"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   22
      Left            =   360
      TabIndex        =   54
      Top             =   5580
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   21
      Left            =   7080
      TabIndex        =   53
      Top             =   4080
      Width           =   1050
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   20
      Left            =   240
      TabIndex        =   52
      Top             =   4560
      Width           =   720
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Relationship:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   19
      Left            =   4080
      TabIndex        =   51
      Top             =   4080
      Width           =   1125
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   18
      Left            =   120
      TabIndex        =   50
      Top             =   4125
      Width           =   840
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Address"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   17
      Left            =   120
      TabIndex        =   49
      Top             =   1755
      Width           =   1290
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   16
      Left            =   6735
      TabIndex        =   48
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sex:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   15
      Left            =   4800
      TabIndex        =   47
      Top             =   2205
      Width           =   405
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Course:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   14
      Left            =   5610
      TabIndex        =   46
      Top             =   3030
      Width           =   675
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "College:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   13
      Left            =   240
      TabIndex        =   45
      Top             =   3030
      Width           =   705
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Minor:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   12
      Left            =   9555
      TabIndex        =   44
      Top             =   3000
      Width           =   540
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Placebirth:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   10
      Left            =   90
      TabIndex        =   43
      Top             =   2610
      Width           =   930
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   9
      Left            =   240
      TabIndex        =   42
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   8
      Left            =   3690
      TabIndex        =   41
      Top             =   2205
      Width           =   405
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Major:"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   11
      Left            =   7800
      TabIndex        =   40
      Top             =   3000
      Width           =   540
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   6
      Left            =   4920
      TabIndex        =   39
      Top             =   1440
      Width           =   1140
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   2880
      TabIndex        =   38
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   480
      TabIndex        =   37
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ID number"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   885
   End
End
Attribute VB_Name = "FrmPIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   Personal Information Module (SPI Add on Module)
'Description    :   This will allow adding and updating of Personal Information
'                   of students. You can set all the necessary information here.
'                   This is active only if you try to add/udate a student record
'                   on the SPI, EVALUATION, COLLEGE RECORD Main Module (the button
'                   Add and Edit Students will make this form available).
'Date           :   January 5, 2005
'Programmer     :   iehjsuckers
'Last Update    :   February 2, 2005
'Comment        :   For further development ???!
'*********************************************************************************
'user defined functions
Function getcourses()
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
End Function
'end of user defined functions

Private Sub Badd_Click()
On Error GoTo xx
DE.rsPIS.AddNew
Badd.Enabled = False
Bdelete.Caption = "Cancel"
Beof.Enabled = False
Bbof.Enabled = False
BPrev.Enabled = False
Bnext.Enabled = False
'txtidnum.SetFocus
Exit Sub
xx:
MsgBox Err.Description, vbCritical, "Error"
Badd.Enabled = True
Bdelete.Caption = "Delete"
Beof.Enabled = True
Bbof.Enabled = True
BPrev.Enabled = True
Bnext.Enabled = True
SetRs "Select * from names"
End Sub

Private Sub Bbof_Click()
  DE.rsPIS.MoveFirst
  Me.Caption = "Personal Information: Record " & DE.rsPIS.AbsolutePosition & " of " & DE.rsPIS.RecordCount
End Sub

Private Sub BDelete_Click()
On Error GoTo err_handler

If Bdelete.Caption = "Cancel" Then
    Badd.Enabled = True
    Bdelete.Caption = "Delete"
    Beof.Enabled = True
    Bbof.Enabled = True
    BPrev.Enabled = True
    Bnext.Enabled = True
    DE.rsPIS.CancelBatch adAffectAllChapters
Else
    MsgBox "    Do you want to Delete?", vbYesNo + vbQuestion, "Confirmation"
    If vbYes Then
        DE.rsPIS.Delete
    End If
End If
SetRs "Select * from names"
Exit Sub
err_handler:
MsgBox Err.Description, vbCritical, "Error"
Badd.Enabled = True
Bdelete.Caption = "Delete"
Beof.Enabled = True
Bbof.Enabled = True
BPrev.Enabled = True
Bnext.Enabled = True
SetRs "Select * from names"
End Sub

Private Sub Beof_Click()
  DE.rsPIS.MoveLast
  Me.Caption = "Personal Information: Record " & DE.rsPIS.AbsolutePosition & " of " & DE.rsPIS.RecordCount
End Sub

Private Sub Bnext_Click()
If DE.rsPIS.EOF Then DE.rsPIS.MoveLast Else DE.rsPIS.MoveNext
Me.Caption = "Personal Information: Record " & DE.rsPIS.AbsolutePosition & " of " & DE.rsPIS.RecordCount
End Sub

Private Sub BPrev_Click()
  If DE.rsPIS.BOF Then DE.rsPIS.MoveFirst Else DE.rsPIS.MovePrevious
  Me.Caption = "Personal Information: Record " & DE.rsPIS.AbsolutePosition & " of " & DE.rsPIS.RecordCount
End Sub

Private Sub Bsave_Click()
On Error GoTo err_handler
Retrim
DE.rsPIS.Update
Badd.Enabled = True
Bdelete.Caption = "Delete"
Beof.Enabled = True
Bbof.Enabled = True
BPrev.Enabled = True
Bnext.Enabled = True
MsgBox "Record Saved. ", vbInformation, "Message"
DE.rsPIS.Close
SetRs "Select * from names"
'update also all the files in subjects enrolled
Dim rs As New Recordset
With rs
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Update SubjectsEnrolled set lnam = '" & txtlnam.Text & "' ,fnam='" & txtfnam.Text & "' , mnam = '" & txtmnam.Text & "' , IDnum = '" & txtIdnum.Text & "' where " & _
        " mnam = '" & txtmnam.Text & "' and fnam = '" & txtfnam.Text & "' and lnam = '" & txtlnam.Text & "'", DE.Con
End With
Set rs = Nothing
Exit Sub
err_handler:
MsgBox Err.Description, vbCritical, "Error"
Badd.Enabled = True
Bdelete.Caption = "Delete"
Beof.Enabled = True
Bbof.Enabled = True
BPrev.Enabled = True
Bnext.Enabled = True
DE.rsPIS.Cancel
SetRs "Select * from names"
End Sub

Private Sub BtnCancel_Click()
    If Bdelete.Caption = "Cancel" Then BDelete_Click
    Unload Me
End Sub

Private Sub BtnOk_Click()
    Bsave_Click
    Unload Me
End Sub

Private Sub Form_Load()
    Select Case FrmMainPIS.isadd
    Case True   'addnew
        Badd_Click
        clearText
    Case False  'load something
    End Select
    'set sex
    TxtSex.Clear
    TxtSex.AddItem "M"
    TxtSex.AddItem "F"
    'set status
    TxtCS.Clear
    TxtCS.AddItem "S"
    TxtCS.AddItem "M"
    TxtCS.AddItem "W"
    'set department
    TxtCollege.Clear
    TxtCollege.AddItem "CAS"
    TxtCollege.AddItem "EDUCATION"
    TxtCollege.AddItem "AGRICULTURE"
    'set income
    TxtIncome.Clear
    TxtIncome.AddItem "A 240000-up"
    TxtIncome.AddItem "B 19000-239000"
    TxtIncome.AddItem "C 13000-189000"
    TxtIncome.AddItem "D 80000-129000"
    TxtIncome.AddItem "E 40000-79000"
    TxtIncome.AddItem "F Below 39000"
    getcourses
    ReBind
End Sub

Private Sub txtBday_LostFocus()
On Error GoTo Erb
    If IsDate(txtBday.Text) Or Trim(txtBday.Text) = "" Then
        'compute the age
        Dim dte As String, sp
        dte = Format(txtBday.Text, "mm/dd/yyyy")
        sp = Split(dte, "/", , vbTextCompare)
        Dim myage As Integer
        myage = Int(Format(Date, "yyyy")) - sp(2)
        If Int(sp(0)) > Int(Format(Date, "mm")) Then
            myage = myage - 1
        Else
            If Int(sp(0)) = Int(Format(Date, "mm")) Then
                If Int(sp(1)) < Int(Format(Date, "dd")) Then
                    myage = myage - 1
                End If
            End If
        End If
        txtAge.Text = myage
    Else
        MsgBox "Please specify a valid date.", vbInformation, "Error"
        txtBday.SetFocus
        SendKeys "{HOME}+{END}"
    End If
    
    Exit Sub
Erb:
    'Goto nothing
    MsgBox "Type a valid date please.", vbCritical, "Error"
    txtBday.SetFocus
    SendKeys "{HOME}+{END}"
End Sub

