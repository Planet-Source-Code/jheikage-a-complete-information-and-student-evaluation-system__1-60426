VERSION 5.00
Begin VB.Form Frmlogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form Login"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frmlogin.frx":030A
   ScaleHeight     =   2220
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txtserver 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtpass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtuser 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last System Update: January 27,2005. Check for updates."
      Height          =   210
      Left            =   135
      TabIndex        =   7
      Top             =   2040
      Width           =   4230
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   3600
      Picture         =   "Frmlogin.frx":45B9
      Stretch         =   -1  'True
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   1035
   End
End
Attribute VB_Name = "Frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   Log in
'Description    :   Logs the user in the system. It makes use of SQLSERVER log in
'                   you must set the Server name, a user name and password supplied
'                   by your server administrator.
'Date           :   January 1, 2005
'Programmer     :   iehjsuckers
'Comment        :   ?
'***********************************************************************************
Private Sub cmdok_Click()
Dim x As Boolean
x = connect(txtuser.Text, txtpass.Text, Txtserver.Text)
If x = False Then
    MsgBox "Invalid user...", vbCritical, "Error"
Else
    MyLoginPass = txtpass.Text
    MainForm.ssbr.Panels(2).Text = "Server: " & StrConv(Txtserver.Text, vbProperCase)
    MainForm.ssbr.Panels(3).Text = "Logged User: " & StrConv(txtuser.Text, vbProperCase)
    MainForm.Show
    Unload Me
End If
End Sub

Private Sub Form_Load()
retriveRegs
End Sub

Private Function retriveRegs()
    Dim server As String, user As String
    server = GetSetting("CSUMIS", "INFO", "Server", "SQL SERVER NAME")
    user = GetSetting("CSUMIS", "INFO", "User", "User")
    Txtserver.Text = server
    txtuser.Text = user
End Function

Private Sub txtpass_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdok_Click
End Sub

Private Sub Txtserver_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub Txtserver_LostFocus()
    SaveSetting "CSUMIS", "INFO", "Server", Txtserver.Text
End Sub

Private Sub txtuser_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtuser_LostFocus()
    SaveSetting "CSUMIS", "INFO", "User", txtuser.Text
End Sub
