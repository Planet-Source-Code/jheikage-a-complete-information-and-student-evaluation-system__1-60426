VERSION 5.00
Begin VB.Form FrmConfirm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Confirmation Box:Please enter your password"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3855
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
   ScaleHeight     =   480
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCon 
      Appearance      =   0  'Flat
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "FrmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module Name    :   Confirmation Module
'Description    :   This will confirm the identity of the user (appears only if you
'                   want to delete or move a certain object/record in your database.
'Date           :   February 20, 2005
'Programmer     :   iehjsuckers
'Last Update    :   February 28, 2005
'Comment        :   ?
'***********************************************************************************
Private Sub Form_Load()
IsTrue = False
End Sub

Private Sub TxtCon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'check here
    If TxtCon.Text = MyLoginPass Then
        IsTrue = True
    End If
    Unload Me
End If
End Sub
