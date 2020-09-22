VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Login"
   ClientHeight    =   2235
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4755
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320.512
   ScaleMode       =   0  'User
   ScaleWidth      =   4464.687
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   100
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "Name"
      Text            =   "tan"
      Top             =   1200
      Width           =   2835
   End
   Begin VB.ComboBox cmbUser 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "tan"
      Top             =   840
      Width           =   2895
   End
   Begin Project1.isButton cmdLogin 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      icon            =   "frmLogin.frx":169B2
      style           =   5
      caption         =   "Login"
      iconsize        =   18
      captionalign    =   2
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frmLogin.frx":1728E
   End
   Begin Project1.isButton cmdCancel 
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      icon            =   "frmLogin.frx":172B6
      style           =   5
      caption         =   "Cancel"
      iconsize        =   17
      captionalign    =   2
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frmLogin.frx":2C42A
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select your username and enter your password in the space provided bellow."
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   3315
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   0
      Left            =   345
      TabIndex        =   4
      Top             =   870
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   120
      Picture         =   "frmLogin.frx":2C452
      Stretch         =   -1  'True
      Top             =   120
      Width           =   585
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
End
End Sub

Private Sub cmdLogin_Click()
If cmbUser.Text = "" Then
    MsgBox "Please input your username!", vbExclamation, Me.Caption
    Exit Sub
End If
If txtPassword.Text = "" Then
    MsgBox "Please input your password!", vbExclamation, Me.Caption
    Exit Sub
End If
If logInSucceeded(cmbUser.Text, txtPassword.Text) = True Then
    Unload Me
Else
    MsgBox "Invalid Login!", vbCritical, Me.Caption
    cmbUser.SetFocus
End If
End Sub

Private Sub Form_Load()
cmdLogin.CaptionAlign = isbright
cmdcancel.CaptionAlign = isbright
sql = "Select username from iwadco_user"
'Call LoadCombo(sql, cmbUser)
Call UnloadAllExceptOne(Me.Name)
End Sub

Private Sub txtPassword_GotFocus()
On Error Resume Next
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdLogin_Click
End If
End Sub

Private Sub txtUserName_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPassword.SetFocus
End If
End Sub
