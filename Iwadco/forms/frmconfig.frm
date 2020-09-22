VERSION 5.00
Begin VB.Form frmSetDB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Settings"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDb 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin Project1.isButton cmdcancel 
      Height          =   420
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   741
      Icon            =   "frmconfig.frx":0000
      Style           =   5
      Caption         =   "&Cancel"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.isButton cmdsave 
      Height          =   420
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   741
      Icon            =   "frmconfig.frx":001C
      Style           =   5
      Caption         =   "&Connect"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox TxtPass 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtConDb 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Database:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "UserName :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SQL-Server:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "frmSetDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
End
End Sub

Private Sub cmdsave_Click()
Dim errMsg As String, noErr As String
errMsg = "Please complete the following requirements!" & Chr(10) & "-----------------------------"
noErr = errMsg
If txtConDb.Text = "" Then
    errMsg = errMsg & Chr(10) & "Server name"
End If
If txtDb.Text = "" Then
    errMsg = errMsg & Chr(10) & "Database"
End If
If errMsg <> noErr Then
    MsgBox errMsg, vbExclamation, Me.Caption
    Exit Sub
End If
Set txt = fso.CreateTextFile(App.Path & "\config.ini", True)
Set fileforopen = fso.GetFile(App.Path & "\config.Ini")
txt.WriteLine Encode(txtConDb.Text & ":" & txtUser.Text & ":" & TxtPass.Text & ":" & txtDb.Text)
txt.Close
Set fso = Nothing
Set txt = Nothing
Set fileforopen = Nothing
Unload Me
End Sub

