VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmbckupandrestor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup & Restore Database"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Backup"
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1155
   End
End
Attribute VB_Name = "frmbckupandrestor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CN1 As New ADODB.Connection

Private Sub Command1_Click()
If Text1.Text = "" Then MsgBox "Please enter backup filename!", vbExclamation, Me.Caption: Exit Sub
CN1.Execute "BACKUP DATABASE iwadco_db TO DISK = '" & Me.Dir1.Path & "\" & Text1.Text & ".bak' WITH INIT, RETAINDAYS = 0"
MsgBox "Backup database complete!", vbInformation, Me.Caption
End Sub

Private Sub Command2_Click()
On Error GoTo errtrap
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
CN1.Execute "RESTORE DATABASE iwadco_db FROM DISK='" & CommonDialog1.FileName & "' WITH MOVE 'iwadco_db_Data' TO '" & Dir1.Path & "\iwadco_db_Data.mdf',MOVE 'iwadco_db_Log' TO '" & Dir1.Path & "\iwadco_db_Log.ldf',REPLACE"
MsgBox "Restore Database Complete!", vbInformation, Me.Caption
errtrap:
Select Case Err.Number
Case 0
Case Else
MsgBox Err.Description, vbExclamation, Me.Caption
Exit Sub
End Select
End Sub

Private Sub Form_Load()
If CN1.State = adStateOpen Then CN1.Close
If fso.FileExists(App.Path & "\config.ini") = True Then
    If loadtxt(App.Path & "\config.ini") = True Then
        CN1.ConnectionString = "DRIVER=SQL Server;SERVER=" & sqlserverdata(0) & ";UID=" & sqlserverdata(1) & ";PWD=" & sqlserverdata(2) & ";APP=Visual Basic;DATABASE=master"
        CN1.Open
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell App.Path & "\prj_iwadco.exe", vbNormalFocus
End
End Sub
