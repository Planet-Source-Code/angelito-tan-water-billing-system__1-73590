VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3240
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   3240
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   2880
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   3000
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   0
      MousePointer    =   11  'Hourglass
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer

Private Sub Form_Load()
If App.PrevInstance = True Then
    MsgBox "Iwadco System Is Already Open", vbInformation, "Iwadco System"
    End
End If
X = 0
lblStatus.Caption = "Initializing required datus..."
Set enc_md5 = New MD5
lblStatus.Caption = "Connecting to Server Database..."
End Sub






Private Sub Timer1_Timer()
If X = 0 Then
    If fso.FileExists(App.Path & "\config.ini") = True Then
        If loadtxt(App.Path & "\config.ini") = True Then
            X = 1
            If DBConnect(sqlserverdata(0), sqlserverdata(1), sqlserverdata(2), sqlserverdata(3)) = False Then MsgBox "Remote server not found or Access Denied!", vbExclamation, Me.Caption: frmSetDB.Show 1
        End If
    Else
        X = -1
        frmSetDB.Show 1
    End If
Else
    If DBConnect(sqlserverdata(0), sqlserverdata(1), sqlserverdata(2), sqlserverdata(3)) = True Then
        lblStatus.Caption = "Complete..."
        rsCheck
        rs.Open "SELECT * FROM iwadco_user WHERE status='E'", CN, adOpenStatic, adLockOptimistic
        If rs.RecordCount = 0 Then
            frmAddUser.Show
        Else
            'frmLogin.Show
            frmMain.Show
        End If
        Unload Me
        Timer1.Enabled = False
    Else
        X = -1
        frmSetDB.txtConDb.Text = sqlserverdata(0)
        frmSetDB.txtUser.Text = sqlserverdata(1)
        frmSetDB.TxtPass.Text = sqlserverdata(2)
        frmSetDB.txtDb.Text = sqlserverdata(3)
        Kill App.Path & "\config.ini"
        MsgBox "Remote server not found or Access Denied!", vbExclamation, Me.Caption
    End If
End If
X = X + 1
End Sub
