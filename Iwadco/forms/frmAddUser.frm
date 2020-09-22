VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMadduser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add User Account"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   8655
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         MaxLength       =   15
         TabIndex        =   8
         Top             =   240
         Width           =   2325
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   6120
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   630
         Width           =   2325
      End
      Begin VB.TextBox txtConfirm 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   6120
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1030
         Width           =   2325
      End
      Begin VB.TextBox txtLName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   1
         Top             =   240
         Width           =   2325
      End
      Begin VB.TextBox txtFName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   2
         Top             =   630
         Width           =   2325
      End
      Begin VB.TextBox txtMName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1030
         Width           =   2325
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1440
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1420
         Width           =   3285
      End
      Begin VB.TextBox txtPhone 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2170
         Width           =   2325
      End
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2560
         Width           =   2325
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   7
         Top             =   2990
         Width           =   2325
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1420
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Top             =   1830
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   89522177
         CurrentDate     =   39392
      End
      Begin Project1.isButton cmdUpdate 
         Height          =   450
         Left            =   1440
         TabIndex        =   13
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   794
         Icon            =   "frmAddUser.frx":6852
         Style           =   5
         Caption         =   "Update"
         IconSize        =   22
         IconAlign       =   1
         CaptionAlign    =   2
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.isButton cmdCancel 
         Height          =   450
         Left            =   2880
         TabIndex        =   14
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   794
         Icon            =   "frmAddUser.frx":712C
         Style           =   5
         Caption         =   "Cancel"
         IconSize        =   22
         IconAlign       =   1
         CaptionAlign    =   2
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Left            =   4920
         TabIndex        =   26
         Top             =   360
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
         Left            =   4920
         TabIndex        =   25
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm:"
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
         Index           =   2
         Left            =   4920
         TabIndex        =   24
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type:"
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
         Index           =   3
         Left            =   4920
         TabIndex        =   23
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
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
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
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
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name:"
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
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
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
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
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
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile:"
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
         Height          =   240
         Index           =   9
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   690
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
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
         Height          =   240
         Index           =   10
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   555
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Hired:"
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
         Index           =   11
         Left            =   4920
         TabIndex        =   15
         Top             =   1920
         Width           =   1560
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmAddUser.frx":1E1C6
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Add, Edit Account Profile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1080
      TabIndex        =   28
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Account Profile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   960
      TabIndex        =   27
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
Dim errMsg As String, noErr As String
errMsg = "Please complete the following requirements!" & Chr(10) & "-----------------------------"
noErr = errMsg
If txtUserName.Text = "" Then
    errMsg = errMsg & Chr(10) & "Username"
End If
If txtPassword = "" Then
    errMsg = errMsg & Chr(10) & "Password"
End If
If txtPassword.Text <> txtConfirm.Text Then
    errMsg = errMsg & Chr(10) & "Password no match."
End If
If errMsg <> noErr Then
    MsgBox errMsg, vbExclamation, Me.Caption
    Exit Sub
End If
rsCheck
'ichcheck kung ung new username na ininput ay nagamit na ng ibang users
If frmSysUser.ListView1.ListItems.Count > 1 Then
    rs.Open "SELECT * FROM iwadco_user WHERE username ='" & txtUserName.Text & "' AND id <> " & frmSysUser.ListView1.SelectedItem.Text & " AND status='E'", CN, adOpenStatic, adLockOptimistic
End If
If rs.RecordCount <> 0 Then
    MsgBox "The username you have entered is used by another employee.", vbInformation, Me.Caption
    Exit Sub
End If
If rs.RecordCount = 0 Then
    rsCheck
    If Me.Caption = "Edit User Account" Then
        sql = "SELECT * FROM iwadco_user WHERE id = " & frmSysUser.ListView1.SelectedItem.Text & " AND status='E'"
        rs.Open sql
    Else
        rs.Open "iwadco_user"
        rs.AddNew
    End If
    rs(1) = txtUserName.Text
    rs(2) = LCase(enc_md5.DigestStrToHexStr(txtPassword.Text))
    rs(3) = Combo1.ListIndex + 1
    rs(4) = txtLname.Text
    rs(5) = txtfname.Text
    rs(6) = txtmname.Text
    rs(7) = txtAddress.Text
    rs(8) = txtPhone.Text
    rs(9) = txtMobile.Text
    rs(10) = txtemail.Text
    rs(11) = DTPicker1.value
    rs.Update
    MsgBox "Transaction Complete!", vbInformation, Me.Caption
    sql = "SELECT iwadco_user.id AS ID, iwadco_user.username AS Username, iwadco_actype.account_type AS 'Account Type' FROM iwadco_user INNER JOIN iwadco_actype ON iwadco_user.type = iwadco_actype.id AND iwadco_user.status='E'"
    lstview.lstDatabase sql, frmSysUser.ListView1, 2
    If Me.Caption <> "Edit User Account" Then
        If MsgBox("Do you want to add another user?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            clear_it
        Else
            Unload Me
            If frmMain.Visible = False Then
                frmLogin.Show
            End If
        End If
    Else
        Unload Me
    End If
End If
End Sub

Private Sub Form_Activate()
If Me.Caption = "Edit User Account" Then
    Call editdaw
End If
End Sub

Private Sub Form_Load()
sql = "SELECT account_type FROM iwadco_actype"
Call LoadCombo(sql, Combo1)
cmdUpdate.CaptionAlign = isbright
cmdUpdate.IconAlign = isbLeft
cmdcancel.Caption = isbright
cmdcancel.IconAlign = isbLeft
cmdcancel.Caption = "Close"
End Sub

Sub clear_it()
txtUserName.Text = ""
txtPassword.Text = ""
txtConfirm.Text = ""
End Sub

Sub editdaw()
sql = "SELECT iwadco_actype.account_type AS account_type,* FROM iwadco_user INNER JOIN iwadco_actype ON iwadco_user.type = iwadco_actype.id WHERE (iwadco_user.id = " & frmSysUser.ListView1.SelectedItem.Text & ")"
rsCheck
rs.Open sql, CN, adOpenStatic, adLockOptimistic
txtUserName.Text = rs("username")
txtLname.Text = rs("lname")
txtfname.Text = rs("fname")
txtmname.Text = rs("mname")
txtAddress.Text = rs("address")
txtPhone.Text = rs("phone")
txtMobile.Text = rs("mobile")
txtemail.Text = rs("email")
Combo1.SetFocus
Call selectCombo(rs("type"), Combo1)
DTPicker1.value = rs("datehired")
End Sub

