VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRMaddcoor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Coordinators Settings"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9225
   Icon            =   "frmAddCoordinator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   9015
      Begin VB.TextBox txtemail 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   6030
         TabIndex        =   8
         Top             =   1560
         Width           =   2685
      End
      Begin VB.ComboBox cmbBillingDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmAddCoordinator.frx":617A
         Left            =   6050
         List            =   "frmAddCoordinator.frx":61DB
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1950
         Width           =   855
      End
      Begin VB.TextBox txtmobile 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6030
         TabIndex        =   7
         Top             =   1150
         Width           =   2655
      End
      Begin VB.TextBox txttel 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6030
         TabIndex        =   6
         Top             =   770
         Width           =   2655
      End
      Begin VB.TextBox txtlname 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Top             =   770
         Width           =   2775
      End
      Begin VB.TextBox txtadd 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1950
         Width           =   2805
      End
      Begin VB.TextBox txtmname 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   3
         Top             =   1550
         Width           =   2775
      End
      Begin VB.TextBox txtfname 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Top             =   1150
         Width           =   2775
      End
      Begin VB.TextBox txtacctno 
         BackColor       =   &H00E6FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   1665
      End
      Begin Project1.isButton cmdsave 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "frmAddCoordinator.frx":6252
         Style           =   5
         Caption         =   "&Save"
         IconSize        =   18
         IconAlign       =   1
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.isButton cmdcancel 
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   2880
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "frmAddCoordinator.frx":1B3C4
         Style           =   5
         Caption         =   "&Cancel"
         IconSize        =   18
         IconAlign       =   1
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   6030
         TabIndex        =   5
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   47579137
         CurrentDate     =   39377
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   140
         TabIndex        =   22
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tel # :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4920
         TabIndex        =   20
         Top             =   880
         Width           =   570
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile # :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4920
         TabIndex        =   19
         Top             =   1250
         Width           =   825
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4920
         TabIndex        =   18
         Top             =   1590
         Width           =   585
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Hired :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4935
         TabIndex        =   17
         Top             =   480
         Width           =   990
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Account No :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Billing Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4920
         TabIndex        =   15
         Top             =   2000
         Width           =   945
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Add or Edit Coordinator Information"
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
      TabIndex        =   13
      Top             =   480
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmAddCoordinator.frx":3245E
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Coordinator Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   960
      TabIndex        =   12
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmAddCoor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecordEDIT As Boolean
Dim ActivateAcount As Long

Private Sub cmdcancel_Click()
Unload Me
Set frmAddCoor = Nothing
End Sub

Private Sub cmdsave_Click()
Dim errMsg As String, noErr As String
errMsg = "Please complete the following requirements!" & Chr(10) & "-----------------------------"
noErr = errMsg
If txtLname.Text = "" Then
    errMsg = errMsg & Chr(10) & "Lastname"
End If
If txtfname.Text = "" Then
    errMsg = errMsg & Chr(10) & "Firstname"
End If
If txtmname.Text = "" Then
    errMsg = errMsg & Chr(10) & "Middlename"
End If
If txtadd.Text = "" Then
    errMsg = errMsg & Chr(10) & "Address"
End If
If cmbBillingDate.Text = "" Then
    errMsg = errMsg & Chr(10) & "Billing Date Setting"
End If
If errMsg <> noErr Then
    MsgBox errMsg, vbExclamation, Me.Caption
    Exit Sub
End If
If RecordEDIT = False Then
    rsCheck
    rs.Open "SELECT * FROM iwadco_coor"
    rs.AddNew
    rs(11) = Format(Now, "mm/dd/yyyy hh:mm:ss")
    rs(12) = empID
Else
    rsCheck
    rs.Open "SELECT * FROM iwadco_coor WHERE id =" & txtacctno.Text
    rs(9) = Format(Now, "mm/dd/yyyy hh:mm:ss")
    rs(10) = empID
End If
rs(1) = txtLname.Text
rs(2) = txtfname.Text
rs(3) = txtmname.Text
rs(4) = txtadd.Text
rs(5) = txttel.Text
rs(6) = txtMobile.Text
rs(7) = txtemail.Text
rs(8) = DTPicker1.value
rs(14) = cmbBillingDate.Text
rs.Update
If RecordEDIT = False Then
    MsgBox "New record has succesfully been saved", vbInformation, Me.Caption
    If MsgBox("Do you want to add new record", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        txtacctno.Text = rs.RecordCount + 1
        txtLname.Text = ""
        txtfname.Text = ""
        txtmname.Text = ""
        txtadd.Text = ""
        txttel.Text = ""
        txtMobile.Text = ""
        txtemail.Text = ""
    Else
        Unload Me
        Set frmAddCoor = Nothing
    End If
Else
    MsgBox "Record has been succesfully updated !", vbInformation, Me.Caption
    Unload Me
End If
lstview.lstDatabase "SELECT id as 'Account No',lname+', '+fname+' '+mname as Name,address as 'Address',tel as 'Phone',mobile as 'Mobile',email as 'Email',datehired as 'Date Registered' FROM iwadco_coor WHERE status='E' ORDER BY id ASC", frmCoordinator.Lstcoor, 1
End Sub

Private Sub Form_Activate()
If Right(Me.Caption, 3) = "Add" Then
    txtacctno.Text = ""
    txtLname.Text = ""
    txtfname.Text = ""
    txtmname.Text = ""
    txtadd.Text = ""
    txttel.Text = ""
    txtMobile.Text = ""
    txtemail.Text = ""
    rsCheck
    rs.Open "SELECT * FROM iwadco_coor", CN, adOpenStatic, adLockOptimistic
    If rs.RecordCount >= 1 Then
        txtacctno.Text = rs.RecordCount + 1
    Else
        txtacctno.Text = 1
    End If
    RecordEDIT = False
Else
    sql = "SELECT * FROM iwadco_coor WHERE id = " & frmCoordinator.Lstcoor.SelectedItem.Text
    rsCheck
    rs.Open sql
    txtacctno.Text = rs(0)
    txtLname.Text = rs(1)
    txtfname.Text = rs(2)
    txtmname.Text = rs(3)
    txtadd.Text = rs(4)
    txttel.Text = rs(5)
    txtMobile.Text = rs(6)
    txtemail.Text = rs(7)
    DTPicker1.value = rs(8)
    cmbBillingDate.Text = rs(14)
    Call selectCombo(rs(14), cmbBillingDate)
    RecordEDIT = True
    txtLname.SetFocus
End If
End Sub
