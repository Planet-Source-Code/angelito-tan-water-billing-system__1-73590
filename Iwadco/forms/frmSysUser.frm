VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRMSYSUSER 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System User Profile Form"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   Icon            =   "frmSysUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysUser.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysUser.frx":6DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysUser.frx":D64E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select a user below"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   80
      TabIndex        =   2
      Top             =   960
      Width           =   7440
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2655
      End
      Begin Project1.isButton cmdClose 
         Height          =   735
         Left            =   5880
         TabIndex        =   3
         Top             =   4800
         Width           =   1335
         _extentx        =   2355
         _extenty        =   1296
         icon            =   "frmSysUser.frx":137D8
         style           =   6
         caption         =   "&Close"
         iconsize        =   26
         captionalign    =   4
         iconalign       =   1
         inonthemestyle  =   0
         tooltiptitle    =   ""
         tooltipicon     =   0
         tooltiptype     =   0
         ttforecolor     =   0
         font            =   "frmSysUser.frx":2A874
      End
      Begin Project1.isButton cmdDelete 
         Height          =   735
         Left            =   1560
         TabIndex        =   4
         Top             =   4800
         Width           =   1335
         _extentx        =   2355
         _extenty        =   1296
         icon            =   "frmSysUser.frx":2A89C
         style           =   6
         caption         =   "&Remove"
         iconsize        =   26
         captionalign    =   4
         iconalign       =   1
         inonthemestyle  =   0
         tooltiptitle    =   ""
         tooltipicon     =   0
         tooltiptype     =   0
         ttforecolor     =   0
         font            =   "frmSysUser.frx":2B178
      End
      Begin Project1.isButton cmdEdit 
         Height          =   735
         Left            =   3000
         TabIndex        =   5
         Top             =   4800
         Width           =   1335
         _extentx        =   2355
         _extenty        =   1296
         icon            =   "frmSysUser.frx":2B1A0
         style           =   6
         caption         =   "&Edit"
         iconsize        =   26
         captionalign    =   4
         iconalign       =   1
         inonthemestyle  =   0
         tooltiptitle    =   ""
         tooltipicon     =   0
         tooltiptype     =   0
         ttforecolor     =   0
         font            =   "frmSysUser.frx":3132C
      End
      Begin Project1.isButton cmdAdd 
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   4800
         Width           =   1335
         _extentx        =   2355
         _extenty        =   1508
         icon            =   "frmSysUser.frx":31354
         style           =   6
         caption         =   "&Add"
         iconsize        =   26
         captionalign    =   4
         iconalign       =   1
         inonthemestyle  =   0
         tooltiptitle    =   ""
         tooltipicon     =   0
         tooltiptype     =   0
         ttforecolor     =   0
         font            =   "frmSysUser.frx":31C30
      End
      Begin Project1.isButton cmdSearch 
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         icon            =   "frmSysUser.frx":31C58
         style           =   5
         caption         =   "&Search"
         iconsize        =   20
         captionalign    =   2
         iconalign       =   1
         inonthemestyle  =   0
         tooltiptitle    =   ""
         tooltipicon     =   0
         tooltiptype     =   0
         ttforecolor     =   0
         font            =   "frmSysUser.frx":32EDC
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3855
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin Project1.isButton cmdRefresh 
         Height          =   735
         Left            =   4440
         TabIndex        =   9
         Top             =   4800
         Width           =   1335
         _extentx        =   2355
         _extenty        =   1296
         icon            =   "frmSysUser.frx":32F04
         style           =   6
         caption         =   "&Refresh"
         iconsize        =   26
         captionalign    =   4
         iconalign       =   1
         inonthemestyle  =   4
         highlightcolor  =   -2147483629
         tooltiptitle    =   ""
         tooltipicon     =   0
         tooltiptype     =   0
         ttforecolor     =   0
         font            =   "frmSysUser.frx":39090
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmSysUser.frx":390B8
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Add, Edit , Delete Existing User Profile Account"
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
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "System User Profiles"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmSysUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
'check if user level is administrator
If empType > 1 Then MsgBox "You are not authorized to Add record!", vbExclamation, Me.Caption: Exit Sub
frmAddUser.Caption = "Add User Account"
frmAddUser.Show 1
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
'check if user level is administrator
If empType > 1 Then MsgBox "You are not authorized to delte record!", vbExclamation, Me.Caption: Exit Sub

'check if the record to be deleted is the active user
If ListView1.SelectedItem.Text = empID Then MsgBox "You cannot delete an active acount!", vbExclamation, Me.Caption: Exit Sub

If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
    rsCheck
    
    'set the record status as Disable
    rs.Open "UPDATE iwadco_user SET status='D' WHERE id=" & ListView1.SelectedItem.Text
    MsgBox "Record has been deleted!", vbInformation, Me.Caption
    
    'refresh listview in frmSysUser
    sql = "SELECT iwadco_user.id AS ID, iwadco_user.username AS Username, iwadco_actype.account_type AS 'Account Type' FROM iwadco_user INNER JOIN iwadco_actype ON iwadco_user.type = iwadco_actype.id AND iwadco_user.status='E'"
    lstview.lstDatabase sql, ListView1, 2
End If
End Sub

Private Sub cmdedit_Click()
'check if user level is administrator
If empType > 1 And ListView1.SelectedItem.Text <> empID Then MsgBox "You are not authorized to modify other accounts!", vbExclamation, Me.Caption: Exit Sub
If ListView1.SelectedItem.Text <> empID Then MsgBox "You cannot edit other accounts!", vbExclamation, Me.Caption: Exit Sub
If ListView1.SelectedItem.Text = "" Then MsgBox "Please select record to modify!": Exit Sub
Load frmAddUser
frmAddUser.Caption = "Edit User Account"
frmAddUser.Show 1
End Sub

Private Sub cmdRefresh_Click()
sql = "SELECT iwadco_user.id AS ID, iwadco_user.username AS Username, iwadco_actype.account_type AS 'Account Type' FROM iwadco_user INNER JOIN iwadco_actype ON iwadco_user.type = iwadco_actype.id AND iwadco_user.status='E'"
lstview.lstDatabase sql, ListView1, 2
End Sub

Private Sub cmdSearch_Click()
sql = "SELECT iwadco_user.id AS ID, iwadco_user.username AS Username, iwadco_actype.account_type AS 'Account Level' FROM iwadco_user INNER JOIN iwadco_actype ON iwadco_user.type = iwadco_actype.id Where (username LIKE '" & txtSearch.Text & "%')"
rsCheck
lstview.lstDatabase sql, ListView1, 1
End Sub

Private Sub Form_Load()
If empID = Null Then
    cmdedit.Enabled = False
    cmdDelete.Enabled = False
End If
sql = "SELECT iwadco_user.id AS ID, iwadco_user.username AS Username, iwadco_actype.account_type AS 'Account Type' FROM iwadco_user INNER JOIN iwadco_actype ON iwadco_user.type = iwadco_actype.id AND iwadco_user.status='E'"
lstview.lstDatabase sql, ListView1, 1
cmdAdd.IconAlign = isbTop
cmdAdd.CaptionAlign = isbBottom
cmdedit.IconAlign = isbTop
cmdedit.CaptionAlign = isbBottom
cmdDelete.IconAlign = isbTop
cmdDelete.CaptionAlign = isbBottom
cmdclose.IconAlign = isbTop
cmdclose.CaptionAlign = isbBottom
cmdRefresh.IconAlign = isbTop
cmdRefresh.CaptionAlign = isbBottom
'search buttong
cmdSearch.CaptionAlign = isbright
End Sub

Private Sub Form_Unload(Cancel As Integer)
If empID = Null Then
    frmLogin.Show 1
End If
End Sub

Private Sub ListView1_DblClick()
cmdedit_Click
End Sub

