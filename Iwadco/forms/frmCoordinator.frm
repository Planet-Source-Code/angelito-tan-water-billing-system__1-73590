VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMCoordinator 
   Caption         =   "Coordinator"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   Icon            =   "frmCoordinator.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   10050
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCoordinator.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCoordinator.frx":0B24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstCoor 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "List of coordinator records were in you can add or edit coordinator information"
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
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmCoordinator.frx":6CAE
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coordinators Lists"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   525
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3990
   End
End
Attribute VB_Name = "frmCoordinator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
UnloadAllExceptOne (Me.Name)
lstview.lstDatabase "SELECT id as 'Account No',lname+', '+fname+' '+mname as Name,address as 'Address',tel as 'Phone',mobile as 'Mobile',email as 'Email',datehired as 'Date Registered',billingdate as 'Billing Date' FROM iwadco_coor WHERE status='E' ORDER BY id ASC", Lstcoor, 2
frmMain.Picture2.Visible = True
frmMain.Picture3.Visible = True
End Sub

Private Sub Form_Resize()
Lstcoor.Width = Me.Width - 300
Lstcoor.Height = Me.Height - frmMain.Picture3.Height - frmMain.StatusBar1.Height - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Picture2.Visible = False
frmMain.Picture3.Visible = False
End Sub

Private Sub lstCoor_DblClick()
Call edit
End Sub

Private Sub lstCoor_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
    PopupMenu frmMain.mnuNEW
End If
End Sub

Public Sub ADD()
If user_priv("add") = False Then MsgBox "You are not allowed to add transaction!", vbExclamation, Me.Caption: Exit Sub
FormShow frmAddCoor, False
End Sub

Public Sub edit()
If user_priv("update") = False Then MsgBox "You are not allowed to add record!", vbExclamation, Me.Caption: Exit Sub
FormShow frmAddCoor, True
End Sub

Public Sub Ref()
lstview.lstDatabase "SELECT id as 'Account No',lname+', '+fname+' '+mname as Name,address as 'Address',tel as 'Phone',mobile as 'Mobile',email as 'Email',datehired as 'Date Registered' FROM iwadco_coor WHERE status='E' ORDER BY id ASC", Lstcoor, 2
End Sub

Public Sub del()
'check if the user has the priviledge to delete record
If user_priv("delete") = False Then MsgBox "You are not allowed to delete record!", vbExclamation, Me.Caption: Exit Sub

If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
    rsCheck
    
    'set the record status as Disable
    rs.Open "UPDATE iwadco_coor SET status='D' WHERE id=" & Lstcoor.SelectedItem.Text
    MsgBox "Record has been deleted!", vbInformation, Me.Caption
    
    'refresh listview in frmSysUser
    sql = "SELECT iwadco_user.id AS ID, iwadco_user.username AS Username, iwadco_actype.account_type AS 'Account Type' FROM iwadco_user INNER JOIN iwadco_actype ON iwadco_user.type = iwadco_actype.id AND iwadco_user.status='E'"
    lstview.lstDatabase "SELECT id as 'Account No',lname+', '+fname+' '+mname as Name,address as 'Address',tel as 'Phone',mobile as 'Mobile',email as 'Email',datehired as 'Date Registered' FROM iwadco_coor WHERE status='E' ORDER BY id ASC", Lstcoor, 2
End If
End Sub

Public Sub clse()
Unload Me
Set frmCoordinator = Nothing
End Sub
