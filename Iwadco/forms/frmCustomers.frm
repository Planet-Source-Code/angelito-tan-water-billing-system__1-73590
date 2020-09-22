VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmconsumer 
   Caption         =   "Consumers Profile"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   13455
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   2775
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   240
      Top             =   5880
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
            Picture         =   "frmCustomers.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomers.frx":08DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomers.frx":0E74
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstCustomer 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList2"
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Select Coordinator:"
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
      Top             =   1080
      Width           =   1890
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Select Area:"
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
      Left            =   5520
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "List of consumer records were in you can add or edit consumer information"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   360
      Picture         =   "frmCustomers.frx":76D6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consumers List"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   200
      Width           =   3375
   End
End
Attribute VB_Name = "frmConsumer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim shws As Boolean
Dim X As Integer
Dim coor_ As Integer
Dim area_ As Integer

Private Sub cmddelete_Click()
Call del
End Sub

Private Sub cmdedit_Click()
If user_priv("update") = False Then MsgBox "You are not allowed to edit record!", vbExclamation, Me.Caption: Exit Sub
FormShow frmAddConsumer, True
End Sub

Private Sub cmdNew_Click()
If user_priv("add") = False Then MsgBox "You are not allowed to add record!", vbExclamation, Me.Caption: Exit Sub
frmAddConsumer.Show 1
End Sub

Private Sub cmdRefresh_Click()
sql = "SELECT iwadco_cons.id as 'Account No',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name',iwadco_cons.address as 'Address',iwadco_cons.tel as 'Phone',iwadco_cons.mobile as 'Mobile',iwadco_cons.email as 'Email',iwadco_cons.dateregistered as 'Date Registered',iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname as 'Coordinator Name' from iwadco_cons,iwadco_coor WHERE iwadco_cons.status = 'E' AND  iwadco_coor.id = iwadco_cons.coor_id ORDER BY iwadco_cons.id ASC"
lstview.lstDatabase sql, lstCustomer, 1
End Sub

Private Sub Combo1_Click()
sql = "SELECT id FROM iwadco_area WHERE area = '" & Combo1.Text & "'"
rsCheck
rs.Open sql, CN, adOpenStatic, adLockOptimistic
area_ = rs(0)
lstview.lstDatabase "SELECT iwadco_cons.id as 'Account No',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name',iwadco_cons.address as 'Address',iwadco_cons.tel as 'Phone',iwadco_cons.mobile as 'Mobile',iwadco_cons.email as 'Email',iwadco_cons.dateregistered as 'Date Registered',iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname as 'Coordinator Name' from iwadco_cons,iwadco_coor WHERE iwadco_cons.status = 'E' AND iwadco_coor.id = iwadco_cons.coor_id AND coor_id = " & coor_ & " AND area_id = " & area_ & " ORDER BY iwadco_cons.id ASC", lstCustomer, 2
End Sub

Private Sub Combo3_Click()
Dim X As Integer
Combo1.Clear
sql = "SELECT id FROM iwadco_coor WHERE lname+', '+fname+' '+mname = '" & Combo3.Text & "'"
rsCheck
rs.Open sql, CN, adOpenStatic, adLockOptimistic
coor_ = rs(0)
sql = "SELECT area FROM iwadco_area WHERE coor_id=" & coor_
rsCheck
rs.Open sql, CN, adOpenStatic, adLockOptimistic
If rs.RecordCount <> 0 Then
    For X = 1 To rs.RecordCount
        Combo1.AddItem rs(0)
        rs.MoveNext
    Next
End If
lstview.lstDatabase "SELECT iwadco_cons.id as 'Account No',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name',iwadco_cons.address as 'Address',iwadco_cons.tel as 'Phone',iwadco_cons.mobile as 'Mobile',iwadco_cons.email as 'Email',iwadco_cons.dateregistered as 'Date Registered',iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname as 'Coordinator Name' from iwadco_cons,iwadco_coor WHERE iwadco_cons.status = 'E' AND iwadco_coor.id = iwadco_cons.coor_id AND coor_id = " & coor_ & " ORDER BY iwadco_cons.id ASC", lstCustomer, 2
End Sub

Private Sub Form_Activate()
UnloadAllExceptOne (Me.Name)
lstview.lstDatabase "SELECT iwadco_cons.id as 'Account No',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name',iwadco_cons.address as 'Address',iwadco_cons.tel as 'Phone',iwadco_cons.mobile as 'Mobile',iwadco_cons.email as 'Email',iwadco_cons.dateregistered as 'Date Registered',iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname as 'Coordinator Name' from iwadco_cons,iwadco_coor WHERE iwadco_cons.status = 'E' AND iwadco_coor.id = iwadco_cons.coor_id ORDER BY iwadco_cons.id ASC", lstCustomer, 2
frmMain.Picture2.Visible = True
frmMain.Picture3.Visible = True
End Sub

Private Sub Form_Load()
Dim X As Integer
Combo3.Clear
sql = "SELECT lname+', '+fname+' '+mname as coor_name FROM iwadco_coor"
rsCheck
rs.Open sql, CN, adOpenStatic, adLockOptimistic
If rs.RecordCount <> 0 Then
    For X = 1 To rs.RecordCount
        Combo3.AddItem rs(0)
        rs.MoveNext
    Next
End If
End Sub

Private Sub Form_Resize()
X = X + 1
    If X = 1 Then
    lstCustomer.Width = Me.Width - 300
    lstCustomer.Height = Me.Height - frmMain.Picture3.Height - frmMain.StatusBar1.Height - 750
    Else
    lstCustomer.Width = Me.Width - 300
    lstCustomer.Height = Me.Height - frmMain.Picture3.Height - frmMain.StatusBar1.Height - 750
    Me.WindowState = 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Picture2.Visible = False
frmMain.Picture3.Visible = False
End Sub

Private Sub lstCustomer_DblClick()
Call edit
End Sub

Private Sub lstCustomer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CONSUMERID = ""
End Sub

Private Sub lstCustomer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu frmMain.mnuNEW
End Sub

Public Sub Ref()
sql = "SELECT iwadco_cons.id as 'Account No',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name',iwadco_cons.address as 'Address',iwadco_cons.tel as 'Phone',iwadco_cons.mobile as 'Mobile',iwadco_cons.email as 'Email',iwadco_cons.dateregistered as 'Date Registered',iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname as 'Coordinator Name' from iwadco_cons,iwadco_coor WHERE iwadco_cons.status = 'E' AND iwadco_coor.id = iwadco_cons.coor_id ORDER BY iwadco_cons.id ASC"
lstview.lstDatabase sql, lstCustomer, 1
End Sub

Public Sub ADD()
If user_priv("add") = False Then MsgBox "You are not allowed to add transaction!", vbExclamation, Me.Caption: Exit Sub
FormShow frmAddConsumer, False
End Sub

Public Sub edit()
If lstCustomer.ListItems.Count > 0 Then
If user_priv("update") = False Then MsgBox "You are not allowed to edit record!", vbExclamation, Me.Caption: Exit Sub
FormShow frmAddConsumer, True
End If
End Sub

Public Sub clse()
Unload Me
Set frmConsumer = Nothing
End Sub

Public Sub del()
On Error Resume Next
'check if the user has the priviledge to delete record
If user_priv("delete") = False Then MsgBox "You are not allowed to delete record!", vbExclamation, Me.Caption: Exit Sub
If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
    rsCheck
    'set the record status as Disable
    rs.Open "UPDATE iwadco_cons SET status='D' WHERE id='" & lstCustomer.SelectedItem.Text & "'"
    MsgBox "Record has been deleted!", vbInformation, Me.Caption
    lstview.lstDatabase "SELECT iwadco_cons.id as 'Account No',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name',iwadco_cons.address as 'Address',iwadco_cons.tel as 'Phone',iwadco_cons.mobile as 'Mobile',iwadco_cons.email as 'Email',iwadco_cons.dateregistered as 'Date Registered',iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname as 'Coordinator Name' from iwadco_cons,iwadco_coor WHERE iwadco_cons.status = 'E' AND iwadco_coor.id = iwadco_cons.coor_id ORDER BY iwadco_cons.id ASC", lstCustomer, 2
    'REFRESH ITEM IN FRMMAIN
    rsCheck
    rs.Open
    lstview.lstDatabase "SELECT iwadco_cons.id as 'Account No',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name',iwadco_cons.address as 'Address',iwadco_cons.tel as 'Phone',iwadco_cons.mobile as 'Mobile',iwadco_cons.email as 'Email',iwadco_cons.dateregistered as 'Date Registered',iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname as 'Coordinator Name' from iwadco_cons,iwadco_coor WHERE iwadco_cons.status = 'E' AND iwadco_coor.id = iwadco_cons.coor_id ORDER BY iwadco_cons.id ASC", frmMain.lstemp, 2
    
End If
End Sub
