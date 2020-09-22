VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRMPAYMENTS 
   Caption         =   "Payments Form"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   120
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   4680
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
            Picture         =   "frmpayments.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpayments.frx":059A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
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
      Height          =   855
      Left            =   840
      TabIndex        =   14
      Top             =   1440
      Width           =   10215
      Begin VB.TextBox txtacctno 
         BackColor       =   &H80000018&
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
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin Project1.isButton cmdsearch 
         Height          =   315
         Left            =   3480
         TabIndex        =   1
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         Icon            =   "frmpayments.frx":6724
         Style           =   5
         Caption         =   "&Search"
         IconAlign       =   1
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Consumer Account Number"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   21
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         ForeColor       =   &H00BB5900&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         ForeColor       =   &H00BB5900&
         Height          =   195
         Left            =   5400
         TabIndex        =   16
         Top             =   315
         Width           =   450
      End
      Begin VB.Label lbldte 
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   6000
         TabIndex        =   15
         Top             =   315
         Width           =   3540
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Payment Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   840
      TabIndex        =   13
      Top             =   4440
      Width           =   10215
      Begin MSComctlLib.ListView lstpayments 
         Height          =   3015
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483634
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
   End
   Begin Project1.isButton cmdclose 
      Height          =   420
      Left            =   2400
      TabIndex        =   7
      Top             =   7920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmpayments.frx":6740
      Style           =   5
      Caption         =   "&Close"
      IconSize        =   20
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
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
   Begin Project1.isButton cmdOk 
      Height          =   420
      Left            =   840
      TabIndex        =   6
      Top             =   7920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmpayments.frx":1D7DA
      Style           =   5
      Caption         =   "&Ok"
      IconSize        =   20
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Account Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   840
      TabIndex        =   8
      Top             =   2400
      Width           =   10215
      Begin VB.TextBox txtmobile 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   770
         Width           =   3375
      End
      Begin VB.TextBox txttel 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtaddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   770
         Width           =   3855
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         ForeColor       =   &H00BB5900&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   405
         Width           =   510
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address :"
         ForeColor       =   &H00BB5900&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone # :"
         ForeColor       =   &H00BB5900&
         Height          =   195
         Left            =   5460
         TabIndex        =   10
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile # :"
         ForeColor       =   &H00BB5900&
         Height          =   195
         Left            =   5520
         TabIndex        =   9
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "NOTICE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   25
      Top             =   7920
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "This account has been under a promisorry note."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "NOTICE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   23
      Top             =   7920
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Label Label10 
         Caption         =   "This consumer account has been disconnected ."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   675
      TabIndex        =   19
      Top             =   3795
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   600
      Picture         =   "frmpayments.frx":34874
      Top             =   3960
      Width           =   195
   End
   Begin VB.Image imag1 
      Height          =   555
      Index           =   0
      Left            =   600
      MousePointer    =   7  'Size N S
      Picture         =   "frmpayments.frx":34E7E
      Top             =   4440
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can view Customer's Payments Details and Transact Payment."
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
      Height          =   210
      Index           =   1
      Left            =   1800
      TabIndex        =   18
      Top             =   960
      Width           =   5475
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   1
      Left            =   840
      Picture         =   "frmpayments.frx":35488
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   870
      Left            =   1680
      TabIndex        =   17
      Top             =   120
      Width           =   3600
   End
End
Attribute VB_Name = "frmpayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdclose_Click()
Unload Me
Set frmpayments = Nothing
End Sub

Private Sub cmdok_Click()
    If lstpayments.ListItems.Count > 0 Then
    frmpay.Show 1
    End If
End Sub

Private Sub cmdSearch_Click()
rsCheck
sql = "SELECT * FROM iwadco_cons WHERE id = '" & txtacctno.Text & "' AND status<>'I' AND tappingStatus='I'"
Debug.Print sql
rs.Open sql, CN, adOpenStatic, adLockOptimistic

If rs.RecordCount = 1 Then
        If MsgBox("Tapping fee is Incomplete!" & Chr(10) & "Complete the payment?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            frmTappingFee.Show 1
        End If
End If

'---If disconnected
rsCheck
rs.Open "SELECT * FROM iwadco_cons WHERE id = '" & txtacctno.Text & "'AND status ='X'", CN, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
    Frame4.Visible = True
Else
    Frame4.Visible = False
End If
'--
rsCheck
rs.Open "SELECT lname,mname,fname,tel,mobile,address,iwadco_readings.account_no,iwadco_readings.status FROM iwadco_cons,iwadco_readings WHERE iwadco_cons.id ='" & txtacctno.Text & "'AND iwadco_readings.account_no = iwadco_cons.id"
'find consumer account
If rs.RecordCount > 0 Then
    txtname.Text = rs("lname") & ", " & rs("fname") & ", " & rs("mname")
    txttel.Text = rs("tel")
    txtAddress.Text = rs("address")
    txtMobile.Text = rs("mobile")
    
    '=1 list all the incomplete payments if there's
    sql = "SELECT id as 'ID Number',account_no as 'Account Number',billto as 'Billing Date',due_date as 'Due Date',total_amount as 'Total Amount',amountpaid as 'Amount Paid', promisorry as 'Promisorry Note' FROM iwadco_readings WHERE account_no = '" & txtacctno.Text & "' AND deletedby=0 AND status = 'I'" ' and due_date >='" & Format(Now, "mm/dd/yyyy") & "'"
    lstview.lstDatabase sql, lstpayments, 2
Else
    '=0 then exit sub
    MsgBox "Account number has not been read!", vbExclamation, Me.Caption
    txtname.Text = ""
    txttel.Text = ""
    txtAddress.Text = ""
    txtMobile.Text = ""
    lstpayments.ListItems.Clear
    Exit Sub
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
'popupmenu
If CONSUMERID <> "" Then
    txtacctno.Text = CONSUMERID
    cmdSearch_Click
    txtacctno.SetFocus
    SendKeys "{end}"
End If
UnloadAllExceptOne (Me.Name)
formBoolean = True
End Sub

Private Sub Form_Load()
lbldte.Caption = Format(Date, "dddd mmmm dd,yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
CONSUMERID = ""
formBoolean = False
End Sub

Private Sub lstpayments_Click()
If lstpayments.ListItems.Count > 0 Then
    If lstpayments.SelectedItem.SubItems(6) = "Yes       " Then
        Frame5.Visible = True
        Else
        Frame5.Visible = False
    End If
End If
End Sub

Private Sub lstpayments_DblClick()
Call cmdok_Click
End Sub

Private Sub Timer1_Timer()
If CONSUMERID <> "" Then
    txtacctno.Text = CONSUMERID
    Else
    Exit Sub
End If
End Sub

Private Sub txtacctno_Change()
txtacctno.Text = str_Filter(txtacctno, 48, 57, 45)
'SET
CONSUMERID = txtacctno.Text
End Sub

Private Sub txtacctno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Call cmdSearch_Click
End If
End Sub

Public Sub dblclick()
cmdSearch_Click
End Sub
