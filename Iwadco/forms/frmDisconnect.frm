VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDisconnect 
   Caption         =   "Disconnection"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12465
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   12465
   WindowState     =   2  'Maximized
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
      TabIndex        =   11
      Top             =   1080
      Width           =   2775
   End
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
      TabIndex        =   9
      Top             =   1080
      Width           =   2775
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
      Left            =   6840
      TabIndex        =   7
      Top             =   7800
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
         TabIndex        =   8
         Top             =   240
         Width           =   5415
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
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   12255
      Begin MSComctlLib.ListView lstpayments 
         Height          =   5775
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   7320
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   4560
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
            Picture         =   "frmDisconnect.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisconnect.frx":059A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Project1.isButton cmdclose 
      Height          =   420
      Left            =   2400
      TabIndex        =   2
      Top             =   7800
      Width           =   1335
      _extentx        =   2355
      _extenty        =   741
      icon            =   "frmDisconnect.frx":6724
      style           =   5
      caption         =   "&Close"
      iconsize        =   20
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frmDisconnect.frx":1D7C0
   End
   Begin Project1.isButton cmdOk 
      Height          =   420
      Left            =   240
      TabIndex        =   3
      Top             =   7800
      Width           =   2055
      _extentx        =   2355
      _extenty        =   741
      icon            =   "frmDisconnect.frx":1D7E8
      style           =   5
      caption         =   "&Disconnect"
      iconsize        =   20
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frmDisconnect.frx":34884
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
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
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
      TabIndex        =   10
      Top             =   1200
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disconnection"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   0
      Width           =   5115
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   1
      Left            =   240
      Picture         =   "frmDisconnect.frx":348AC
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts to be disconnect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   720
      Width           =   2565
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   75
      TabIndex        =   4
      Top             =   3675
      Width           =   135
   End
End
Attribute VB_Name = "frmDisconnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstv As New DataViewer
Dim coor_ As Integer
Dim area_ As Integer

Private Sub cmdclose_Click()
Unload Me
Set frmDisconnect = Nothing
End Sub

Private Sub cmdok_Click()
Dim X, Y As Integer

If lstpayments.ListItems.Count > 0 Then
    For Y = 1 To lstpayments.ListItems.Count
    rsCheck
    rs.Open "SELECT * FROM iwadco_readings WHERE account_no = '" & frmDisconnect.lstpayments.ListItems(Y).Text & "'"
       If rs.RecordCount > 0 Then
         If lstpayments.ListItems(Y).Checked = True Then
            If rs("promisorry") = "Yes       " Then
            MsgBox "Account number " & rs("account_no") & " is under promisorry note", vbInformation, "Disconnection"
            Exit Sub
            End If
         End If
       End If
    Next
End If

If lstpayments.ListItems.Count > 0 Then
    If MsgBox("Are you sure you want to disconnect this consumer ?", vbExclamation + vbYesNo) = vbYes Then
        For X = 1 To lstpayments.ListItems.Count
            If lstpayments.ListItems(X).Checked = True Then
            rsCheck
            rs.Open "SELECT * FROM iwadco_cons WHERE id = '" & lstpayments.ListItems(X).Text & "'"
            rs("status") = "X"
            rs("datedisconnected") = Format(Now, "mm/dd/yyyy")
            rs.Update
            End If
        Next
    MsgBox "Cosumer disconnection succesful", vbInformation, "Disconnection"
    sql = "SELECT     iwadco_cons.id AS 'Account Number', iwadco_cons.lname + ', ' + iwadco_cons.fname AS 'Full Name', iwadco_readings.billfrom AS 'Bill From'," & _
                      "iwadco_readings.billto AS 'Bill To', iwadco_readings.due_date AS 'Due Date', iwadco_readings.consume AS Consume," & _
                      "iwadco_readings.excess AS Excess, iwadco_readings.total_amount - iwadco_typcon.min_rate AS 'Amount Excess', " & _
                      "iwadco_readings.arrears AS Arrears, iwadco_readings.amount_excess AS 'Present Billing',iwadco_readings.wtax as 'W/tax',iwadco_readings.total_amount as 'Amount Due'" & _
        "FROM         iwadco_cons INNER JOIN " & _
                      "iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id INNER JOIN " & _
                      "iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no " & _
        "WHERE DATEADD(DAY,5,iwadco_readings.due_date)<='" & Format(Now, "mm/dd/yyyy") & "' " & _
        "AND iwadco_readings.status='I' AND iwadco_cons.status ='E' AND iwadco_readings.amountpaid = 0"
        If Combo3.Text <> "" Then
            sql = sql & " AND coor_id=" & coor_
        End If
        If Combo1.Text <> "" Then
            sql = sql & " AND area_id=" & area_
        End If
        lstv.lstDatabase sql, lstpayments, 1
    End If
End If
End Sub

Private Sub Combo1_Click()
sql = "SELECT id FROM iwadco_area WHERE area = '" & Combo1.Text & "'"
rsCheck
rs.Open sql, CN, adOpenStatic, adLockOptimistic
area_ = rs(0)
sql = "SELECT     iwadco_cons.id AS 'Account Number', iwadco_cons.lname + ', ' + iwadco_cons.fname AS 'Full Name', iwadco_readings.billfrom AS 'Bill From'," & _
                      "iwadco_readings.billto AS 'Bill To', iwadco_readings.due_date AS 'Due Date', iwadco_readings.consume AS Consume," & _
                      "iwadco_readings.excess AS Excess, iwadco_readings.total_amount - iwadco_typcon.min_rate AS 'Amount Excess', " & _
                      "iwadco_readings.arrears AS Arrears, iwadco_readings.amount_excess AS 'Present Billing',iwadco_readings.wtax as 'W/tax',iwadco_readings.total_amount as 'Amount Due'" & _
        "FROM         iwadco_cons INNER JOIN " & _
                      "iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id INNER JOIN " & _
                      "iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no " & _
        "WHERE DATEADD(DAY,5,iwadco_readings.due_date)<='" & Format(Now, "mm/dd/yyyy") & "' " & _
        "AND iwadco_readings.status='I' AND iwadco_cons.status ='E' AND iwadco_readings.amountpaid =0 AND coor_id=" & coor_ & " AND area_id = " & area_ & ""
        Debug.Print sql
lstv.lstDatabase sql, lstpayments, 1
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
sql = "SELECT     iwadco_cons.id AS 'Account Number', iwadco_cons.lname + ', ' + iwadco_cons.fname AS 'Full Name', iwadco_readings.billfrom AS 'Bill From'," & _
                      "iwadco_readings.billto AS 'Bill To', iwadco_readings.due_date AS 'Due Date', iwadco_readings.consume AS Consume," & _
                      "iwadco_readings.excess AS Excess, iwadco_readings.total_amount - iwadco_typcon.min_rate AS 'Amount Excess', " & _
                      "iwadco_readings.arrears AS Arrears, iwadco_readings.amount_excess AS 'Present Billing',iwadco_readings.wtax as 'W/tax',iwadco_readings.total_amount as 'Amount Due'" & _
        "FROM         iwadco_cons INNER JOIN " & _
                      "iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id INNER JOIN " & _
                      "iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no " & _
        "WHERE DATEADD(DAY,5,iwadco_readings.due_date)<='" & Format(Now, "mm/dd/yyyy") & "' " & _
        "AND iwadco_readings.status='I' AND iwadco_cons.status ='E' AND iwadco_readings.amountpaid =0 AND coor_id=" & coor_
lstv.lstDatabase sql, lstpayments, 1
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
sql = "SELECT     iwadco_cons.id AS 'Account Number', iwadco_cons.lname + ', ' + iwadco_cons.fname AS 'Full Name', iwadco_readings.billfrom AS 'Bill From'," & _
                      "iwadco_readings.billto AS 'Bill To', iwadco_readings.due_date AS 'Due Date', iwadco_readings.consume AS Consume," & _
                      "iwadco_readings.excess AS Excess, iwadco_readings.total_amount - iwadco_typcon.min_rate AS 'Amount Excess', " & _
                      "iwadco_readings.arrears AS Arrears, iwadco_readings.amount_excess AS 'Present Billing',iwadco_readings.wtax as 'W/tax',iwadco_readings.total_amount as 'Amount Due'" & _
        "FROM         iwadco_cons INNER JOIN " & _
                      "iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id INNER JOIN " & _
                      "iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no " & _
        "WHERE DATEADD(DAY,5,iwadco_readings.due_date)<='" & Format(Now, "mm/dd/yyyy") & "' " & _
        "AND iwadco_readings.status='I' AND iwadco_cons.status ='E' AND iwadco_readings.amountpaid =0"
lstv.lstDatabase sql, lstpayments, 1
End Sub

Private Sub lstpayments_Click()
If lstpayments.ListItems.Count > 0 Then
    rsCheck
    rs.Open "SELECT * FROM iwadco_readings WHERE account_no ='" & lstpayments.SelectedItem.Text & "'"
    If rs("promisorry") = "Yes       " Then
        Frame5.Visible = True
        Else
        Frame5.Visible = False
    End If
End If
End Sub
