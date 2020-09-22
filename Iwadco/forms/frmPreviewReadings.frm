VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPreviewReadings 
   Caption         =   "Readings Preview"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   9720
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
      Top             =   5805
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
            Picture         =   "frmPreviewReadings.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewReadings.frx":08DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   6525
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
            Picture         =   "frmPreviewReadings.frx":0E74
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstCustomer 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6588
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
   Begin VB.Label Label3 
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
      Index           =   6
      Left            =   8040
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Index           =   5
      Left            =   6720
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Index           =   4
      Left            =   5400
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Index           =   3
      Left            =   4080
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Index           =   2
      Left            =   2760
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Index           =   1
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1440
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
      Index           =   0
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
   Begin VB.Image Image1 
      Height          =   585
      Left            =   360
      Picture         =   "frmPreviewReadings.frx":76D6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check Readings"
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
      TabIndex        =   2
      Top             =   120
      Width           =   3450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview of readings, this is where you check the readings of the consumers"
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
      TabIndex        =   1
      Top             =   645
      Width           =   7335
   End
End
Attribute VB_Name = "frmPreviewReadings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim coor_ As Integer
Dim area_ As Integer
Dim prevdfrom As Date
Dim prevdto As Date

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

prevdfrom = smonth & "/01/" & syear
prevdto = DateAdd("m", 1, smonth & "/01/" & syear) - 1
sql = "SELECT Account_no as 'Account Number',readingno,lname+', '+fname+' '+mname as Name,iwadco_typcon.type,previous_reading as 'Previous Reading',present_reading as 'Present Readings',consume as 'Consume' ,excess as 'Excess',amount_excess as 'Amount Excess',arrears as 'Arrears',total_amount as 'Total Amount',iwadco_readings.amountpaid as 'Amount Paid' FROM iwadco_readings INNER JOIN iwadco_cons ON iwadco_readings.account_no = iwadco_cons.id INNER JOIN iwadco_typcon ON iwadco_cons.class=iwadco_typcon.id WHERE billto BETWEEN '" & prevdfrom & "' AND '" & prevdto & "' AND deletedby=0  ORDER BY lname+', '+fname+' '+mname"
lstview.lstDatabase sql, lstCustomer, 1

End Sub

Private Sub Form_Resize()
lstCustomer.Width = Me.Width - 300
lstCustomer.Height = Me.Height - frmMain.Picture3.Height - frmMain.StatusBar1.Height - 1000
End Sub

Private Sub Combo1_Click()
sql = "SELECT id FROM iwadco_area WHERE area = '" & Combo1.Text & "'"
rsCheck
rs.Open sql, CN, adOpenStatic, adLockOptimistic
area_ = rs(0)
'sql = "SELECT     iwadco_cons.id AS 'Account Number', iwadco_cons.lname + ', ' + iwadco_cons.fname AS 'Full Name', iwadco_readings.billfrom AS 'Bill From'," & _
                      "iwadco_readings.billto AS 'Bill To', iwadco_readings.due_date AS 'Due Date', iwadco_readings.consume AS Consume," & _
                      "iwadco_readings.excess AS Excess, iwadco_readings.total_amount - iwadco_typcon.min_rate AS 'Amount Excess', " & _
                      "iwadco_readings.arrears AS Arrears, iwadco_readings.amount_excess AS 'Present Billing',iwadco_readings.wtax as 'W/tax',iwadco_readings.total_amount as 'Amount Due'" & _
        "FROM         iwadco_cons INNER JOIN " & _
                      "iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id INNER JOIN " & _
                      "iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no " & _
        "WHERE DATEADD(DAY,5,iwadco_readings.due_date)<='" & Format(Now, "mm/dd/yyyy") & "' " & _
        "AND iwadco_readings.status='I' AND iwadco_cons.status ='E' AND iwadco_readings.amountpaid =0 AND coor_id=" & coor_ & " AND area_id = " & area_ & ""
sql = "SELECT Account_no as 'Account Number',readingno,lname+', '+fname+' '+mname as Name,iwadco_typcon.type,previous_reading as 'Previous Reading',present_reading as 'Present Readings',consume as 'Consume' ,excess as 'Excess',amount_excess as 'Amount Excess',arrears as 'Arrears',total_amount as 'Total Amount',iwadco_readings.amountpaid as 'Amount Paid' FROM iwadco_readings INNER JOIN iwadco_cons ON iwadco_readings.account_no = iwadco_cons.id INNER JOIN iwadco_typcon ON iwadco_cons.class=iwadco_typcon.id WHERE billto BETWEEN '" & prevdfrom & "' AND '" & prevdto & "' AND deletedby=0  AND coor_id=" & coor_ & " AND area_id = " & area_ & "" & "ORDER BY lname+', '+fname+' '+mname "
lstview.lstDatabase sql, lstCustomer, 1
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
'sql = "SELECT     iwadco_cons.id AS 'Account Number', iwadco_cons.lname + ', ' + iwadco_cons.fname AS 'Full Name', iwadco_readings.billfrom AS 'Bill From'," & _
                      "iwadco_readings.billto AS 'Bill To', iwadco_readings.due_date AS 'Due Date', iwadco_readings.consume AS Consume," & _
                      "iwadco_readings.excess AS Excess, iwadco_readings.total_amount - iwadco_typcon.min_rate AS 'Amount Excess', " & _
                      "iwadco_readings.arrears AS Arrears, iwadco_readings.amount_excess AS 'Present Billing',iwadco_readings.wtax as 'W/tax',iwadco_readings.total_amount as 'Amount Due'" & _
        "FROM         iwadco_cons INNER JOIN " & _
                      "iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id INNER JOIN " & _
                      "iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no " & _
        "WHERE DATEADD(DAY,5,iwadco_readings.due_date)<='" & Format(Now, "mm/dd/yyyy") & "' " & _
        "AND iwadco_readings.status='I' AND iwadco_cons.status ='E' AND iwadco_readings.amountpaid =0 AND coor_id=" & coor_
sql = "SELECT Account_no as 'Account Number',readingno,lname+', '+fname+' '+mname as Name,iwadco_typcon.type,previous_reading as 'Previous Reading',present_reading as 'Present Readings',consume as 'Consume' ,excess as 'Excess',amount_excess as 'Amount Excess',arrears as 'Arrears',total_amount as 'Total Amount',iwadco_readings.amountpaid as 'Amount Paid' FROM iwadco_readings INNER JOIN iwadco_cons ON iwadco_readings.account_no = iwadco_cons.id INNER JOIN iwadco_typcon ON iwadco_cons.class=iwadco_typcon.id WHERE billto BETWEEN '" & prevdfrom & "' AND '" & prevdto & "' AND deletedby=0  AND coor_id=" & coor_ & " ORDER BY lname+', '+fname+' '+mname "
lstview.lstDatabase sql, lstCustomer, 1
End Sub


