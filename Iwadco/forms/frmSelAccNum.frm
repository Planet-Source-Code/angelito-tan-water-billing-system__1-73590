VERSION 5.00
Begin VB.Form frmSelAccNum 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Account Number"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
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
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1095
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   600
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin Project1.isButton cmdclose 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      icon            =   "frmSelAccNum.frx":0000
      style           =   6
      caption         =   "&Close"
      iconsize        =   20
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frmSelAccNum.frx":1709C
   End
   Begin Project1.isButton cmdSearch 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      icon            =   "frmSelAccNum.frx":170C4
      style           =   6
      caption         =   "&Search"
      iconsize        =   17
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frmSelAccNum.frx":2DA88
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Select Billing Month"
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
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter Account Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   2145
   End
End
Attribute VB_Name = "frmSelAccNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdSearch_Click()
'On Error Resume Next
Dim todate As Date
    rsCheck
    sql = "SELECT billingdate FROM iwadco_cons,iwadco_coor WHERE coor_id=iwadco_coor.id AND iwadco_cons.id = '" & Me.txtSearch.Text & "'"
    rs.Open sql
    If rs.RecordCount = 0 Then
        MsgBox "No record found!", vbExclamation, Me.Caption
        Exit Sub
    End If
    todate = Combo1.ListIndex + 1 & "/" & rs(0) & "/" & Combo2.Text
'            sql = _
"        SELECT     iwadco_readings.readingno, a.lname + ', ' + a.fname + ' ' + a.mname AS 'Account Name', a.id AS 'Account No', a.address AS Address, " & _
"                      iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname AS Coordinato, iwadco_typcon.type AS 'Type Of Connection', " & _
"                      iwadco_readings.billfrom, iwadco_readings.billto, iwadco_readings.due_date AS 'Due Date', iwadco_readings.previous_reading AS Previous, " & _
"                      iwadco_readings.present_reading AS Present, iwadco_readings.consume AS 'Total Used', iwadco_readings.excess AS Excess, " & _
"                      iwadco_readings.amount_excess-iwadco_typcon.min_rate  AS 'Amount Excess', iwadco_typcon.min_rate, iwadco_readings.arrears AS Arrears, " & _
"                      iwadco_readings.total_amount  " & _
"- " & _
"                          (SELECT     change " & _
"                            From iwadco_payments " & _
"                            WHERE      iwadco_payments.ConID = a.id AND iwadco_payments.id = " & _
"                                                       (SELECT     MAX(id) " & _
"                                                         From iwadco_payments " & _
"                                                         WHERE      iwadco_payments.ConID = a.id)) AS 'Total Amount Due'"
'sql = sql & _
"                      , iwadco_readings.wtax, iwadco_typcon.min_rate + iwadco_readings.amount_excess-iwadco_typcon.min_rate  AS 'Total Amount', DATEADD(DAY, 5, iwadco_readings.due_date) AS dis_date, " & _
"                      3500 - a.amountPaid AS tappings, iwadco_readings.arrears + iwadco_readings.total_amount AS Total_bill, iwadco_onexcss.per_cubic_m, " & _
"                          (SELECT     change " & _
"                            From iwadco_payments " & _
"                            WHERE      iwadco_payments.ConID = a.id AND iwadco_payments.id = " & _
"                                                       (SELECT     MAX(id) " & _
"                                                         From iwadco_payments " & _
"                                                         WHERE      iwadco_payments.ConID = a.id)) AS adv_payments " & _
"FROM         iwadco_cons a INNER JOIN " & _
"                      iwadco_coor ON a.coor_id = iwadco_coor.id INNER JOIN " & _
"                      iwadco_typcon ON a.class = iwadco_typcon.id INNER JOIN " & _
"                      iwadco_readings ON a.id = iwadco_readings.account_no INNER JOIN " & _
"                      iwadco_onexcss ON a.class = iwadco_onexcss.typeid " & _
    "WHERE      (iwadco_readings.billto = '" & todate & "') AND (a.id = '" & txtSearch.Text & "')" & " AND iwadco_readings.deletedby=0"
    sql = "spSOAIndi'" & txtSearch.Text & "','" & todate & "'"
    Debug.Print sql
    If DataEnvironment1.rscmdSOA.State = adStateOpen Then DataEnvironment1.rscmdSOA.Close
    DataEnvironment1.rscmdSOA.Open sql, CN, adOpenStatic, adLockOptimistic
    If DataEnvironment1.rscmdSOA.RecordCount = 0 Then
        MsgBox "No Record Found!", vbInformation, Me.Caption
        Exit Sub
    Else
        rptSOA.DataMember = "cmdSOA"
        rptSOA.Sections("Section1").Controls("Label18").Caption = Combo1.Text
        rptSOA.Sections("Section1").Controls("Label30").Caption = Combo1.Text
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim year As Integer
Combo1.AddItem "January"
Combo1.AddItem "February"
Combo1.AddItem "March"
Combo1.AddItem "April"
Combo1.AddItem "May"
Combo1.AddItem "June"
Combo1.AddItem "July"
Combo1.AddItem "August"
Combo1.AddItem "September"
Combo1.AddItem "October"
Combo1.AddItem "November"
Combo1.AddItem "December"
year = Format(Now, "yyyy")
For year = year - 50 To year + 50
    Combo2.AddItem year
Next
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Combo2.SetFocus
If Combo2.Text = Format(Now, "yyyy") Then: Timer1.Enabled = False: Exit Sub
SendKeys "{DOWN}"
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Combo1.SetFocus
If Combo1.Text = Format(Now, "mmmm") Then: Timer2.Enabled = False: Exit Sub
SendKeys "{DOWN}"
End Sub


Private Sub txtSearch_Change()
txtSearch.Text = str_Filter(txtSearch, 48, 57, 45)
End Sub
