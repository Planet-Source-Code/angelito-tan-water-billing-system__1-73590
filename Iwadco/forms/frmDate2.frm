VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDate2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Reports"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "Monthly"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   3360
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Specific Date"
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   3360
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Daily Report"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   6375
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   4200
         TabIndex        =   14
         Top             =   170
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   46202881
         CurrentDate     =   39529
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Top             =   170
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   46202881
         CurrentDate     =   39529
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   3480
         TabIndex        =   16
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
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
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   6375
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   180
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         Format          =   46202881
         CurrentDate     =   39519
      End
      Begin VB.Label Label4 
         Caption         =   "Daily Report for:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   230
         Width           =   1935
      End
   End
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
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3720
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
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   120
      Top             =   120
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
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2880
      Width           =   2775
   End
   Begin MSComctlLib.ListView Lstcoor 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
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
   Begin Project1.isButton cmdclose 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "frmDate2.frx":0000
      Style           =   0
      Caption         =   "&Close"
      IconSize        =   20
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
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
   Begin Project1.isButton cmdselect 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "frmDate2.frx":1709A
      Style           =   0
      Caption         =   "&Select"
      IconSize        =   20
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   65535
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
   Begin MSComctlLib.ImageList imglst 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDate2.frx":173B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Coordinator's Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6615
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
      Left            =   600
      TabIndex        =   7
      Top             =   3840
      Width           =   1875
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
      Left            =   600
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmDate2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdclose_Click()
Unload Me
Set frmBillingDate = Nothing
End Sub

Private Sub cmdselect_Click()
On Error Resume Next
Dim AreaID As Integer
smonth = Combo1.ListIndex + 1
syear = Val(Combo2.Text)
rsCheck
rs.Open "SELECT * FROM iwadco_area WHERE area = '" & Combo3.Text & "'", CN, adOpenStatic, adLockOptimistic
If rs.RecordCount <> 0 Then
    AreaID = rs(0)
Else
    AreaID = 0
End If
If Me.Caption = "Aging" Then
    If DataEnvironment1.rscmdAging.State = adStateOpen Then DataEnvironment1.rscmdAging.Close
    sql = "SHAPE {SELECT iwadco_coor.id, iwadco_coor.lname+', '+iwadco_coor.fname+' '+iwadco_coor.mname as CoorName,SUM(previous_reading) as sumPrev,SUM(present_reading) as sumPres,SUM(consume) as sumCons,SUM(excess) as sumExc,sum(amount_excess) as sumAmExc,sum(total_amount-iwadco_readings.amountpaid) as sumTot FROM iwadco_coor INNER JOIN iwadco_cons ON iwadco_coor.id=iwadco_cons.coor_id INNER JOIN iwadco_readings ON iwadco_readings.account_no = iwadco_cons.id WHERE (iwadco_readings.due_date <= '" & Format(Now, "mm/dd/yyyy") & "') AND (iwadco_readings.total_amount > iwadco_readings.amountpaid AND iwadco_readings.deletedby=0 AND coor_id=" & Lstcoor.SelectedItem.Text
    If AreaID <> 0 Then
         sql = sql & " AND area_id=" & AreaID & ""
    End If
    sql = sql & " AND iwadco_readings.billto BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/01/" & syear) - 1 & "'"
    sql = sql & " ) GROUP BY iwadco_coor.id,iwadco_coor.lname,iwadco_coor.fname,iwadco_coor.mname" & _
    "}  AS cmdAging APPEND ({SELECT iwadco_cons.lname + ', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname AS ConsName, iwadco_typcon.type, iwadco_readings.billfrom, iwadco_readings.billto, iwadco_readings.previous_reading, iwadco_readings.present_reading, iwadco_readings.consume, iwadco_readings.excess, iwadco_typcon.min_rate, iwadco_readings.amount_excess, iwadco_readings.total_amount-iwadco_readings.amountpaid as total_amount, iwadco_coor.id AS CoorID, iwadco_cons.id FROM iwadco_coor INNER JOIN iwadco_cons ON iwadco_coor.id = iwadco_cons.coor_id INNER JOIN iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id INNER JOIN iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no WHERE (iwadco_readings.due_date <= '" & Format(Now, "mm/dd/yyyy") & "') AND (iwadco_readings.total_amount > iwadco_readings.amountpaid)  AND iwadco_readings.deletedby=0 AND coor_id=" & Lstcoor.SelectedItem.Text
    If AreaID <> 0 Then
         sql = sql & " AND area_id=" & AreaID & ""
    End If
    sql = sql & " AND iwadco_readings.billto BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/01/" & syear) - 1 & "'  ORDER BY iwadco_cons.lname,iwadco_cons.fname,iwadco_cons.mname"
    sql = sql & "}  AS Command3 RELATE 'id' TO 'coorid') AS Command3"
    
    Debug.Print sql
    DataEnvironment1.rscmdAging.Open sql
    rptAging.Show
    Unload Me
Else
With DataEnvironment1
    If .rscmdPayments.State = adStateOpen Then .rscmdPayments.Close
    sql = "SELECT     readings.total_amount, payments.amountpayed as amountpayed, readings.total_amount - (SELECT SUM(amountpayed) FROM iwadco_payments WHERE (dateofpayment <= payments.dateofpayment) and (id = payments.id)) AS balance, payments.dateofpayment, payments.change, iwadco_cons.id, iwadco_cons.lname + ', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname AS consname,account_no,invoice FROM         iwadco_payments payments INNER JOIN  iwadco_readings readings ON payments.id = readings.id INNER JOIN    iwadco_cons ON readings.account_no = iwadco_cons.id WHERE readings.deletedby=0 AND coor_id=" & Lstcoor.SelectedItem.Text
    If AreaID <> 0 Then
         sql = sql & " AND area_id=" & AreaID & ""
    End If
    If Option1.Value = True Then
        sql = sql & " AND payments.dateofpayment = '" & DTPicker1.Value & "'"
    ElseIf Option2.Value = True Then
        sql = sql & " AND payments.dateofpayment BETWEEN '" & DTPicker2.Value & "' AND '" & DTPicker3.Value & "'"
    Else
        sql = sql & " AND readings.billto BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/01/" & syear) - 1 & "'"
    End If
    sql = sql & " ORDER BY iwadco_cons.lname,iwadco_cons.fname,iwadco_cons.mname"
    'sql=sql & " GROUP BY readings.
    'sql = sql & " UNION " & _
"(SELECT     a.total_amount, 0 + 0 AS amountpayed, a.total_amount - a.amountpaid AS balance, " & _
"                        0 + 0 AS dateofpayment, 0 + 0 AS change, iwadco_cons.id, iwadco_cons.lname + ', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname AS consname, " & _
"                        a.account_no, '' AS invoice " & _
" FROM         iwadco_readings a INNER JOIN " & _
"                        iwadco_cons ON a.account_no = iwadco_cons.id " & _
" WHERE     a.deletedby = 0 AND coor_id = 19 AND a.billto BETWEEN '2/1/2008' AND '2/29/2008' AND NOT EXIST (SELECT * FROM iwadco_payments INNER JOIN iwadco_readings ON iwadco_payments.conid=iwadco_readings.account_no WHERE iwadco_readings.deletedby=0 AND coor_id=" & Lstcoor.SelectedItem.Text & " ))"
    Debug.Print sql
    .rscmdPayments.Open sql
    sql = "SELECT SUM(iwadco_readings.total_amount) AS sumTot,SUM(iwadco_payments.amountpayed) AS sumAmP,SUM(iwadco_readings.total_amount - iwadco_readings.amountpaid+iwadco_payments.change) as sumBalance,SUM(change) as sumChange FROM         iwadco_payments INNER JOIN  iwadco_readings ON iwadco_payments.id = iwadco_readings.id INNER JOIN    iwadco_cons ON iwadco_readings.account_no = iwadco_cons.id WHERE iwadco_readings.deletedby=0 AND coor_id=" & Lstcoor.SelectedItem.Text
    If AreaID <> 0 Then
        sql = sql & " AND area_id=" & AreaID & ""
    End If
    If Option1.Value = True Then
        sql = sql & " AND iwadco_payments.dateofpayment = '" & DTPicker1.Value & "'"
    ElseIf Option2.Value = True Then
        sql = sql & " AND iwadco_payments.dateofpayment BETWEEN '" & DTPicker2.Value & "' AND '" & DTPicker3.Value & "'"
    Else
        sql = sql & " AND iwadco_readings.billto BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/01/" & syear) - 1 & "'"
    End If
    Debug.Print sql
    If rs.State = adStateOpen Then rs.Close
    rs.Open sql, CN, adOpenStatic, adLockOptimistic
    If Option1.Value = True Then
        rptPayments.Sections("Section4").Controls("Label8").Caption = "Daily Report ( " & DTPicker1.Value & " )"
    ElseIf Option2.Value = True Then
        rptPayments.Sections("Section4").Controls("Label8").Caption = "From " & DTPicker2.Value & " To " & DTPicker3.Value
    Else
        rptPayments.Sections("Section4").Controls("Label8").Caption = Combo1.Text & " " & Combo2.Text
    End If
    rptPayments.Sections("section4").Controls("Label7").Caption = Me.Lstcoor.SelectedItem.SubItems(1) & " ( " & Combo3.Text & " ) "
    rptPayments.Sections("Section5").Controls("Label11").Caption = Format(rs("sumTot"), "###,##0.00")
    rptPayments.Sections("Section5").Controls("Label12").Caption = Format(rs("sumAmP"), "###,##0.00")
    rptPayments.Sections("Section5").Controls("Label15").Caption = Format(rs("sumBalance"), "###,##0.00")
    rptPayments.Sections("Section5").Controls("Label9").Caption = Format(rs("sumChange"), "###,##0.00")
    Unload Me
    rptPayments.Show
End With
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
'Combo1.Text = Format(Now, "mmmm")
Combo2.Text = Format(Now, "yyyy")
'Timer1.Enabled = True
Timer2.Enabled = True
With Lstcoor
    .Icons = imglst
    .SmallIcons = imglst
End With
lstview.lstDatabase "SELECT id as 'ID Number',lname+', '+fname+' '+mname as Name FROM iwadco_coor ORDER BY id ASC", Lstcoor, 1
End Sub

Private Sub Lstcoor_Click()
On Error Resume Next
Combo3.Clear
rsCheck
rs.Open "SELECT area FROM iwadco_area where coor_id = " & Lstcoor.SelectedItem.Text, CN, adOpenStatic, adLockOptimistic
While Not rs.EOF
    Combo3.AddItem rs(0)
    rs.MoveNext
Wend
End Sub

Private Sub Option1_Click()
If Me.Caption = "Aging" Then
    Option3.Value = True
End If
If Option1.Value = True Then
    Frame1.Visible = True
    Frame2.Visible = False
Else
    Frame1.Visible = False
End If
End Sub

Private Sub Option2_Click()
If Me.Caption = "Aging" Then
    Option3.Value = True
End If
If Option2.Value = True Then
    Frame2.Visible = True
    Frame1.Visible = False
Else
    Frame2.Visible = False
End If
End Sub

Private Sub Option3_Click()
Frame1.Visible = False
Frame2.Visible = False
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

