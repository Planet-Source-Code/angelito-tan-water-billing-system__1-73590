VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDate 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6840
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   8
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   2040
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   4080
      Width           =   6015
      Begin VB.CheckBox Check1 
         Caption         =   "Forfeited Commisions"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
      Begin Project1.isButton cmdcancel 
         Height          =   495
         Left            =   4560
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
         _extentx        =   2355
         _extenty        =   873
         icon            =   "frmDate.frx":0000
         style           =   5
         caption         =   "&Cancel"
         iconalign       =   1
         inonthemestyle  =   0
         tooltiptitle    =   ""
         tooltipicon     =   0
         tooltiptype     =   0
         font            =   "frmDate.frx":001C
      End
      Begin Project1.isButton cmdok 
         Height          =   495
         Left            =   3000
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
         _extentx        =   2355
         _extenty        =   873
         icon            =   "frmDate.frx":0044
         style           =   5
         caption         =   "&Ok"
         iconalign       =   1
         inonthemestyle  =   0
         tooltiptitle    =   ""
         tooltipicon     =   0
         tooltiptype     =   65535
         font            =   "frmDate.frx":0060
      End
   End
   Begin MSComctlLib.ListView Lstcoor 
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   1200
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
      Icons           =   "imglst"
      SmallIcons      =   "imglst"
      ColHdrIcons     =   "imglst"
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
   Begin MSComctlLib.ImageList imglst 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmDate.frx":0088
            Key             =   ""
         EndProperty
      EndProperty
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
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmDate.frx":0622
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Year And Month Of Billing Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   3645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Date Form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   675
      Left            =   1200
      TabIndex        =   6
      Top             =   240
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   6
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   6600
      X2              =   0
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   6
      X1              =   0
      X2              =   6600
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdcancel_Click()
Unload Me
Set frmBillingDate = Nothing
End Sub

Private Sub cmdok_Click()
Dim AreaID As Integer
smonth = Combo1.ListIndex + 1
syear = Val(Combo2.Text)
With DataEnvironment1
rsCheck
rs.Open "SELECT * FROM iwadco_area WHERE area = '" & Combo3.Text & "'", CN, adOpenStatic, adLockOptimistic
If rs.RecordCount <> 0 Then
    AreaID = rs(0)
Else
    AreaID = 0
End If
    If .rsCommand1.State = adStateOpen Then .rsCommand1.Close
    If Check1.value = 0 Then
    '    sql = "SHAPE {SELECT     iwadco_coor.id, iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname AS CoorName, ROUND(SUM(iwadco_commisions.amount), 2) AS samount, " & _
                      "ROUND(SUM(iwadco_commisions.commision), 2) AS scom, ROUND(SUM(iwadco_commisions.w_tax), 2) AS stax, ROUND(SUM(iwadco_commisions.total_com), 2) " & _
                      "AS tot_com, ROUND(SUM(iwadco_commisions.amount / 1.12), 2) AS sumEvat, ROUND(SUM(iwadco_commisions.amount - iwadco_commisions.amount / 1.12), 2) " & _
                      "AS shello " & _
"FROM         iwadco_coor INNER JOIN " & _
"                      iwadco_cons ON iwadco_coor.id = iwadco_cons.coor_id INNER JOIN " & _
"                      iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no INNER JOIN " & _
"                      iwadco_commisions ON iwadco_commisions.account_no = iwadco_cons.id " & _
        " WHERE iwadco_commisions.commision<>0 AND iwadco_readings.deletedby IS NULL AND iwadco_commisions.date BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/1/" & syear) - 1 & "' AND iwadco_readings.billto BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/1/" & syear) - 1 & "' GROUP BY iwadco_coor.id,iwadco_coor.lname,iwadco_coor.fname,iwadco_coor.mname}  AS Command1 APPEND ({select iwadco_commisions.id,iwadco_commisions.invoice,iwadco_commisions.date,iwadco_commisions.account_no,ROUND(iwadco_commisions.amount,2) as amount,ROUND(iwadco_commisions.amount/1.12,2) as evat,ROUND(iwadco_commisions.amount-(iwadco_commisions.amount/1.12),2) as hello,ROUND(iwadco_commisions.commision,2) as commision" & _
        ",ROUND(iwadco_commisions.w_tax,2) as w_tax,ROUND(iwadco_commisions.total_com,2) as total_com,iwadco_cons.coor_id from iwadco_commisions inner join iwadco_cons on iwadco_cons.id=iwadco_commisions.account_no WHERE iwadco_commisions.commision<>0 AND iwadco_commisions.date BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/1/" & syear) - 1 & "'}  AS Command2 RELATE 'id' TO 'coor_id') AS Command2"
        'MsgBox sql
        sql = "SHAPE {SELECT     iwadco_coor.id, ROUND(SUM(iwadco_commisions.amount), 2) AS samount, ROUND(SUM(iwadco_commisions.commision), 2) AS scom, " & _
"                      ROUND(SUM(iwadco_commisions.w_tax), 2) AS stax, ROUND(SUM(iwadco_commisions.total_com), 2) AS tot_com, " & _
"                      ROUND(SUM(iwadco_commisions.amount / 1.12), 2) AS sumEvat, ROUND(SUM(iwadco_commisions.amount - iwadco_commisions.amount / 1.12), 2) " & _
"                      AS shello, iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname AS CoorName " & _
"FROM         iwadco_commisions INNER JOIN " & _
"                      iwadco_cons ON iwadco_commisions.account_no = iwadco_cons.id INNER JOIN " & _
"                      iwadco_coor ON iwadco_cons.coor_id = iwadco_coor.id INNER JOIN " & _
"                      iwadco_readings ON iwadco_readings.id=iwadco_commisions.readings_id " & _
"WHERE     (iwadco_commisions.commision <> 0) AND (iwadco_readings.billto BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/1/" & syear) - 1 & "') AND coor_id=" & Lstcoor.SelectedItem.Text
If AreaID <> 0 Then
     sql = sql & " AND area_id=" & AreaID & ""
End If
sql = sql & _
" GROUP BY iwadco_coor.id, iwadco_coor.lname, iwadco_coor.fname, iwadco_coor.mname ORDER BY iwadco_coor.lname}  AS Command1 APPEND ({SELECT     iwadco_commisions.id, iwadco_commisions.invoice, iwadco_commisions.[date], iwadco_commisions.account_no, " & _
"                      ROUND(iwadco_commisions.amount, 2) AS amount, ROUND(iwadco_commisions.amount / 1.12, 2) AS evat, " & _
"                      ROUND(iwadco_commisions.amount - iwadco_commisions.amount / 1.12, 2) AS hello, ROUND(iwadco_commisions.commision, 2) AS commision, " & _
"                      ROUND(iwadco_commisions.w_tax, 2) AS w_tax, ROUND(iwadco_commisions.total_com, 2) AS total_com, iwadco_cons.coor_id, iwadco_cons.lname+', '+iwadco_cons.fname+' ' + iwadco_cons.mname as consName " & _
"FROM         iwadco_commisions INNER JOIN " & _
"                      iwadco_cons ON iwadco_cons.id = iwadco_commisions.account_no INNER JOIN " & _
"                      iwadco_readings ON iwadco_readings.id=iwadco_commisions.readings_id " & _
"WHERE     (iwadco_commisions.commision <> 0) AND (iwadco_readings.billto BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/1/" & syear) - 1 & "') AND coor_id=" & Lstcoor.SelectedItem.Text
If AreaID <> 0 Then
     sql = sql & " AND area_id=" & AreaID & ""
End If
sql = sql & _
"                      }  AS Command2 RELATE 'id' TO 'coor_id') AS Command2"
    Else
        sql = "SHAPE {SELECT     iwadco_coor.id, ROUND(SUM(iwadco_commisions.amount), 2) AS samount, ROUND(SUM(iwadco_commisions.commision), 2) AS scom, " & _
"                      ROUND(SUM(iwadco_commisions.w_tax), 2) AS stax, ROUND(SUM(iwadco_commisions.total_com), 2) AS tot_com, " & _
"                      ROUND(SUM(iwadco_commisions.amount / 1.12), 2) AS sumEvat, ROUND(SUM(iwadco_commisions.amount - iwadco_commisions.amount / 1.12), 2) " & _
"                      AS shello, iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname AS CoorName " & _
"FROM         iwadco_commisions INNER JOIN " & _
"                      iwadco_cons ON iwadco_commisions.account_no = iwadco_cons.id INNER JOIN " & _
"                      iwadco_coor ON iwadco_cons.coor_id = iwadco_coor.id INNER JOIN " & _
"                      iwadco_readings ON iwadco_readings.id=iwadco_commisions.readings_id " & _
"WHERE     (iwadco_commisions.commision = 0) AND (iwadco_readings.billto BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/1/" & syear) - 1 & "') AND coor_id=" & Lstcoor.SelectedItem.Text
If AreaID <> 0 Then
     sql = sql & " AND area_id=" & AreaID & ""
End If
sql = sql & _
" GROUP BY iwadco_coor.id, iwadco_coor.lname, iwadco_coor.fname, iwadco_coor.mname}  AS Command1 APPEND ({SELECT     iwadco_commisions.id, iwadco_commisions.invoice, iwadco_commisions.[date], iwadco_commisions.account_no, " & _
"                      ROUND(iwadco_commisions.amount, 2) AS amount, ROUND(iwadco_commisions.amount / 1.12, 2) AS evat, " & _
"                      ROUND(iwadco_commisions.amount - iwadco_commisions.amount / 1.12, 2) AS hello, ROUND(iwadco_commisions.commision, 2) AS commision, " & _
"                      ROUND(iwadco_commisions.w_tax, 2) AS w_tax, ROUND(iwadco_commisions.total_com, 2) AS total_com, iwadco_cons.coor_id, iwadco_cons.lname+', '+iwadco_cons.fname+' ' + iwadco_cons.mname as consName " & _
"FROM         iwadco_commisions INNER JOIN " & _
"                      iwadco_cons ON iwadco_cons.id = iwadco_commisions.account_no INNER JOIN " & _
"                      iwadco_readings ON iwadco_readings.id=iwadco_commisions.readings_id " & _
"WHERE     (iwadco_commisions.commision = 0) AND (iwadco_readings.billto BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/1/" & syear) - 1 & "') AND coor_id=" & Lstcoor.SelectedItem.Text
If AreaID <> 0 Then
     sql = sql & " AND area_id=" & AreaID & ""
End If
sql = sql & _
"                      }  AS Command2 RELATE 'id' TO 'coor_id') AS Command2"
'"WHERE     (iwadco_commisions.commision <> 0) AND (iwadco_readings.billto BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/1/" & syear) - 1 & "') "
End If
    Debug.Print sql
    .rsCommand1.Open sql
    Unload Me
    rptCommisions.Show
    rptCommisions.Sections("section4").Controls("label3").Caption = "Report  Date: " & Format(smonth & " " & syear, "mmmm yyyy")
End With
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

