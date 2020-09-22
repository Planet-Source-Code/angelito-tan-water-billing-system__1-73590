VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelCoorRep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Readings Report"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Dont Include Special Accounts"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Now"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   3240
      Width           =   495
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   240
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3240
      Width           =   2775
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
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin MSComctlLib.ListView Lstcoor 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   480
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
      Left            =   5280
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      icon            =   "frmSelCoorRep.frx":0000
      style           =   0
      caption         =   "&Close"
      iconsize        =   20
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "frmSelCoorRep.frx":1709C
   End
   Begin Project1.isButton cmdselect 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      icon            =   "frmSelCoorRep.frx":170C4
      style           =   0
      caption         =   "&Select"
      iconsize        =   20
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   65535
      font            =   "frmSelCoorRep.frx":173E0
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
            Picture         =   "frmSelCoorRep.frx":17408
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
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
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
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1875
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmSelCoorRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdclose_Click()
Unload Me
Set frmSlctCoordinator = Nothing
End Sub

Private Sub cmdselect_Click()
'On Error GoTo err
Dim billdate As Date
Dim todate As Date
Dim asa As String
If Combo1.Text = "" Or Combo2.Text = "" Then
    MsgBox "Please select billing date!", vbExclamation, Me.Caption
    Exit Sub
End If
If Lstcoor.ListItems.Count >= 1 Then
    CoorID = Lstcoor.SelectedItem.Text
    rsCheck
    sql = "SELECT billingdate FROM iwadco_coor WHERE id = " & Lstcoor.SelectedItem.Text
    rs.Open sql
    todate = Combo1.ListIndex + 1 & "/" & rs(0) & "/" & Combo2.Text
    aa1 = Combo1.Text
    aa2 = Combo2.Text
    aa3 = -1
    asa = Lstcoor.SelectedItem.SubItems(1)
    If Combo3.Text <> "" Then
    rsCheck
    rs.Open "SELECT * FROM iwadco_area WHERE area = '" & Combo3.Text & "'", CN, adOpenStatic, adLockOptimistic
    aa3 = rs(0)
    asdasd = Combo3.Text
    End If
    If Me.Caption = "Reading" Then
        If DataEnvironment1.rscmdReadingsTots.State = adStateOpen Then DataEnvironment1.rscmdReadingsTots.Close
        'sql = "SELECT     iwadco_cons.id, SUM(iwadco_readings.previous_reading) AS sumPrev, SUM(iwadco_readings.present_reading) AS sumPres, SUM(iwadco_readings.consume) AS sumCons, SUM(iwadco_readings.arrears) AS sumArr, SUM(iwadco_readings.total_amount) AS sumTot,SUM(iwadco_readings.wtax) AS 'sumTax', SUM(iwadco_typcon.min_rate + iwadco_readings.amount_excess) AS 'sumTotAmount' FROM iwadco_cons INNER JOIN iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no INNER JOIN iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id GROUP BY iwadco_cons.id"
        sql = "SHAPE {SELECT SUM(previous_reading) AS sumPrev,SUM(present_reading) AS sumPres,SUM(consume) AS sumCons,SUM(arrears) AS sumArr,SUM(total_amount) AS sumTot, iwadco_readings.billto,SUM(iwadco_readings.wtax) AS 'sumTax', SUM(iwadco_readings.total_amount-iwadco_readings.wtax-iwadco_readings.arrears) AS 'sumTotAmount',SUM(iwadco_readings.amountpaid) as sumAmtPaid FROM iwadco_cons INNER JOIN iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no INNER JOIN iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id "
        sql = sql & " WHERE billto = '" & todate & "' AND coor_id=" & CoorID
        If aa3 > 0 Then
            sql = sql & " AND area_id = " & aa3
        End If
        If Check1.Value = 1 Then
            sql = sql & " AND iwadco_cons.class<>6"
        End If
        sql = sql & " AND iwadco_readings.deletedby=0 AND iwadco_cons.status='E' GROUP BY iwadco_readings.billto}  AS cmdReadingsTots APPEND ({SELECT iwadco_cons.id,lname+', '+fname+' '+mname as name,iwadco_typcon.type,previous_reading,present_reading,consume,arrears,total_amount,billto, iwadco_readings.wtax,iwadco_readings.total_amount-iwadco_readings.wtax-iwadco_readings.arrears as 'Total Amount',iwadco_readings.amountpaid FROM iwadco_cons INNER JOIN iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no INNER JOIN iwadco_typcon ON iwadco_cons.class=iwadco_typcon.id WHERE billto = '" & todate & "' AND coor_id=" & CoorID
        If aa3 > 0 Then
            sql = sql & " AND area_id = " & aa3
        End If
        If Check1.Value = 1 Then
            sql = sql & " AND iwadco_cons.class<>6"
        End If
        sql = sql & " AND iwadco_readings.deletedby=0 AND iwadco_cons.status='E' }  AS cmdReadings RELATE 'billto' TO 'billto') AS cmdReadings"
        Debug.Print sql
        DataEnvironment1.rscmdReadingsTots.Open sql
        'If DataEnvironment1.rscmdReadingsTots.RecordCount = 0 Then
        '    MsgBox "No Record Found!", vbInformation, Me.Caption
        '    Exit Sub
        'Else
        '    Unload Me
        'End If
        If DataEnvironment1.rscmdReadingsTots.RecordCount = 0 Then
            MsgBox "No Record Found!", vbInformation, Me.Caption
            Exit Sub
        Else
            Unload Me
            If aa1 = "" Or aa2 = "" Then Exit Sub
            rptReadings.DataMember = "cmdReadingsTots"
            rptReadings.Sections("Section4").Controls("lblDate").Caption = aa1 & " " & aa2
            rptReadings.Sections("Section4").Controls("lblcoorname").Caption = asa & " Area - " & asdasd
            aa1 = ""
            aa2 = ""
            aa3 = -1
        End If
    ElseIf Me.Caption = "Reading Form" Then
    
        rsCheck
        sql = "SELECT billingdate FROM iwadco_coor WHERE id = " & Lstcoor.SelectedItem.Text
        rs.Open sql
        todate = Combo1.ListIndex + 1 & "/" & rs(0) & "/" & Combo2.Text
        'If DataEnvironment1.rscmdReadingsTots.State = adStateOpen Then DataEnvironment1.rscmdReadingsTots.Close
        If Combo1.ListIndex = 0 Then
            todate = 12 & "/" & rs(0) & "/" & Val(Combo2.Text) - 1
        Else
            todate = (Combo1.ListIndex - 1) & "/" & rs(0) & "/" & Combo2.Text
        End If
        'sql = "SHAPE {SELECT SUM(previous_reading) AS sumPrev,SUM(present_reading) AS sumPres,SUM(consume) AS sumCons,SUM(arrears+total_amount) AS sumArr,SUM(total_amount) AS sumTot, iwadco_readings.billto FROM iwadco_cons INNER JOIN iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no"
        'sql = sql & " WHERE billto = '" & todate & "' AND coor_id=" & CoorID
        'If aa3 > 0 Then
        '    sql = sql & " AND area_id = " & aa3
        'End If
        'sql = sql & " AND iwadco_readings.deletedby=0 GROUP BY iwadco_readings.billto}  AS cmdReadingsTots APPEND ({SELECT iwadco_cons.id,lname+', '+fname+' '+mname as name,iwadco_typcon.type,previous_reading,present_reading,consume,arrears+total_amount as arrears,total_amount,billto FROM iwadco_cons INNER JOIN iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no INNER JOIN iwadco_typcon ON iwadco_cons.class=iwadco_typcon.id WHERE billto = '" & todate & "' AND coor_id=" & CoorID
        'If aa3 > 0 Then
        '    sql = sql & " AND area_id = " & aa3
        'End If
        'Debug.Print sql
        'sql = sql & " AND iwadco_readings.deletedby=0}  AS cmdReadings RELATE 'billto' TO 'billto') AS cmdReadings"
        If DataEnvironment1.rsCommand6.State = 1 Then DataEnvironment1.rsCommand6.Close
        sql = "spReadingForm'" & todate & "','" & CoorID & "','" & aa3 & "'"
        Debug.Print sql
        DataEnvironment1.rsCommand6.Open sql
        If DataEnvironment1.rsCommand6.RecordCount = 0 Then
            MsgBox "No Record Found!", vbInformation, Me.Caption
            Exit Sub
        Else
            Unload Me
            If aa1 = "" Or aa2 = "" Then Exit Sub
            DataReport4.DataMember = "command6"
            'DataReport4.Sections("Section4").Controls("lblDate").Caption = aa1 & " " & aa2
            'DataReport4.Sections("Section4").Controls("lblcoorname").Caption = asa & " Area - " & asdasd
            DataReport4.Show
            aa1 = ""
            aa2 = ""
            aa3 = -1
        End If
    Else
    If Combo3.Text = "" Then
        MsgBox "Please select area!", vbInformation, "SOA"
        Exit Sub
    End If
        
sql = "spSOA'" & todate & "','" & Int(Lstcoor.SelectedItem.Text) & "','" & aa3 & "'"
Debug.Print sql
        If DataEnvironment1.rscmdSOA.State = adStateOpen Then DataEnvironment1.rscmdSOA.Close
        DataEnvironment1.rscmdSOA.Open sql, CN, adOpenStatic, adLockOptimistic
        If DataEnvironment1.rscmdSOA.RecordCount = 0 Then
            MsgBox "No Record Found!", vbInformation, Me.Caption
            DataEnvironment1.rscmdSOA.Close
            Exit Sub
        Else
            If Me.Caption = "SOA Summary" Then
                ChngPrinterOrientationLandscape Me
                rptSOASummary.DataMember = "cmdSOA"
                rptSOASummary.Sections("section2").Controls("label4").Caption = "Coordinators Name: " & Lstcoor.SelectedItem.SubItems(1) & "( Area - " & Combo3.Text & " )                "
                rptSOASummary.Sections("section2").Controls("label7").Caption = "Billing Date: " & Combo1.Text & " " & Combo2.Text
                sql = _
                "SELECT     SUM(iwadco_readings.amount_excess-iwadco_typcon.min_rate) AS 'sumAE', SUM(iwadco_readings.arrears) AS sumArrears," & _
                              "SUM((iwadco_readings.total_amount+(3500-iwadco_cons.amountPaid))) AS 'sumTotal', SUM(iwadco_readings.wtax) as 'sumTax',SUM(iwadco_readings.total_amount-iwadco_readings.wtax-iwadco_readings.arrears) as 'sumAD' " & _
                "FROM         iwadco_cons INNER JOIN " & _
                              "iwadco_coor ON iwadco_cons.coor_id = iwadco_coor.id INNER JOIN " & _
                              "iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id INNER JOIN " & _
                              "iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no " & _
                "WHERE     iwadco_readings.billto = '" & todate & "' AND iwadco_coor.id = " & Int(Lstcoor.SelectedItem.Text) & " AND iwadco_readings.deletedby=0"
                sql = sql & " AND area_id = " & aa3
                Debug.Print sql
                rsCheck
                rs.Open sql, CN, adOpenStatic, adLockOptimistic
                'rptSOASummary.Sections("section5").Controls("label6").Caption = Format(rs(0), "###,##0.00")
                rptSOASummary.Sections("section5").Controls("label8").Caption = Format(rs(1), "###,##0.00")
                rptSOASummary.Sections("section5").Controls("label9").Caption = Format(rs(2), "###,##0.00")
                rptSOASummary.Sections("section5").Controls("label10").Caption = Format(rs(3), "###,##0.00")
                rptSOASummary.Sections("section5").Controls("label11").Caption = Format(rs(4), "###,##0.00")
                Unload Me
            Else
                rptSOA.DataMember = "cmdSOA"
                rptSOA.Sections("Section1").Controls("Label18").Caption = Combo1.Text
                rptSOA.Sections("Section1").Controls("Label30").Caption = Combo1.Text
                'rsCheck
                'rs.Open sql
                'rptSOA.Sections("Section1").Controls("Label27").Caption = DateAdd("d", 5, rs("Due Date")) '& rs("Due Date")
                Unload Me
            End If
        End If
    End If
End If
err:
Exit Sub
End Sub

Private Sub Command1_Click()
Timer2.Enabled = True
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Lstcoor_Click()
Combo3.Clear
rsCheck
rs.Open "SELECT area FROM iwadco_area where coor_id = " & Lstcoor.SelectedItem.Text, CN, adOpenStatic, adLockOptimistic
While Not rs.EOF
    Combo3.AddItem rs(0)
    rs.MoveNext
Wend
End Sub

Private Sub lstCoor_DblClick()
cmdselect_Click
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
With Lstcoor
    .Icons = imglst
    .SmallIcons = imglst
End With
lstview.lstDatabase "SELECT id as 'ID Number',lname+', '+fname+' '+mname as Name FROM iwadco_coor ORDER BY id ASC", Lstcoor, 1
End Sub

Private Sub Timer1_Timer()
On Error GoTo err
Combo1.SetFocus
If Combo1.Text = Format(Now, "mmmm") Then
    Timer1.Enabled = False
    Exit Sub
End If
SendKeys "{DOWN}"
err:
Select Case err
Case 0
Case Else
Timer1.Enabled = False
End Select
End Sub

Private Sub Timer2_Timer()
On Error GoTo err
Combo2.SetFocus
If Combo2.Text = Format(Now, "yyyy") Then
    Timer2.Enabled = False
    Exit Sub
End If
SendKeys "{DOWN}"
err:
Select Case err
Case 0
Case Else
Timer2.Enabled = False
End Select
End Sub



