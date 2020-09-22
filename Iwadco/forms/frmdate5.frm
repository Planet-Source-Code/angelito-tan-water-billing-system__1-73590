VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmdate5 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6840
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20185089
      CurrentDate     =   39465
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1800
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20185089
      CurrentDate     =   39465
   End
   Begin Project1.isButton cmdcancel 
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "frmdate5.frx":0000
      Style           =   5
      Caption         =   "&Cancel"
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
   Begin Project1.isButton cmdok 
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "frmdate5.frx":001C
      Style           =   5
      Caption         =   "&Ok"
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
            Picture         =   "frmdate5.frx":0038
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmdate5.frx":05D2
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Summary Report"
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
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Date Of Report"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   720
      Width           =   2130
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Date From:"
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
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Date To:"
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
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   1560
   End
End
Attribute VB_Name = "frmdate5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
ChngPrinterOrientationLandscape Me
'sql = "" & _
"SELECT     a.coor_id, iwadco_coor.lname + ', ' + iwadco_coor.fname AS coor_name, " & _
"                          (SELECT     COUNT(id) " & _
"                            From iwadco_cons " & _
"                            WHERE      coor_id = a.coor_id AND iwadco_cons.status='E' AND dateregistered BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AS cntCons, " & _
"                          (SELECT     COUNT(id) " & _
"                            From iwadco_cons " & _
"                            WHERE      coor_id = a.coor_id AND iwadco_cons.status='X' AND datedisconnected BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AS disCons, " & _
"                          (SELECT     SUM(consume) " & _
"                            FROM          iwadco_readings INNER JOIN " & _
"                                                   iwadco_cons ON iwadco_cons.id = iwadco_readings.account_no " & _
"                            WHERE      iwadco_cons.coor_id = a.coor_id AND iwadco_readings.trxdate BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND iwadco_readings.deletedby=0 AND iwadco_cons.status='E' AND iwadco_cons.class<>6) AS sumConsume, " & _
"                          (SELECT     SUM(excess) " & _
"                            FROM          iwadco_readings INNER JOIN " & _
"                                                   iwadco_cons ON iwadco_cons.id = iwadco_readings.account_no " & _
"                            WHERE      iwadco_cons.coor_id = a.coor_id AND iwadco_readings.trxdate BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND iwadco_readings.deletedby=0 AND iwadco_cons.status='E' AND iwadco_cons.class<>6) AS sumExcess, " & _
"                          (SELECT     SUM(amount_excess - iwadco_typcon.min_rate) " & _
"                            FROM          iwadco_readings INNER JOIN " & _
"                                                   iwadco_cons ON iwadco_cons.id = iwadco_readings.account_no INNER JOIN " & _
"                                                   iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id " & _
"                            WHERE      iwadco_cons.coor_id = a.coor_id AND iwadco_readings.trxdate BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND iwadco_readings.deletedby=0 AND iwadco_cons.status='E' AND iwadco_cons.class<>6) AS sumAmtExcess, " & _
"                          (SELECT     SUM(wtax) " & _
"                            FROM          iwadco_readings INNER JOIN " & _
"                                                   iwadco_cons ON iwadco_cons.id = iwadco_readings.account_no " & _
"                            WHERE      iwadco_cons.coor_id = a.coor_id AND iwadco_readings.trxdate BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND iwadco_readings.deletedby=0 AND iwadco_cons.status='E' AND iwadco_cons.class<>6) AS sumWtax, "
'sql = sql & " (SELECT     SUM(total_amount - wtax) " & _
"                            FROM          iwadco_readings INNER JOIN " & _
"                                                   iwadco_cons ON iwadco_cons.id = iwadco_readings.account_no " & _
"                            WHERE      iwadco_cons.coor_id = a.coor_id AND iwadco_readings.trxdate BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND iwadco_readings.deletedby=0 AND iwadco_cons.status='E' AND iwadco_cons.class<>6) AS billing, " & _
"                          (SELECT     SUM(total_amount) " & _
"                            FROM          iwadco_readings INNER JOIN " & _
"                                                   iwadco_cons ON iwadco_cons.id = iwadco_readings.account_no " & _
"                            WHERE      iwadco_cons.coor_id = a.coor_id AND iwadco_readings.trxdate BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND iwadco_readings.deletedby=0 AND iwadco_cons.status='E' AND iwadco_cons.class<>6) AS sumAmtDue, " & _
"                          (SELECT     SUM(iwadco_readings.amountpaid) " & _
"                            FROM          iwadco_readings INNER JOIN " & _
"                                                   iwadco_cons ON iwadco_cons.id = iwadco_readings.account_no " & _
"                            WHERE      iwadco_cons.coor_id = a.coor_id AND iwadco_readings.trxdate BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND iwadco_readings.deletedby=0 AND iwadco_cons.status='E' AND iwadco_cons.class<>6) AS sumAmtPaid, " & _
"                          (SELECT     SUM(iwadco_readings.total_amount - iwadco_readings.amountpaid) " & _
"                            FROM          iwadco_readings INNER JOIN " & _
"                                                   iwadco_cons ON iwadco_cons.id = iwadco_readings.account_no " & _
"                            WHERE      iwadco_cons.coor_id = a.coor_id AND iwadco_readings.trxdate BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND iwadco_readings.deletedby=0 AND iwadco_cons.status='E' AND iwadco_cons.class<>6) AS sumArrears " & _
"FROM         iwadco_cons a INNER JOIN " & _
"                      iwadco_coor ON iwadco_coor.id = a.coor_id " & _
"GROUP BY a.coor_id, iwadco_coor.lname, iwadco_coor.fname "

Debug.Print sql
If DataEnvironment1.rscmdSummary.State = 1 Then DataEnvironment1.rscmdSummary.Close
DataEnvironment1.rscmdSummary.Open "spListSummary'" & DTPicker1.Value & "','" & DTPicker2.Value & "'", CN, adOpenStatic, adLockOptimistic
DataReport2.Sections("Section4").Controls("Label13").Caption = "Date : " & Format(DTPicker1.Value, "mmmm dd, yyyy") & " to " & Format(DTPicker2.Value, "mmmm dd,yyyy")
Unload Me
End Sub

Private Sub Form_Load()
DTPicker1.Value = Format(Now, "mm/dd/yyyy")
DTPicker2.Value = Format(Now, "mm/dd/yyyy")
End Sub
