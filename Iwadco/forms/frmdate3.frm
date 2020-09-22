VERSION 5.00
Begin VB.Form frmdate3 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   6015
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
         TabIndex        =   3
         Top             =   240
         Width           =   3735
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
      Begin VB.CheckBox Check1 
         Caption         =   "Incomplete Payments Only"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   2415
      End
      Begin Project1.isButton cmdcancel 
         Height          =   495
         Left            =   4560
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Icon            =   "frmdate3.frx":0000
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
         Left            =   3000
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Icon            =   "frmdate3.frx":001C
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
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   6
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tapping Fees"
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
      TabIndex        =   7
      Top             =   240
      Width           =   3675
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
      TabIndex        =   6
      Top             =   840
      Width           =   3645
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmdate3.frx":0038
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "frmdate3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Check1_Click()
    If Timer1.Enabled = True Then Timer1.Enabled = False
    If Timer2.Enabled = True Then Timer2.Enabled = False
If Check1.Value = 1 Then
    Combo1.Enabled = False
    Combo2.Enabled = False
Else
    Combo1.Enabled = True
    Combo2.Enabled = True
End If
End Sub

Private Sub cmdcancel_Click()
Unload Me
Set frmBillingDate = Nothing
End Sub

Private Sub cmdok_Click()
smonth = Combo1.ListIndex + 1
syear = Val(Combo2.Text)
With DataEnvironment1
    If Check1.Value = 0 Then
        sql = "SHAPE {select iwadco_coor.id,iwadco_coor.lname+', '+iwadco_coor.fname as coorname,SUM(amountPaid) as samountPaid,SUM(3500-amountPaid) as sbalance from iwadco_coor inner join iwadco_cons on iwadco_coor.id=iwadco_cons.coor_id inner join iwadco_tappingfee on iwadco_cons.id=iwadco_tappingfee.account_no where iwadco_tappingfee.dateofpayment BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/1/" & syear) - 1 & "' AND iwadco_cons.status='E' " & _
        "group by iwadco_coor.id,iwadco_coor.fname,iwadco_coor.lname}  AS Command4 APPEND ({select iwadco_coor.id as coorid,iwadco_cons.id,iwadco_cons.lname+', '+iwadco_cons.fname as name, iwadco_cons.dateregistered,iwadco_typcon.type,iwadco_cons.amountPaid,(3500-iwadco_cons.amountPaid) as balance from iwadco_cons inner join iwadco_coor on iwadco_coor.id=iwadco_cons.coor_id inner join iwadco_typcon on iwadco_typcon.id=iwadco_cons.class inner join iwadco_tappingfee on iwadco_cons.id=iwadco_tappingfee.account_no where iwadco_tappingfee.dateofpayment BETWEEN '" & smonth & "/1/" & syear & "' AND '" & DateAdd("m", 1, smonth & "/1/" & syear) - 1 & "' ORDER BY iwadco_cons.lname,iwadco_cons.fname,iwadco_cons.mname}  AS cmdtapping RELATE 'id' TO 'coorid') AS cmdtapping"
        Debug.Print sql
    Else
        sql = "SHAPE {select iwadco_coor.id,iwadco_coor.lname+', '+iwadco_coor.fname as coorname,SUM(amountPaid) as samountPaid,SUM(3500-amountPaid) as sbalance from iwadco_coor inner join iwadco_cons on iwadco_coor.id=iwadco_cons.coor_id where iwadco_cons.tappingStatus='I' " & _
        "group by iwadco_coor.id,iwadco_coor.fname,iwadco_coor.lname}  AS Command4 APPEND ({select iwadco_coor.id as coorid,iwadco_cons.id,iwadco_cons.lname+', '+iwadco_cons.fname as name, iwadco_cons.dateregistered,iwadco_typcon.type,iwadco_cons.amountPaid,(3500-iwadco_cons.amountPaid) as balance from iwadco_cons inner join iwadco_coor on iwadco_coor.id=iwadco_cons.coor_id inner join iwadco_typcon on iwadco_typcon.id=iwadco_cons.class where iwadco_cons.tappingStatus='I'  AND iwadco_cons.status='E' ORDER BY iwadco_cons.lname,iwadco_cons.fname,iwadco_cons.mname}  AS cmdtapping RELATE 'id' TO 'coorid') AS cmdtapping"
        Debug.Print sql
    End If
    If .rsCommand4.State = adStateOpen Then .rsCommand4.Close
    .rsCommand4.Open sql
    Unload Me
    DataReport1.Sections("section4").Controls("label4").Caption = DataReport1.Sections("section4").Controls("label4").Caption & Combo1.List(smonth - 1) & " " & syear
    DataReport1.Show
    Timer1.Enabled = False
    Timer2.Enabled = False
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


