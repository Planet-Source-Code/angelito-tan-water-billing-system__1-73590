VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDate4 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6825
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6825
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consumer"
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "Disconnected"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Registered"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
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
      Format          =   46333953
      CurrentDate     =   39465
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
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
      Format          =   46333953
      CurrentDate     =   39465
   End
   Begin Project1.isButton cmdcancel 
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "frmDate4.frx":0000
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
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "frmDate4.frx":001C
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
   Begin MSComctlLib.ListView Lstcoor 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   2520
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
            Picture         =   "frmDate4.frx":0038
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
      TabIndex        =   14
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Coordinators Name"
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
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TO"
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
      Left            =   3360
      TabIndex        =   11
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FROM"
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
      Left            =   3360
      TabIndex        =   10
      Top             =   1440
      Width           =   525
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
      TabIndex        =   9
      Top             =   720
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consumer Reports"
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
      TabIndex        =   8
      Top             =   120
      Width           =   5160
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmDate4.frx":05D2
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmDate4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
Dim area_ As Integer
If Combo3.Text <> "" Then
rsCheck
rs.Open "SELECT id FROM iwadco_area WHERE area='" & Combo3.Text & "'", CN, adOpenStatic, adLockOptimistic
area_ = rs(0)
Else
area_ = 0
End If
With DataEnvironment1
If .rsCommand5.State = adStateOpen Then .rsCommand5.Close
'.rsCommand5.Open "SHAPE {select iwadco_coor.id,iwadco_coor.lname+', '+iwadco_coor.fname+' '+iwadco_coor.mname as coor_name from iwadco_coor inner join iwadco_cons on iwadco_cons.coor_id=iwadco_coor.id where iwadco_cons.dateregistered between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "'}  AS Command5 APPEND ({SELECT     iwadco_cons.id, iwadco_cons.lname + ' , ' + iwadco_cons.fname + ' ' + iwadco_cons.mname AS cons_name, iwadco_area.area, iwadco_typcon.type, " & _
                     " iwadco_cons.tel , iwadco_cons.coor_id " & _
                     " FROM         iwadco_cons INNER JOIN " & _
                     " iwadco_area ON iwadco_cons.area_id = iwadco_area.id INNER JOIN " & _
                     " iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id where iwadco_cons.dateregistered between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "'}  AS cmdConsumers RELATE 'id' TO 'coor_id') AS cmdConsumers"
sql = "SHAPE {select iwadco_coor.id,iwadco_coor.lname+', '+iwadco_coor.fname+' '+iwadco_coor.mname as coor_name from iwadco_coor "
If Lstcoor.SelectedItem.Text <> "" Then sql = sql & " where id = '" & Lstcoor.SelectedItem.Text & "'"
sql = sql & "}  AS Command5 APPEND ({SELECT     iwadco_cons.id, iwadco_cons.lname + ' , ' + iwadco_cons.fname + ' ' + iwadco_cons.mname AS cons_name, iwadco_area.area, iwadco_typcon.type, " & _
                      "iwadco_cons.tel , iwadco_cons.coor_id " & _
"FROM         iwadco_cons INNER JOIN " & _
"                      iwadco_area ON iwadco_cons.area_id = iwadco_area.id INNER JOIN " & _
"                      iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id "
If Option1.Value = True Then
    sql = sql & " where iwadco_cons.status='E' AND iwadco_cons.dateregistered between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "' "
Else
    sql = sql & " where iwadco_cons.status='X' AND iwadco_cons.DateDisconnected BETWEEN '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "' "
End If
If area_ <> 0 Then
    sql = sql & " AND area_id = " & area_
End If
sql = sql & " ORDER BY iwadco_cons.lname,iwadco_cons.fname,iwadco_cons.mname}  AS cmdConsumers RELATE 'id' TO 'coor_id') AS cmdConsumers "
Debug.Print sql
.rsCommand5.Open sql
    rptRepCons.Sections("section4").Controls("label13").Caption = Format(DTPicker1.Value, "mmmm dd, yyyy")
    rptRepCons.Sections("section4").Controls("label12").Caption = Format(DTPicker2.Value, "mmmm dd, yyyy")
    ChngPrinterOrientationPortrait Me
    Unload Me
    Load rptRepCons
End With
End Sub

Private Sub Form_Load()
lstview.lstDatabase "SELECT id as 'ID Number',lname+', '+fname+' '+mname as Name FROM iwadco_coor ORDER BY id ASC", Lstcoor, 1
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
