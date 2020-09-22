VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDate6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Commision Summary Report"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   3480
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin Project1.isButton cmdcancel 
         Height          =   495
         Left            =   4560
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Icon            =   "frmDate6.frx":0000
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
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Icon            =   "frmDate6.frx":001C
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
   Begin MSComctlLib.ListView Lstcoor 
      Height          =   2175
      Left            =   120
      TabIndex        =   5
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
            Picture         =   "frmDate6.frx":0038
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   6
      X1              =   0
      X2              =   6600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   6600
      X2              =   0
      Y1              =   3120
      Y2              =   3120
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
      TabIndex        =   7
      Top             =   240
      Width           =   4815
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
      Picture         =   "frmDate6.frx":05D2
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "frmDate6"
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
    If .rscmdComSum.State = adStateOpen Then .rsCommand1.Close
    
        sql = "spComSum'" & smonth & "/1/" & syear & "','" & DateAdd("m", 1, smonth & "/1/" & syear) - 1 & "'," & Lstcoor.SelectedItem.Text
    Debug.Print sql
    
    .rscmdComSum.Open sql
    Unload Me
    DataReport5.Sections("section4").Controls("label11").Caption = DataReport5.Sections("section4").Controls("label11").Caption & "- Report  Date: " & Format(smonth & " " & syear, "mmmm yyyy")
    DataReport5.Show
    
    
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



