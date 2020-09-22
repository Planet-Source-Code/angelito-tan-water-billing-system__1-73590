VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_utl_con_min_rate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connection Minimum Rate"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10800
   Icon            =   "frmutilities.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imglst 
      Left            =   7080
      Top             =   360
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
            Picture         =   "frmutilities.frx":617A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmutilities.frx":C304
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   10575
      Begin Project1.isButton cmdCancel 
         Height          =   420
         Left            =   1560
         TabIndex        =   2
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmutilities.frx":1248E
         Style           =   5
         Caption         =   "&Cancel"
         IconSize        =   20
         IconAlign       =   1
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
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
      Begin Project1.isButton cmdedit 
         Height          =   420
         Left            =   120
         TabIndex        =   3
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmutilities.frx":29528
         Style           =   5
         Caption         =   "&Edit"
         IconSize        =   20
         IconAlign       =   1
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
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
      Begin MSComctlLib.ListView LST_CON_MIN_RATE 
         Height          =   3615
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmutilities.frx":3E69A
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Change connection name, minimum rate and per cubic meter"
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
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Connection Minimum Rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "FRM_UTL_CON_MIN_RATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Double

Private Sub cmdcancel_Click()
Unload Me
Set FRM_UTL_CON_MIN_RATE = Nothing
End Sub

Private Sub cmdedit_Click()
If conName <> "" Then
    frm_edit_con_min.Show 1
End If
End Sub

Private Sub Form_Load()
'Initialize and fill up the ListView
With LST_CON_MIN_RATE
    .Icons = imglst
    .SmallIcons = imglst
End With
'listview items
sql = "SELECT type as 'Connection Type', min_rate as 'Minimum Rate', per_cubic_m as 'Per Cubic Meter'FROM iwadco_typcon,iwadco_onexcss WHERE iwadco_typcon.id = iwadco_onexcss.typeid AND iwadco_typcon.status = 'E' AND iwadco_onexcss.status ='E'"
lstview.lstDatabase sql, LST_CON_MIN_RATE, 2


End Sub

Private Sub Form_Unload(Cancel As Integer)
conName = ""
End Sub

Private Sub LST_CON_MIN_RATE_Click()
conName = LST_CON_MIN_RATE.SelectedItem.Text
rsCheck
rs.Open "SELECT * FROM iwadco_typcon WHERE type ='" & conName & "'"
FK_ID = rs(0)
End Sub

Private Sub LST_CON_MIN_RATE_DblClick()
Call cmdedit_Click
End Sub
