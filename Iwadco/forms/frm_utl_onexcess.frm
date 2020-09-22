VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_utl_onexcess_min 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cubic Per Meter"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10800
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Connection Type And Cubic Per Meter"
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
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   10575
      Begin MSComctlLib.ListView LST_ONEXCSS_MIN_RATE 
         Height          =   3495
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   6165
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
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin Project1.isButton cmdCancel 
      Height          =   420
      Left            =   1680
      TabIndex        =   0
      Top             =   4440
      Width           =   1335
      _extentx        =   2355
      _extenty        =   741
      icon            =   "frm_utl_onexcess.frx":0000
      style           =   5
      caption         =   "&Cancel"
      iconsize        =   20
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frm_utl_onexcess.frx":1709C
   End
   Begin Project1.isButton cmdedit 
      Height          =   420
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
      _extentx        =   2355
      _extenty        =   741
      icon            =   "frm_utl_onexcess.frx":170C4
      style           =   5
      caption         =   "&Edit"
      iconsize        =   20
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frm_utl_onexcess.frx":2C238
   End
End
Attribute VB_Name = "FRM_UTL_ONEXCESS_MIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
Set FRM_UTL_ONEXCESS_MIN = Nothing
End Sub

Private Sub cmdedit_Click()
If conName <> "" Then
    FRM_EDIT_ONEXCESS_MIN.Show 1
End If
End Sub

Private Sub Form_Load()
sql = "SELECT typeid as 'Connection Type', per_cubic_m as 'Cubic Per Meter' FROM iwadco_onexcss WHERE status = 'E'"
lstview.lstDatabase sql, LST_ONEXCSS_MIN_RATE, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
conName = ""
End Sub

Private Sub LST_ONEXCSS_MIN_RATE_Click()
conName = LST_ONEXCSS_MIN_RATE.SelectedItem.Text
End Sub

Private Sub LST_ONEXCSS_MIN_RATE_DblClick()
Call cmdedit_Click
End Sub

