VERSION 5.00
Begin VB.Form frm_edit_onexcess_min 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Connection Type and Per Cubic Meter"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtPCM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txttype 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frm_edit_onexcess.frx":0000
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Per Cubic Meter (P):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Connection Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
   Begin Project1.isButton cmdcance 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      icon            =   "frm_edit_onexcess.frx":0006
      style           =   5
      caption         =   "&Cancel"
      iconsize        =   20
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frm_edit_onexcess.frx":170A2
   End
   Begin Project1.isButton cmdsave 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      icon            =   "frm_edit_onexcess.frx":170CA
      style           =   5
      caption         =   "&Save"
      iconsize        =   20
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frm_edit_onexcess.frx":2E166
   End
End
Attribute VB_Name = "FRM_EDIT_ONEXCESS_MIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcance_Click()
Unload Me
Set FRM_EDIT_ONEXCESS_MIN = Nothing
End Sub

Private Sub cmdsave_Click()
rs(1) = txttype.Text
rs(2) = txtPCM.Text
rs.Update
MsgBox "Changes Has Been Succesfully Made !", vbInformation, Me.Caption
Unload Me
sql = "SELECT con_type as 'Connection Type', per_cubic_m as 'Cubic Per Meter' FROM iwadco_onexcss WHERE status = 'E'"
lstview.lstDatabase sql, FRM_UTL_ONEXCESS_MIN.LST_ONEXCSS_MIN_RATE, 1
End Sub

Private Sub Form_Load()
rsCheck
rs.Open "SELECT * FROM iwadco_onexcss WHERE typeid ='" & conName & "'"
txttype.Text = rs(1)
txtPCM.Text = rs(2)
End Sub

