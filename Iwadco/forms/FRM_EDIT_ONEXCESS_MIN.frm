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
         Text            =   "FRM_EDIT_ONEXCESS_MIN.frx":0000
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
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FRM_EDIT_ONEXCESS_MIN.frx":0006
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
   Begin Project1.isButton cmdsave 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FRM_EDIT_ONEXCESS_MIN.frx":170A0
      Style           =   5
      Caption         =   "&Save"
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
lstview.lstDatabase sql, FRM_UTL_ONEXCESS_MIN.LST_ONEXCSS_MIN_RATE
End Sub

Private Sub Form_Load()
rsCheck
rs.Open "SELECT * FROM iwadco_onexcss WHERE typeid ='" & conName & "'"
txttype.Text = rs(1)
txtPCM.Text = rs(2)
End Sub

