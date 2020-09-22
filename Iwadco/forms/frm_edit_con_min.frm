VERSION 5.00
Begin VB.Form frm_edit_con_min 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connection Name, Minimum Amount, Per Cubic Meter Settings"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6705
   Icon            =   "frm_edit_con_min.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6495
      Begin VB.TextBox txtpercm 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   3
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "frm_edit_con_min.frx":617A
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtminAmount 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   2
         Top             =   1320
         Width           =   2055
      End
      Begin Project1.isButton cmdcancel 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "frm_edit_con_min.frx":6180
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
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "frm_edit_con_min.frx":1D21A
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Per Cubic Meter ( Php ) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Connection Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Minimum Amount ( Php ):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2295
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Default Settings Of Inpart Water Development Corporation (IWADCO)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frm_edit_con_min.frx":342B4
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frm_edit_con_min"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
rs("type") = txtname.Text
rs("min_rate") = txtminAmount.Text
rs.Update
rs("per_cubic_m") = txtpercm.Text
rs.Update
MsgBox "Changes Has Been Succesfully Made !", vbInformation, "Default Setting"
sql = "SELECT type as 'Connection Type', min_rate as 'Minimum Rate', per_cubic_m as 'Per Cubic Meter'FROM iwadco_typcon,iwadco_onexcss WHERE iwadco_typcon.id = iwadco_onexcss.typeid AND iwadco_typcon.status = 'E' AND iwadco_onexcss.status ='E'"
lstview.lstDatabase sql, FRM_UTL_CON_MIN_RATE.LST_CON_MIN_RATE, 2
Unload Me
End Sub

Private Sub Form_Load()
rsCheck
rs.Open "SELECT * FROM iwadco_typcon,iwadco_onexcss WHERE type ='" & conName & "'AND typeid =" & FK_ID
txtname.Text = rs("type")
txtminAmount.Text = rs("min_rate")
txtpercm.Text = rs("per_cubic_m")
End Sub

Private Sub txtminAmount_Change()
txtminAmount.Text = str_Filter(txtminAmount, 48, 57, 46)
End Sub

Private Sub txtpercm_Change()
txtpercm.Text = str_Filter(txtpercm, 48, 57, 46)
End Sub
