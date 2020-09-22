VERSION 5.00
Begin VB.Form frm_edit_default 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Default Excess Settings"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   5895
      Begin VB.TextBox txttype 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "frm_edit_default.frx":0000
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtVALUE 
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
         Left            =   1920
         TabIndex        =   1
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Connection Type:"
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
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "VALUE :"
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
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
         Width           =   855
      End
   End
   Begin Project1.isButton cmdcancel 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "frm_edit_default.frx":0006
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
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "frm_edit_default.frx":170A0
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Default Excess Settings Of Inpart Water Development Corporation (IWADCO)"
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
      TabIndex        =   7
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frm_edit_default.frx":2E13A
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "FRM_EDIT_DEFAULT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
Unload Me
Set FRM_EDIT_DEFAULT = Nothing
End Sub

Private Sub cmdsave_Click()
rs(1) = txttype.Text
rs(2) = txtVALUE.Text
rs.Update
MsgBox "Changes Has Been Succesfully Made !", vbInformation, Me.Caption
Unload Me
sql = "SELECT type as 'Default Type', _value as 'Value' FROM iwadco_default"
lstview.lstDatabase sql, FRM_ULT_DEFAULT.LST_DEFAULT_SETTINGS, 1
End Sub

Private Sub Form_Load()
rsCheck
rs.Open "SELECT * FROM iwadco_default WHERE type ='" & conName & "'"
txttype.Text = rs(1)
txtVALUE.Text = rs(2)
End Sub
