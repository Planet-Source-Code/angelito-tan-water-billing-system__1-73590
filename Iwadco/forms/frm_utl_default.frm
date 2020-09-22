VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_ult_default 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Default Settings"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10755
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   10575
      Begin MSComctlLib.ListView LST_DEFAULT_SETTINGS 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4683
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin Project1.isButton cmdCancel 
      Height          =   420
      Left            =   1920
      TabIndex        =   0
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frm_utl_default.frx":0000
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
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frm_utl_default.frx":1709A
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
   Begin MSComctlLib.ImageList imglst 
      Left            =   7080
      Top             =   240
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
            Picture         =   "frm_utl_default.frx":2C20C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_utl_default.frx":32396
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default Configuration Settings"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   0
      Width           =   8295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Excess Settings"
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
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   360
      Picture         =   "frm_utl_default.frx":38520
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "FRM_ULT_DEFAULT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
Unload Me
Set FRM_ULT_DEFAULT = Nothing
End Sub

Private Sub cmdedit_Click()
If conName <> "" Then
    FRM_EDIT_DEFAULT.Show 1
End If
End Sub

Private Sub Form_Load()
'Call dummy_db_connect
'rsCheck
With LST_DEFAULT_SETTINGS
    .Icons = imglst
    .SmallIcons = imglst
End With
sql = "SELECT type as 'Default Type', _value as 'Value' FROM iwadco_default"
lstview.lstDatabase sql, LST_DEFAULT_SETTINGS, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
conName = ""
End Sub



Private Sub LST_DEFAULT_SETTINGS_Click()
conName = LST_DEFAULT_SETTINGS.SelectedItem.Text
End Sub

Private Sub LST_DEFAULT_SETTINGS_DblClick()
Call cmdedit_Click
End Sub

