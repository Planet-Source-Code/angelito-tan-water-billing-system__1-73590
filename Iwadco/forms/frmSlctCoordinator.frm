VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMSLCTCOORDINATOR 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Coordinator's Name"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6960
   Icon            =   "frmSlctCoordinator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   120
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
            Picture         =   "frmSlctCoordinator.frx":617A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSlctCoordinator.frx":6714
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lstcoor 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
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
   Begin Project1.isButton cmdclose 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "frmSlctCoordinator.frx":C89E
      Style           =   6
      Caption         =   "&Close"
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
   Begin Project1.isButton cmdselect 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "frmSlctCoordinator.frx":23938
      Style           =   6
      Caption         =   "&Select"
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
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmSlctCoordinator.frx":23C52
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Select A Coresponding Coordinator Name"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Coordinator's Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmSlctCoordinator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdclose_Click()
Unload Me
Set frmSlctCoordinator = Nothing
End Sub

Private Sub cmdselect_Click()
If Lstcoor.ListItems.Count >= 1 Then
    CoorID = Lstcoor.SelectedItem.Text
    frmAddArea.txtname.Text = Lstcoor.SelectedItem.SubItems(1)
    Unload Me
End If
End Sub

Private Sub Form_Load()
lstview.lstDatabase "SELECT id as 'ID Number',lname+', '+fname+' '+mname as Name FROM iwadco_coor ORDER BY id ASC", Lstcoor, 2
End Sub

Private Sub lstCoor_DblClick()
cmdselect_Click
End Sub
