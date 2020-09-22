VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRMSLCTAREA 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select a coordinator name"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmslctArea.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.isButton cmddelete 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "frmslctArea.frx":617A
      Style           =   5
      Caption         =   "&Delete"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.isButton cmdedit 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "frmslctArea.frx":1B2EC
      Style           =   5
      Caption         =   "&Edit"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   120
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
            Picture         =   "frmslctArea.frx":3045E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Project1.isButton cmdclose 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "frmslctArea.frx":309F8
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
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "frmslctArea.frx":47A92
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
   Begin MSComctlLib.ListView lst 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5318
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
      ForeColor       =   -2147483640
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Area below"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmslctArea.frx":47DAC
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Coordinator Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmslctArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
Unload Me
Set frmslctArea = Nothing
End Sub

Private Sub cmddelete_Click()
rsCheck
rs.Open "Select * from iwadco_cons WHERE area_id ='" & lst.SelectedItem.Text & "'"
If rs.RecordCount > 0 Then
    MsgBox "Area Record Contain Consumer Area! Cannot delete this record", vbExclamation, Me.Caption
    Else
        If MsgBox("Are you sure you want to delete this record ?", vbInformation + vbYesNo) = vbYes Then
        rsCheck
        rs.Open "SELECT * FROM iwadco_area WHERE id ='" & lst.SelectedItem.Text & "'"
        rs.Delete
        MsgBox "Area Record Delete Succes", vbInformation, Me.Caption
        Unload Me
        Set frmslctArea = Nothing
        End If
End If
End Sub

Private Sub cmdedit_Click()
If lst.ListItems.Count > 0 Then
    If lst.SelectedItem.Text > 0 Then
        isEDIT = True
        frmAddArea.txtAreaNum.Text = lst.SelectedItem.Text
         frmAddArea.txtareaname.Text = lst.SelectedItem.SubItems(1)
        frmAddArea.txtname.Text = lst.SelectedItem.SubItems(2)
        Unload Me
    End If
End If
End Sub

Private Sub cmdselect_Click()
Dim XX As Integer
Dim TmpArea As String
Dim tmpcorID As String

If lst.ListItems.Count >= 1 Then
    AreaID = frmslctArea.lst.SelectedItem.Text
    CoorID = frmslctArea.lst.SelectedItem.SubItems(3)
    frmAddConsumer.txtCorName.Text = frmslctArea.lst.SelectedItem.SubItems(2)
    frmAddConsumer.txtAreaNo.Text = frmslctArea.lst.SelectedItem.SubItems(1)
    TmpArea = AreaID
    tmpcorID = CoorID
        If Len(TmpArea) = 1 Then
            TmpArea = "0" & AreaID
        End If
        If Len(tmpcorID) = 1 Then
            tmpcorID = "0" & CoorID
        End If
    frmAddConsumer.txtacctno.Text = Unique_ID(tmpcorID, TmpArea, CoorID, AreaID)
    z = Unique_ID(tmpcorID, TmpArea, CoorID, AreaID)
    Unload Me
End If


End Sub

Private Sub Form_Load()
sql = "SELECT iwadco_area.id as 'ID', area as 'Area Name', lname+' '+fname+' '+mname as 'Coordinator Incharge',iwadco_coor.id as 'Coordinator ID' FROM iwadco_area,iwadco_coor WHERE iwadco_area.id = iwadco_area.id AND iwadco_coor.id = iwadco_area.coor_id AND iwadco_area.status ='E' ORDER BY iwadco_area.id ASC"
lstview.lstDatabase sql, lst, 1
If Me.Caption = "Area's Record" Then
Me.Caption = "Select a coordinator name"
End If
End Sub

Private Sub lst_DblClick()
Call cmdselect_Click
End Sub

