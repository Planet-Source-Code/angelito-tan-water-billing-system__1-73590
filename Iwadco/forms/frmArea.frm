VERSION 5.00
Begin VB.Form frmaddarea 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Area Settings"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6375
      Begin VB.TextBox txtAreaNum 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
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
         Height          =   360
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   770
         Width           =   2895
      End
      Begin VB.TextBox txtareaname 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   5
         Top             =   1150
         Width           =   2895
      End
      Begin Project1.isButton cmdView 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "frmArea.frx":0000
         Style           =   5
         Caption         =   "View"
         IconSize        =   17
         IconAlign       =   1
         CaptionAlign    =   2
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.isButton cmdopen 
         Height          =   360
         Left            =   5280
         TabIndex        =   3
         Top             =   770
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   635
         Icon            =   "frmArea.frx":618A
         Style           =   1
         Caption         =   "open"
         IconSize        =   17
         IconAlign       =   1
         CaptionAlign    =   2
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.isButton cmdsave 
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "frmArea.frx":C9EC
         Style           =   5
         Caption         =   "&Save"
         IconSize        =   15
         IconAlign       =   1
         CaptionAlign    =   2
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.isButton cmdclose 
         Height          =   375
         Left            =   4680
         TabIndex        =   12
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "frmArea.frx":1324E
         Style           =   5
         Caption         =   "&Close"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   2
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Area Number :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   345
         TabIndex        =   10
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Coordinator's Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   350
         TabIndex        =   9
         Top             =   840
         Width           =   1710
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Area Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   315
         TabIndex        =   8
         Top             =   1200
         Width           =   1020
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmArea.frx":2A2E8
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Allow to add new area name with a corresponding coordinator name"
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
      Left            =   840
      TabIndex        =   11
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Area Settings"
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
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmAddArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
Set frmArea = Nothing
End Sub

Private Sub cmdopen_Click()
frmSlctCoordinator.Show 1
End Sub

Private Sub cmdsave_Click()
Dim AddNew As Variant
Dim errMsg As String, noErr As String
rsCheck
rs.Open "SELECT * FROM iwadco_area"
errMsg = "Please complete the following requirements!" & Chr(10) & "-----------------------------"
noErr = errMsg
If txtname.Text = "" Then
    errMsg = errMsg & Chr(10) & "Coordinator's Name"
    MsgBox errMsg, vbExclamation, Me.Caption
    Exit Sub
ElseIf txtareaname.Text = "" Then
    errMsg = errMsg & Chr(10) & "Area Name"
    MsgBox errMsg, vbExclamation, Me.Caption
    Exit Sub
Else
    If isEDIT <> True Then
        rs.AddNew
        rs.Fields("coor_id") = CoorID
        rs.Fields("area") = txtareaname.Text
        rs.Fields("createdate") = Format(Now, "mm/dd/yyyy")
        rs.Fields("createdby") = empID
        rs.Update
        MsgBox "New area has been succesfully added", vbInformation, "Add Area"
        AddNew = MsgBox("Do want to add another record?", vbInformation + vbYesNo, "Add Area")
            If AddNew = vbYes Then
                rs.MoveLast
                txtAreaNum.Text = rs("id") + 1
                txtname.Text = ""
                txtareaname.Text = ""
                Else
                Unload Me
                Set frmAddArea = Nothing
            End If
     Else
        rsCheck
        rs.Open "SELECT * FROM iwadco_area WHERE id ='" & txtAreaNum.Text & "'"
        If rs.RecordCount > 0 Then
        rs.Fields("coor_id") = CoorID
        rs.Fields("area") = txtareaname.Text
        rs.Fields("lasteditby") = empID
        rs.Fields("lastedit") = Format(Now, "mm/dd/yyyy")
        rs.Update
        MsgBox "Area Records has been updated", vbInformation, Me.Caption
        Unload Me
        Set frmAddArea = Nothing
        Else
        MsgBox "Area Record Not Exist", vbCritical, Me.Caption
        End If
    End If
End If
'set edit to false after saving
isEDIT = False
End Sub

Private Sub cmdview_Click()
frmslctArea.Caption = "Area's Record"
frmslctArea.Show 1
End Sub

Private Sub Form_Activate()
txtname.SetFocus
End Sub

Private Sub Form_Load()
cmdview.IconAlign = isbLeft
cmdview.CaptionAlign = isbright
cmdsave.IconAlign = isbLeft
cmdsave.CaptionAlign = isbright
cmdclose.IconAlign = isbLeft
cmdclose.CaptionAlign = isbright
cmdopen.IconAlign = isbLeft
cmdopen.CaptionAlign = isbright
rsCheck
rs.Open "SELECT * FROM iwadco_area"
rs.MoveLast
txtAreaNum.Text = rs.Fields("id") + 1
End Sub

