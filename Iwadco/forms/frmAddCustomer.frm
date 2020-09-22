VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmaddconsumer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consumer Settings"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9945
   Icon            =   "frmAddCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   9735
      Begin VB.TextBox txtMonthly 
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
         Left            =   6360
         TabIndex        =   31
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox txtLname 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1440
         TabIndex        =   1
         Top             =   770
         Width           =   2925
      End
      Begin VB.TextBox txtacctno 
         BackColor       =   &H00E6FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   1620
      End
      Begin VB.TextBox txtfname 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1440
         TabIndex        =   2
         Top             =   1150
         Width           =   2925
      End
      Begin VB.TextBox txtmname 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1440
         TabIndex        =   3
         Top             =   1550
         Width           =   2925
      End
      Begin VB.TextBox txtadd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1950
         Width           =   2925
      End
      Begin VB.TextBox txttel 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6360
         TabIndex        =   5
         Top             =   360
         Width           =   2925
      End
      Begin VB.TextBox txtmobile 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6360
         TabIndex        =   6
         Top             =   760
         Width           =   2925
      End
      Begin VB.TextBox txtemail 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6360
         TabIndex        =   7
         Top             =   1150
         Width           =   2925
      End
      Begin VB.TextBox txtAreaNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1950
         Width           =   2325
      End
      Begin VB.TextBox txtCorName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2350
         Width           =   2325
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2745
         Width           =   2295
      End
      Begin Project1.isButton cmdopen 
         Height          =   360
         Left            =   8760
         TabIndex        =   14
         Top             =   1950
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   635
         Icon            =   "frmAddCustomer.frx":08CA
         Style           =   1
         Caption         =   "open"
         IconSize        =   17
         IconAlign       =   1
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
      Begin Project1.isButton isButton1 
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   3120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "frmAddCustomer.frx":15A3C
         Style           =   5
         Caption         =   "&Cancel"
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.isButton cmdsave 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "frmAddCustomer.frx":2CAD6
         Style           =   5
         Caption         =   "&Save"
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   6360
         TabIndex        =   8
         Top             =   1560
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         Format          =   50069505
         CurrentDate     =   39380
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Bill"
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
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Customer ID : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   150
         TabIndex        =   29
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Last Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   135
         TabIndex        =   28
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "First Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   150
         TabIndex        =   27
         Top             =   1320
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Middle Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   150
         TabIndex        =   25
         Top             =   2280
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tel # :"
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
         Left            =   4590
         TabIndex        =   24
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mobile # :"
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
         Left            =   4590
         TabIndex        =   23
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Email :"
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
         Left            =   4590
         TabIndex        =   22
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date Registered :"
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
         Left            =   4590
         TabIndex        =   21
         Top             =   1635
         Width           =   1440
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Coordinator Name :"
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
         Left            =   4590
         TabIndex        =   20
         Top             =   2400
         Width           =   1590
      End
      Begin VB.Label Label11 
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
         Left            =   4590
         TabIndex        =   19
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Connection Type :"
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
         Height          =   255
         Left            =   4590
         TabIndex        =   18
         Top             =   2760
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   240
      Picture         =   "frmAddCustomer.frx":41C48
      Stretch         =   -1  'True
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Add or Edit Consumer Account"
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
      Left            =   1080
      TabIndex        =   15
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Consumer Settings"
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
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmAddConsumer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecordEDIT As Boolean
Dim sql As Variant
Dim frmCaption As Variant
Dim ActivateCount As Long
Dim tmpstr As String
Private Sub cmdopen_Click()
frmslctArea.Show 1
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Dim errMsg As String, noErr As String
errMsg = "Please complete the following requirements!" & Chr(10) & "-----------------------------"
noErr = errMsg
If txtLname.Text = "" Then
    errMsg = errMsg & Chr(10) & "Lastname"
End If

If txtfname.Text = "" Then
    errMsg = errMsg & Chr(10) & "Firstname"
End If

If txtmname.Text = "" Then
    errMsg = errMsg & Chr(10) & "Middlename"
End If

If txtadd.Text = "" Then
    errMsg = errMsg & Chr(10) & "Address"
End If
If CoorID = 0 Then
    errMsg = errMsg & Chr(10) & "Coordinator"
End If
If AreaID = 0 Then
    errMsg = errMsg & Chr(10) & "Area"
End If

If Not IsNumeric(txtMonthly.Text) And Combo1.Text = "Flat Rate" Then
    errMsg = errMsg & Chr(10) & "Invalid Monthly Bill!"
End If
    
If Combo1.Text = "" Then
    errMsg = errMsg & Chr(10) & "Connection Type"
End If

If errMsg <> noErr Then
    MsgBox errMsg, vbExclamation, Me.Caption
    Exit Sub
End If

If RecordEDIT = False Then
    rsCheck
    rs.Open "SELECT * FROM iwadco_cons"
    rs.AddNew
    rs(11) = Format(Now, "mm/dd/yyyy hh:mm:ss")
    rs(15) = empID
Else
'if consumer = "" then
    rsCheck
    If Len(CONSUMERID) <= 0 Then
        rs.Open "SELECT * FROM iwadco_cons WHERE id ='" & frmConsumer.lstCustomer.SelectedItem.Text & "'"
    Else
         rs.Open "SELECT * FROM iwadco_cons WHERE id ='" & CONSUMERID & "'"
    End If
    rs(9) = Format(Now, "mm/dd/yyyy hh:mm:ss")
    rs(13) = empID
End If
    rs(0) = txtacctno.Text
    rs(1) = CoorID
    rs(2) = AreaID
If txtMonthly.Text <> "" Then
    rs("monthly") = txtMonthly.Text
End If
rs(3) = Combo1.ListIndex + 1
rs(4) = txtLname.Text
rs(5) = txtfname.Text
rs(6) = txtmname.Text
rs(7) = txtadd.Text
rs(8) = txttel.Text
rs(9) = txtMobile.Text
rs(10) = txtemail.Text
rs(11) = DTPicker1.Value

'MsgBox AreaID

'MsgBox CoorID
'Exit Sub
rs.Update
rsCheck
rs.Open "SELECT account_no FROM iwadco_readings WHERE account_no = '" & frmConsumer.lstCustomer.SelectedItem.Text & "'", CN, adOpenStatic, adLockOptimistic
rs(0) = txtacctno.Text
rs.Update
rsCheck
rs.Open "SELECT account_no FROM iwadco_commisions WHERE account_no = '" & frmConsumer.lstCustomer.SelectedItem.Text & "'", CN, adOpenStatic, adLockOptimistic
rs(0) = txtacctno.Text
rs.Update

CoorID = 0
If RecordEDIT = False Then
MsgBox "New record has succesfully been saved", vbInformation, Me.Caption
    If MsgBox("Do you want to add new record", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        txtacctno.Text = ""
        txtLname.Text = ""
        txtfname.Text = ""
        txtmname.Text = ""
        txtadd.Text = ""
        txttel.Text = ""
        txtMobile.Text = ""
        txtemail.Text = ""
        frmslctArea.Show 1
    Else
        Unload Me
        Set frmAddConsumer = Nothing
    End If
Else
    MsgBox "Record has been succesfully updated !", vbInformation, Me.Caption
    Unload Me
    Set frmAddConsumer = Nothing
End If
'verify

If Len(CONSUMERID) <= 0 Then
    sql = "SELECT iwadco_cons.id as 'Account No',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name',iwadco_cons.address as 'Address',iwadco_cons.tel as 'Phone',iwadco_cons.mobile as 'Mobile',iwadco_cons.email as 'Email',iwadco_cons.dateregistered as 'Date Registered',iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname as 'Coordinator Name' from iwadco_cons,iwadco_coor WHERE iwadco_coor.id = iwadco_cons.coor_id ORDER BY iwadco_cons.id ASC"
    lstview.lstDatabase sql, frmConsumer.lstCustomer, 2
End If
End Sub

Private Sub Combo1_Click()
If Combo1.Text <> "Flat Rate" Then
    txtMonthly.Enabled = False
Else
    txtMonthly.Enabled = True
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
ActivateCount = ActivateCount + 1
If ActivateCount = 1 Then
Combo1.SetFocus
rsCheck
rs.Open "SELECT * FROM iwadco_cons"
If rs.RecordCount <= 0 Then
rsCheck
Call LoadCombo("SELECT type FROM iwadco_typcon ORDER BY id ASC", Combo1)
Exit Sub
End If
Call LoadCombo("SELECT type FROM iwadco_typcon ORDER BY id ASC", Combo1)
    If Right(Me.Caption, 3) = "Add" Then
        rsCheck
        'txtacctno.Text = ""
        txtLname.Text = ""
        txtfname.Text = ""
        txtmname.Text = ""
        txtadd.Text = ""
        txttel.Text = ""
        txtMobile.Text = ""
        txtemail.Text = ""
        RecordEDIT = False
        DTPicker1.Value = Format(Now, "mm/dd/yyyy")
    Else
        If Len(CONSUMERID) <= 0 Then
            sql = "SELECT iwadco_cons.*,iwadco_coor.lname+', '+iwadco_coor.fname+' '+iwadco_coor.mname AS coorName, iwadco_area.area FROM iwadco_cons,iwadco_coor,iwadco_area WHERE iwadco_cons.coor_id = iwadco_coor.id AND area_id = iwadco_area.id AND iwadco_cons.id = '" & frmConsumer.lstCustomer.SelectedItem.Text & "'"
            Else
            sql = "SELECT iwadco_cons.*,iwadco_coor.lname+', '+iwadco_coor.fname+' '+iwadco_coor.mname AS coorName, iwadco_area.area FROM iwadco_cons,iwadco_coor,iwadco_area WHERE iwadco_cons.coor_id = iwadco_coor.id AND area_id = iwadco_area.id AND iwadco_cons.id = '" & CONSUMERID & "'"
        End If
        rsCheck
        rs.Open sql
        If Len(z) > 0 Then
        txtacctno.Text = z
        Else
        txtacctno.Text = rs(0)
        End If
        txtLname.Text = rs(4)
        txtfname.Text = rs(5)
        txtmname.Text = rs(6)
        txtadd.Text = rs(7)
        txttel.Text = rs(8)
        txtMobile.Text = rs(9)
        txtemail.Text = rs(10)
        DTPicker1.Value = rs(11).Value
        txtAreaNo.Text = rs("area")
        txtCorName.Text = rs("coorName")
        
        If Len(rs("monthly")) > 0 Then
        txtMonthly.Text = rs("monthly")
        End If
       ' coorID = rs(1)
       ' AreaID = rs(2)
        Call selectCombo(rs("class"), Combo1)
        RecordEDIT = True
    End If
If txtacctno.Text = "" Then
frmslctArea.Show 1
End If
End If
rsCheck
rs.Open "SELECT * FROM iwadco_cons WHERE id ='" & txtacctno.Text & "'"
If rs.RecordCount > 0 Then
    AreaID = rs("area_id")
    CoorID = rs("coor_id")
End If
End Sub

Private Sub Form_Load()
'MsgBox txtacctno.Text
'Initialize activecount = 0
ActivateCount = 0
cmdopen.IconAlign = isbLeft
cmdopen.CaptionAlign = isbright
isButton1.IconAlign = isbLeft
isButton1.CaptionAlign = isbright
cmdsave.IconAlign = isbLeft
cmdsave.CaptionAlign = isbright

End Sub

Private Sub Form_Unload(Cancel As Integer)
RecordEDIT = False
'produce error dont know y maybe bcoz of active form is differ from form_load
'CONSUMERID = ""
End Sub

Private Sub isButton1_Click()
Unload Me
Set frmAddConsumer = Nothing
z = ""
txtacctno.Text = ""
'good
CONSUMERID = ""
End Sub
