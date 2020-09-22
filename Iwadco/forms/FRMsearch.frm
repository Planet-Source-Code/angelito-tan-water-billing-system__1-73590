VERSION 5.00
Begin VB.Form FRMsearch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7575
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.isButton cmdsearch 
      Height          =   390
      Left            =   600
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
      Icon            =   "FRMsearch.frx":0000
      Style           =   5
      Caption         =   "&Search"
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
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7335
      Begin VB.TextBox txtSN 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   480
         TabIndex        =   5
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtAN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   4
         Top             =   240
         Width           =   3375
      End
      Begin VB.OptionButton SearchByName 
         Caption         =   "Search By  First Name and Last Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton searchbyAC 
         Caption         =   "Search By Account Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   4200
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   2895
      End
   End
   Begin Project1.isButton cmdclose 
      Height          =   390
      Left            =   2160
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
      Icon            =   "FRMsearch.frx":001C
      Style           =   5
      Caption         =   "&Close"
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
      Picture         =   "FRMsearch.frx":0038
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Consumer Account Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "FRMsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
Set FRMsearch = Nothing
End Sub

Private Sub cmdSearch_Click()
If SelectedForm = UCase("frmConsumer") Then
'consumer search
    If searchbyAC.Value = True Then
        If Len(txtAN.Text) > 0 Then                                                                                                                                                                                                                                                                                                                                                                                                                                                     'iwadco_cons.coor_id ORDER BY iwadco_cons.id ASC"
            sql = "SELECT iwadco_cons.id as 'Account No',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name',iwadco_cons.address as 'Address',iwadco_cons.tel as 'Phone',iwadco_cons.mobile as 'Mobile',iwadco_cons.email as 'Email',iwadco_cons.dateregistered as 'Date Registered',iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname as 'Coordinator Name' from iwadco_cons,iwadco_coor WHERE iwadco_cons.id = '" & txtAN.Text & "'AND iwadco_cons.status ='E' AND iwadco_cons.coor_id = iwadco_coor.id"
            lstview.lstDatabase sql, frmConsumer.lstCustomer, 2
            Else
            MsgBox "Account Number Empty", vbExclamation, Me.Caption
        End If
    Else
        If Len(txtSN.Text) > 0 Then
            'SEARCH BY FNAME
                sql = "SELECT iwadco_cons.id as 'Account No',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name',iwadco_cons.address as 'Address',iwadco_cons.tel as 'Phone',iwadco_cons.mobile as 'Mobile',iwadco_cons.email as 'Email',iwadco_cons.dateregistered as 'Date Registered',iwadco_coor.fname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname as 'Coordinator Name' from iwadco_cons,iwadco_coor WHERE  iwadco_cons.fname like '%" & txtSN.Text & "%'AND iwadco_cons.coor_id = iwadco_coor.id AND iwadco_cons.status ='E'"
                rsCheck
                rs.Open sql, CN, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then
                lstview.lstDatabase sql, frmConsumer.lstCustomer, 2
                'SEARCH BY LNAME
                Else
                sql = "SELECT iwadco_cons.id as 'Account No',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name',iwadco_cons.address as 'Address',iwadco_cons.tel as 'Phone',iwadco_cons.mobile as 'Mobile',iwadco_cons.email as 'Email',iwadco_cons.dateregistered as 'Date Registered',iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname as 'Coordinator Name' from iwadco_cons,iwadco_coor WHERE  iwadco_cons.lname like '%" & txtSN.Text & "%'AND iwadco_cons.coor_id = iwadco_coor.id AND iwadco_cons.status ='E'"
                lstview.lstDatabase sql, frmConsumer.lstCustomer, 2
            End If
        Else
            MsgBox "Name is Empty", vbExclamation, Me.Caption
        End If
    End If
Else
    'coordinator search
    If searchbyAC.Value = True Then
        If Len(txtAN.Text) > 0 Then                                                                                                                                                                                                                                                                                                                                                                                                                                                     'iwadco_cons.coor_id ORDER BY iwadco_cons.id ASC"
            sql = "SELECT id as 'Account No',lname+', '+fname+' '+mname as Name,address as 'Address',tel as 'Phone',mobile as 'Mobile',email as 'Email',datehired as 'Date Registered',billingdate as 'Billing Date' FROM iwadco_coor WHERE iwadco_coor.id =" & txtAN.Text & " AND iwadco_coor.status ='E'"
            lstview.lstDatabase sql, frmCoordinator.Lstcoor, 2
            Else
            MsgBox "Account Number Empty", vbExclamation, Me.Caption
        End If
    Else
        If Len(txtSN.Text) > 0 Then
            sql = "SELECT id as 'Account No',lname+', '+fname+' '+mname as Name,address as 'Address',tel as 'Phone',mobile as 'Mobile',email as 'Email',datehired as 'Date Registered',billingdate as 'Billing Date' FROM iwadco_coor WHERE iwadco_coor.fname LIKE '" & txtSN.Text & "%'AND iwadco_coor.status ='E'"
            rsCheck
            rs.Open sql, CN, adOpenKeyset, adLockOptimistic
            'search for fname
            If rs.RecordCount > 0 Then
                lstview.lstDatabase sql, frmCoordinator.Lstcoor, 2
                Else
            'search for lname
                sql = "SELECT id as 'Account No',lname+', '+fname+' '+mname as Name,address as 'Address',tel as 'Phone',mobile as 'Mobile',email as 'Email',datehired as 'Date Registered',billingdate as 'Billing Date' FROM iwadco_coor WHERE iwadco_coor.lname LIKE '" & txtSN.Text & "%'AND iwadco_coor.status ='E'"
                lstview.lstDatabase sql, frmCoordinator.Lstcoor, 2
            End If
            Else
            MsgBox "Name is Empty", vbExclamation, Me.Caption
        End If
    End If
End If
        
End Sub

Private Sub Form_Activate()
txtAN.SetFocus
End Sub

Private Sub searchbyAC_Click()
txtAN.Text = ""
txtAN.BackColor = &H80000018
txtAN.SetFocus
'---
txtSN.Text = ""
txtSN.BackColor = &HFFFFFF
End Sub

Private Sub SearchByName_Click()
txtSN.Text = ""
txtSN.SetFocus
txtSN.BackColor = &H80000018
'---
txtAN.Text = ""
txtAN.BackColor = &HFFFFFF
End Sub

Private Sub txtAN_Change()
txtAN.Text = str_Filter(txtAN, 48, 57, 45)
End Sub

Private Sub txtSN_Change()
txtSN.Text = str_Filter(txtSN, 65, 122, 32)
End Sub
