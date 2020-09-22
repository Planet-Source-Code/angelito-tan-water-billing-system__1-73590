VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_EDITPAYMENTS 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Payments"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4650
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Adjust Customer Payments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2700
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin MSComCtl2.DTPicker txtDOP 
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20643841
         CurrentDate     =   39090
      End
      Begin VB.TextBox txtremarks 
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
         Height          =   795
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtchange 
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
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtinvoice 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtAP 
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
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Remarks :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Date Of Payment :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Change :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Invoice Number :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   405
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Amount Paid :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   795
         Width           =   1335
      End
   End
   Begin Project1.isButton cmdExit 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2900
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      icon            =   "frm_EDITPAYMENTS.frx":0000
      style           =   5
      caption         =   "&Close"
      iconsize        =   18
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frm_EDITPAYMENTS.frx":1709C
   End
   Begin Project1.isButton cmsave 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2900
      Width           =   1215
      _extentx        =   2355
      _extenty        =   661
      icon            =   "frm_EDITPAYMENTS.frx":170C4
      style           =   5
      caption         =   "&Save"
      iconsize        =   17
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frm_EDITPAYMENTS.frx":1D928
   End
End
Attribute VB_Name = "frm_EDITPAYMENTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
Set frm_EDITPAYMENTS = Nothing
End Sub

Private Sub cmsave_Click()
Dim sumof As Currency
rsCheck
rs.Open "SELECT * FROM iwadco_payments WHERE invoice = '" & txtinvoice.Text & "'"
rs("amountpayed") = txtAP.Text
rs("change") = Val(txtchange.Text)
rs("dateofpayment") = Format(txtDOP.Value, "m/dd/yyyy")
rs("remarks") = txtremarks.Text
rs("lastedit") = Format(Now, "m/dd/yyyy")
rs("lasteditby") = empID
rs.Update
rsCheck
rs.Open "SELECT SUM(amountpayed) FROM iwadco_payments WHERE id = " & frmViewRecords.RecList.SelectedItem.SubItems(1)
sumof = rs(0)
rsCheck
rs.Open "SELECT amountpaid FROM iwadco_readings WHERE id = " & frmViewRecords.RecList.SelectedItem.SubItems(1)
rs(0) = sumof
rs.Update
'refresh frm_EDITREADING FORM
sql = "SELECT invoice as 'Invoice Number',amountpayed as 'Amount Paid',change as 'Amount Change',dateofpayment as 'Date Of Payment',remarks as 'Remarks' FROM iwadco_payments WHERE id ='" & frmViewRecords.RecList.SelectedItem.SubItems(1) & "'"
lstview.lstDatabase sql, frmEDITREADING.lstpayments, 1
Unload Me
Set frm_EDITPAYMENTS = Nothing
End Sub


Private Sub txtAP_Change()
txtAP.Text = str_Filter(txtAP, 48, 57, 46)
End Sub

Private Sub txtchange_Change()
txtchange.Text = str_Filter(txtchange, 48, 57, 46)
End Sub
