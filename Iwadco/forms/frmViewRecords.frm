VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRMViewrecords 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record List"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   11640
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   8640
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
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
            Picture         =   "frmViewRecords.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewRecords.frx":059A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search by Account Number"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   11415
      Begin Project1.chameleonButton chameleonSearch 
         Height          =   375
         Left            =   6000
         TabIndex        =   5
         Top             =   360
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         btype           =   3
         tx              =   "&Search"
         enab            =   -1  'True
         font            =   "frmViewRecords.frx":6724
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   14215660
         fcol            =   0
      End
      Begin VB.TextBox txtsearch 
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
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Consumer Account Number :"
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
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
   End
   Begin MSComctlLib.ListView RecList 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10186
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
   Begin Project1.isButton cmdDelete 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   8040
      Width           =   1455
      _extentx        =   2566
      _extenty        =   1296
      icon            =   "frmViewRecords.frx":6750
      style           =   6
      caption         =   "&Remove"
      iconsize        =   26
      captionalign    =   4
      iconalign       =   3
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frmViewRecords.frx":702C
   End
   Begin Project1.isButton cmdClose 
      Height          =   735
      Left            =   3120
      TabIndex        =   8
      Top             =   8040
      Width           =   1575
      _extentx        =   2778
      _extenty        =   1296
      icon            =   "frmViewRecords.frx":7054
      style           =   6
      caption         =   "&Close"
      iconsize        =   26
      captionalign    =   4
      iconalign       =   3
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frmViewRecords.frx":1E0F0
   End
   Begin Project1.isButton cmdedit 
      Height          =   735
      Left            =   1680
      TabIndex        =   9
      Top             =   8040
      Width           =   1335
      _extentx        =   2355
      _extenty        =   1296
      icon            =   "frmViewRecords.frx":1E118
      style           =   6
      caption         =   "&Edit"
      iconsize        =   26
      iconalign       =   3
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frmViewRecords.frx":3328C
   End
   Begin VB.Label Label3 
      Caption         =   "You can search and view consumers account information"
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
      TabIndex        =   6
      Top             =   720
      Width           =   6495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consumers Billing Records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   585
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   6390
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmViewRecords.frx":332B4
      Top             =   210
      Width           =   720
   End
End
Attribute VB_Name = "frmViewRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chameleonSearch_Click()
sql = "SELECT Account_no as 'Account Number',id as 'Record ID', billto as 'Billing Date',due_date as 'Due Date',previous_reading as 'Previous Reading',present_reading as 'Present Readings',consume as 'Consume' ,excess as 'Excess',amount_excess as 'Amount Excess',arrears as 'Arrears',total_amount as 'Total Amount',amountpaid as 'Amount Paid' ,iwadco_db.dbo.checkifNull((SELECT     SUM(change) From iwadco_payments WHERE      iwadco_payments.ConID = a.account_no AND iwadco_payments.id = a.id )) as Change,status FROM iwadco_readings a WHERE account_no = '" & txtsearch.Text & "' AND deletedby=0 ORDER BY id "
Debug.Print sql
lstview.lstDatabase sql, RecList, 1
End Sub

Private Sub cmdclose_Click()
Unload Me
Set frmViewRecords = Nothing
End Sub

Private Sub cmddelete_Click()
If RecList.ListItems.Count <= 0 Then Exit Sub
If user_priv("delete") = False Then MsgBox "You are not allowed to delete transaction!", vbExclamation, Me.Caption: Exit Sub
If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
    rsCheck
    sql = "SELECT * FROM iwadco_payments WHERE id = '" & RecList.SelectedItem.SubItems(1) & "'"
    Debug.Print sql
    rs.Open sql, CN, adOpenStatic, adLockOptimistic
    If rs.RecordCount <> 0 Then
    rs("datedeleted") = Format(Now, "mm/dd/yyyy")
    rs("deletedby") = empID
    rs.Update
    End If
    rsCheck
    rs.Open "SELECT * FROM iwadco_readings WHERE id = '" & RecList.SelectedItem.SubItems(1) & "'", CN, adOpenStatic, adLockOptimistic
    rs("datedeleted") = Format(Now, "mm/dd/yyyy")
    rs("deletedby") = empID
    rs.Update
    MsgBox "Record is deleted!", vbInformation, Me.Caption
End If
sql = "SELECT Account_no as 'Account Number',id as 'Record ID', billto as 'Billing Date',due_date as 'Due Date',previous_reading as 'Previous Reading',present_reading as 'Present Readings',consume as 'Consume' ,excess as 'Excess',amount_excess as 'Amount Excess',arrears as 'Arrears',total_amount as 'Total Amount',amountpaid as 'Amount Paid' ,iwadco_db.dbo.checkifNull((SELECT     SUM(change) From iwadco_payments WHERE      iwadco_payments.ConID = a.account_no AND iwadco_payments.id = a.id )) as Change FROM iwadco_readings a WHERE account_no = '" & txtsearch.Text & "' AND deletedby=0 ORDER BY id "
lstview.lstDatabase sql, RecList, 1
End Sub

Private Sub cmdedit_Click()
If user_priv("update") = False Then MsgBox "You are not allowed to edit transaction!", vbExclamation, Me.Caption: Exit Sub
If RecList.ListItems.Count > 0 Then
    frmEDITREADING.txtarrears.Text = RecList.SelectedItem.SubItems(9)
    frmEDITREADING.txtprevious.Text = RecList.SelectedItem.SubItems(4)
    frmEDITREADING.txtpresent.Text = RecList.SelectedItem.SubItems(5)
    frmEDITREADING.Show 1
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
'PopupMenu
If CONSUMERID <> "" Then
    txtsearch.Text = CONSUMERID
    chameleonSearch_Click
    txtsearch.SetFocus
   ' SendKeys "{end}"
End If

formBoolean = True
cmdDelete.IconAlign = isbTop
cmdDelete.CaptionAlign = isbBottom
cmdClose.IconAlign = isbTop
cmdClose.CaptionAlign = isbBottom
cmdedit.IconAlign = isbTop
cmdedit.CaptionAlign = isbBottom
End Sub

Private Sub Form_Load()                                                                                                                                                                                                                                                                                       'PUBLIC VAR SEE FRMREADING
sql = "SELECT Account_no as 'Account Number',id as 'Record ID', billto as 'Billing Date',due_date as 'Due Date',previous_reading as 'Previous Reading',present_reading as 'Present Readings',consume as 'Consume' ,excess as 'Excess',amount_excess as 'Amount Excess',arrears as 'Arrears',total_amount as 'Total Amount',amountpaid as 'Amount Paid' FROM iwadco_readings WHERE account_no = '" & CONSUMERID & "' AND deletedby=0 ORDER BY id "
lstview.lstDatabase sql, RecList, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
formBoolean = False
End Sub

Private Sub RecList_DblClick()
Call cmdedit_Click
End Sub

Private Sub Timer1_Timer()
txtsearch.Text = CONSUMERID
End Sub

Private Sub txtSearch_Change()
txtsearch.Text = str_Filter(txtsearch, 48, 57, 45)
'set
CONSUMERID = txtsearch.Text
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
chameleonSearch_Click
ElseIf KeyAscii = 27 Then
Unload Me
Set frmViewRecords = Nothing
End If
End Sub
Public Sub dblclick()
chameleonSearch_Click
End Sub
