VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMPAY 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payments"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   735
      Left            =   4800
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Format          =   49283073
      CurrentDate     =   39434
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   3240
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Format          =   49283073
      CurrentDate     =   39434
   End
   Begin Project1.isButton isButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "frmpay.frx":0000
      Style           =   5
      Caption         =   "&Print Receipt"
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
   Begin Project1.isButton cmdclose 
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "frmpay.frx":001C
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
   Begin Project1.isButton cmsave 
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "frmpay.frx":170B6
      Style           =   5
      Caption         =   "&Save"
      IconSize        =   17
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
   Begin VB.Frame Frame2 
      Caption         =   "Payments Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6135
      Begin VB.TextBox txtORNumber 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtAP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtremarks 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   4680
         Width           =   5655
      End
      Begin VB.TextBox txtCA 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox txtcontype 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   440
         Width           =   4095
      End
      Begin VB.TextBox txtCP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Frame Frame3 
         Caption         =   "Full / Partial Payments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2175
         Begin VB.OptionButton Option5 
            Caption         =   "Promisorry Note"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   910
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Full Payments"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   300
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Partial Payments"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OR Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BB5900&
         Height          =   270
         Index           =   0
         Left            =   2400
         TabIndex        =   25
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label Label5 
         BackColor       =   &H00BB5900&
         BackStyle       =   0  'Transparent
         Caption         =   "Advance Payment :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BB5900&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   240
         X2              =   5880
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   4320
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Amount Recieve :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BB5900&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3960
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PHP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   14
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   240
         X2              =   5880
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   240
         X2              =   5880
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label lblamount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Type Of Connection :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BB5900&
         Height          =   255
         Left            =   -480
         TabIndex        =   12
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00BB5900&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BB5900&
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   3480
         Width           =   1395
      End
      Begin VB.Label Label8 
         BackColor       =   &H00BB5900&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Due :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BB5900&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   2415
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payment and  Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   270
      Index           =   1
      Left            =   960
      TabIndex        =   16
      Top             =   360
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   1
      Left            =   120
      Picture         =   "frmpay.frx":1D918
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmpay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim amountCHANGE As Double
Dim advancePAYMENT As Double
Dim amountPAID As Double
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmsave_Click()
On Error Resume Next
Dim rs1 As New ADODB.Recordset
Dim ReceiptPrint As Variant
Dim inv As String
Dim amount_paids As Double
Dim X As Integer
Dim comision As Double
Dim forfeited As Double
inv = Format(Now, "myydhms")
amountCHANGE = 0
'rsCheck
'MsgBox frmpayments.lstpayments.SelectedItem.SubItems(1)
'   rs.Open "SELECT * FROM iwadco_readings WHERE account_no ='" & frmpayments.lstpayments.SelectedItem.SubItems(1) & "'AND status ='I'"
'   MsgBox rs.RecordCount
'   Exit Sub
If Option1.Value = True Then
    If Val(txtCA.Text) - Val(lblamount.Caption) < 0 Then
        MsgBox "Cash amount receive must not be lower than amount due", vbExclamation, "Error Amount Due"
        Exit Sub
    Else
        amountCHANGE = Val(txtCA.Text) - Val(lblamount.Caption)
    End If
ElseIf Option2.Value = True Then
        If Val(txtCA.Text) > Val(lblamount.Caption) Then
        MsgBox "Cash amount receive must not higher than amount due", vbInformation, "Partial Payment"
        Exit Sub
        End If
ElseIf Option5.Value = True Then
        rsCheck
        rs.Open "SELECT * FROM iwadco_readings WHERE id ='" & frmpayments.lstpayments.SelectedItem & "'"
        rs("promisorry") = "Yes"
        rs("promisorrydate") = Format(Now, "mm/dd/yyyy")
        If txtremarks.Text <> "" Then
            rs("promissorynote") = txtremarks.Text
        Else
            MsgBox "Promissory note should not be empty !", vbExclamation, "Promisorry Note"
        Exit Sub
        End If
        rs.Update
        MsgBox "New promisorry added", vbInformation, "Promisorry Note"
        sql = "SELECT id as 'ID Number',account_no as 'Account Number',billto as 'Billing Date',due_date as 'Due Date',total_amount as 'Total Amount',amountpaid as 'Amount Paid', promisorry as 'Promisorry Note' FROM iwadco_readings WHERE id = '" & frmpayments.lstpayments.SelectedItem.Text & "' AND deletedby =0 AND status = 'I'"
        lstview.lstDatabase sql, frmpayments.lstpayments, 1
        Unload Me
        Exit Sub
End If
If Val(lblamount.Caption) = 0 Then
    GoTo goo
End If
'IF NO OR NUMBER IN THE TEXTBOX IS EMPTY !
If txtORNumber.Text = "" Then
    'if TXTORNUMBER IS EMPTY MESSAGE BOX
    MsgBox "Official receipt is required !", vbExclamation, "Official Receipt ERROR !"
    Exit Sub
End If
rsCheck
rs.Open "SELECT * FROM iwadco_payments where invoice ='" & txtORNumber.Text & "'"
'LOOK FOR EXISTING OFFICIAL RECEIPT !
If rs.RecordCount > 0 Then
    MsgBox "Official receipt already exist !", vbExclamation, "Official Receipt ERROR !!"
    Exit Sub
End If
If Val(txtCP.Text) > 0 And Val(txtCA.Text) > 0 Then
    rs.AddNew
    rs("id") = frmpayments.lstpayments.SelectedItem.Text
    rs("conid") = frmpayments.lstpayments.SelectedItem.SubItems(1)
    'OR number
    rs("invoice") = txtORNumber.Text
    rs("amountpayed") = txtCA.Text
    rs("dateofpayment") = Format(Now, "mm/dd/yyyy")
    rs("remarks") = txtremarks.Text
    rs("lastedit") = Format(Now, "mm/dd/yyyy")
    rs("lasteditby") = empID
    rs("change") = amountCHANGE
            rs.Update
        '---------------
        'query another
           
    rsCheck
    rs.Open "SELECT billto,total_amount-arrears as comision,arrears as forfeited FROM iwadco_readings WHERE id= " & frmpayments.lstpayments.SelectedItem.Text & ""
    'MsgBox rs(0)
    DTPicker2.Value = DateAdd("m", 1, rs(0))
    comision = rs(1)
    forfeited = rs(2)
    If rs1.State = adStateOpen Then rs1.Close
    rs1.Open "SELECT * FROM iwadco_cons WHERE id = '" & frmMain.lstemp.SelectedItem.Text & "'", CN, adOpenStatic, adLockOptimistic
    rsCheck
    rs.Open "SELECT * FROM iwadco_commisions", CN, adOpenStatic, adLockOptimistic
    rs.AddNew
    rs(0) = rs1("coor_id")
    rs(1) = txtORNumber.Text
    rs(2) = Format(Now, "mm/dd/yyyy")
    rs(3) = frmMain.lstemp.SelectedItem.Text
    rs(4) = comision
    rs(5) = CDbl(txtCA.Text - forfeited) / 1.12 * 0.2
    rs(6) = rs(5).Value * 0.1
    rs(8) = Format(Now, "mm/dd/yyyy")
    rs(9) = empID
    rs(10) = frmpayments.lstpayments.SelectedItem.Text
    rs.Update
    If forfeited > 0 Then
        rsCheck
        rs.Open "SELECT * FROM iwadco_commisions", CN, adOpenStatic, adLockOptimistic
        rs.AddNew
        rs(0) = rs1("coor_id")
        rs(1) = txtORNumber.Text
        rs(2) = Format(Now, "mm/dd/yyyy")
        rs(3) = frmMain.lstemp.SelectedItem.Text
        rs(4) = forfeited
        rs(5) = 0
        rs(6) = 0
        rs(8) = Format(Now, "mm/dd/yyyy")
        rs(9) = empID
        rs(10) = frmpayments.lstpayments.SelectedItem.Text
        rs.Update
    End If
    amount_paids = txtCA.Text
    rsCheck
    rs.Open "SELECT * FROM iwadco_readings WHERE id< " & frmpayments.lstpayments.SelectedItem.Text & " AND account_no = '" & frmpayments.lstpayments.SelectedItem.SubItems(1) & "'"
    rs.MoveFirst
    rs.MoveLast
    rs.MoveFirst
   'rsCheck
   'rs.Open "SELECT * FROM iwadco_readings WHERE id ='" & frmpayments.lstpayments.SelectedItem.Text & "'AND status ='I'"
    For X = 1 To rs.RecordCount
        'If rs("amountpaid") + amount_paids >= rs("total_amount") Then
            'rs("amountpaid") = rs("total_amount")
            rs("status") = "C"
        'Else
        '    rs("amountpaid") = amount_paids
        'End If
        rs.MoveNext
    Next
goo:
        rsCheck
        'rs.Open "SELECT iwadco_readings.amountpaid,iwadco_readings.total_amount,iwadco_readings.status,iwadco_payments.id,iwadco_payments.amountpayed FROM iwadco_readings,iwadco_payments WHERE iwadco_readings.id ='" & frmpayments.lstpayments.SelectedItem.Text & "'AND iwadco_payments.id = iwadco_readings.id"
        rs.Open "SELECT * FROM iwadco_readings WHERE id ='" & frmpayments.lstpayments.SelectedItem & "'"
        amountPAID = Format(rs("amountpaid"), "###,##0.00")
        amountPAID = amountPAID + txtCA.Text
        'Check if the amountpay has equal or greater than the total amount to pay .
        rs("amountpaid") = Format(amountPAID, "###,##0.00")
            'If amountPAID >= rs("total_amount") Then
            'if amount paid is greater thant equal to total amount to pay then set it the status to complete
            If Option2.Value = False Then
                rs("status") = "C"
            Else
                rs("status") = "I"
            End If
            rs("promisorry") = "No"
            rs.Update
            '----enabled consumer
            rsCheck
            rs.Open "SELECT * FROM iwadco_cons WHERE id = '" & frmpayments.lstpayments.SelectedItem.SubItems(1) & "'"
            rs("status") = "E"
            'End If
        rs.Update
        
            frmpayments.Frame4.Visible = False
            frmpayments.Frame5.Visible = False
        '---------------------
        MsgBox "Transaction Complete!", vbInformation, Me.Caption
Else
    MsgBox "Cannot process payment some of fields may contain no Value or Negative value", vbExclamation, Me.Caption
    Exit Sub
End If
If MsgBox("Do you want to print Receipt", vbYesNo + vbInformation, Me.Caption) = vbYes Then
    Unload Me
    If DataEnvironment1.rscmdReceipt.State = adStateOpen Then DataEnvironment1.rscmdReceipt.Close
    DataEnvironment1.rscmdReceipt.Open "SELECT iwadco_cons.id, iwadco_cons.lname+', '+ iwadco_cons.fname+' '+iwadco_cons.mname as 'Costumer Name', iwadco_readings.billfrom, iwadco_readings.billto, iwadco_readings.due_date, iwadco_readings.previous_reading, iwadco_readings.present_reading, iwadco_readings.consume, iwadco_readings.excess, iwadco_readings.amount_excess, iwadco_readings.arrears, iwadco_readings.total_amount, iwadco_readings.wtax, iwadco_payments.invoice, iwadco_payments.amountpayed, iwadco_payments.change, iwadco_payments.dateofpayment FROM iwadco_cons INNER JOIN iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no INNER JOIN iwadco_payments ON iwadco_readings.id = iwadco_payments.id WHERE invoice = '" & inv & "'"
    rptReceipt.Show
End If
    'refresh listview
    'sql = "SELECT id as 'ID Number',account_no as 'Account Number',billto as 'Billing Date',due_date as 'Due Date',total_amount as 'Total Amount',amountpaid as 'Amount Paid', promisorry as 'Promisorry Note' FROM iwadco_readings WHERE account_no = '" & txtacctno.Text & "' AND deletedby=0 AND status = 'I'" ' and due_date >='" & Format(Now, "mm/dd/yyyy") & "'"
    'lstview.lstDatabase sql, lstpayments, 2

    sql = "SELECT id as 'ID Number',account_no as 'Account Number',billto as 'Billing Date',due_date as 'Due Date',total_amount as 'Total Amount',amountpaid as 'Amount Paid', promisorry as 'Promisorry Note' FROM iwadco_readings WHERE account_no = '" & frmpayments.lstpayments.SelectedItem.SubItems(1) & "' AND deletedby=0 AND status = 'I'" ' and due_date >='" & Format(Now, "mm/dd/yyyy") & "'"
    lstview.lstDatabase sql, frmpayments.lstpayments, 1
    Unload Me
    Set frmpay = Nothing

End Sub

Private Sub Form_Load()
'find connection type and amount to paid
DTPicker1.Value = Format(Now, "mm/dd/yyyy")
rsCheck
rs.Open "SELECT *,type,iwadco_typcon.id FROM iwadco_cons,iwadco_typcon WHERE iwadco_cons.id ='" & frmpayments.lstpayments.SelectedItem.SubItems(1) & "'AND iwadco_typcon.id = iwadco_cons.class"
txtcontype.Text = rs("type")
lblamount.Caption = frmpayments.lstpayments.SelectedItem.SubItems(4) - frmpayments.lstpayments.SelectedItem.SubItems(5)
Option1.Value = True
'find for advance payment
rsCheck
sql = "spComputeTotal'" & frmpayments.lstpayments.SelectedItem.SubItems(1) & "'"
Debug.Print sql
rs.Open sql, CN, adOpenStatic, adLockOptimistic

'If Val(txtCP.Text) - Val(txtAP.Text) < 0 Then

txtAP.Text = (rs("change") - rs("hello"))
lblamount.Caption = Val(txtCP.Text) - Abs(Val(txtAP.Text))
If Val(txtAP.Text) < 0 Then
    txtAP.Text = rs("change")
End If

If Val(lblamount.Caption) < 0 Then
    lblamount.Caption = 0
    txtCA.Enabled = False
    txtCA.Text = "0"
End If
'End If
End Sub

Private Sub isButton1_Click()
Unload Me
End Sub

Private Sub Option1_Click()
txtCP.Text = lblamount.Caption
txtCP.Locked = True
txtremarks.Text = "Full Payment"
txtCP.Locked = True
Label8.Visible = True
Label5.Visible = True
Label1.Visible = True
txtAP.Visible = True
txtCP.Visible = True
txtCA.Visible = True
End Sub

Private Sub Option2_Click()
txtCP.Text = ""
txtCP.Text = lblamount.Caption
txtCA.Text = ""
txtremarks.Text = "Partial Payment"
txtCP.Locked = True
Label8.Visible = True
Label5.Visible = True
Label1.Visible = True
txtAP.Visible = True
txtCP.Visible = True
txtCA.Visible = True
End Sub

Private Sub Option5_Click()

If frmpayments.lstpayments.SelectedItem.SubItems(6) = "Yes       " Then
    MsgBox "This account is already under a promisorry note", vbExclamation, "Promisorry Note"
    Option5.Value = False
    Option1.Value = True
    Else
    txtremarks.Text = ""
    txtremarks.Text = "Can't Pay Now"
    Label8.Visible = False
    Label5.Visible = False
    Label1.Visible = False
    txtAP.Visible = False
    txtCP.Visible = False
    txtCA.Visible = False
End If
End Sub

Private Sub txtCA_Change()
txtCA.Text = str_Filter(txtCA, 48, 57, 46)
If Val(txtCA.Text) >= 1 Then
    If Val(txtCP.Text) >= 1 Then
        txtCA.Locked = False
      
        ElseIf Val(txtCP.Text) <= 0 Then
        txtCA.Locked = True
    End If
End If
End Sub

Private Sub txtCP_Change()
'txtCP.Text = txtCP, 48, 57, 46)
If Val(txtCP.Text) > lblamount.Caption Then
    txtCP.Text = lblamount.Caption
    SendKeys "{end}"
End If
If Val(txtCP.Text) > 0 Then
    txtCA.Locked = False
    Else
    txtCA.Locked = True
End If
End Sub

Private Sub txtORNumber_Change()
txtCP.Text = str_Filter(txtCP, 48, 57, 46)
End Sub
