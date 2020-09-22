VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEDITREADING 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Reading"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6450
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "View and Adjust Customer Payments"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   6255
      Begin MSComctlLib.ListView lstpayments 
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4260
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
         ForeColor       =   4210752
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjust Present and Previous Reading .."
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
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   6255
      Begin VB.CheckBox chcknoreading 
         Caption         =   "No Reading"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1920
         TabIndex        =   13
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtarrears 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtpresent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtprevious 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   1140
         Width           =   3255
      End
      Begin VB.Label lblarrears 
         BackStyle       =   0  'Transparent
         Caption         =   "Arrears :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Present Reading :"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Reading :"
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
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
      End
   End
   Begin Project1.isButton cmdExit 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      icon            =   "frmEDITREADING.frx":0000
      style           =   5
      caption         =   "&Close"
      iconsize        =   18
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frmEDITREADING.frx":1709C
   End
   Begin Project1.isButton cmsave 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
      _extentx        =   2355
      _extenty        =   661
      icon            =   "frmEDITREADING.frx":170C4
      style           =   5
      caption         =   "&Save"
      iconsize        =   17
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "frmEDITREADING.frx":1D928
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   0
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
            Picture         =   "frmEDITREADING.frx":1D950
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEDITREADING.frx":1DEEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This form allow to edit existing Previous and Present Reading"
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
      Height          =   210
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   405
      Width           =   5010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Reading, Payments Form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   45
      Width           =   4635
   End
End
Attribute VB_Name = "frmEDITREADING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim default() As Double
Dim Per_Cubic_Meter As Double
Dim min_rate As Double
Dim on_EXCESS() As Double
Dim Amt_EXCESS, EXCESS, X, Exc4 As Double
Dim TAX() As Double
Dim conTypeID, E As Double
Private Sub chcknoreading_Click()

If chcknoreading.Value = 0 Then
    txtprevious.Enabled = True
    txtpresent.Enabled = True
    Else
    txtprevious.Enabled = False
    txtpresent.Enabled = False
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
Set frmEDITREADING = Nothing
End Sub

Private Sub cmsave_Click()
If chcknoreading.Value = 0 Then
    If txtprevious.Text = "" Or txtpresent.Text = "" Then
    MsgBox "Zero value and Null value are not allowed", vbExclamation, Me.Caption
    Exit Sub
    End If
    '--sql
    'get percubic meter
    rsCheck
    rs.Open "SELECT per_cubic_m FROM iwadco_onexcss,iwadco_cons WHERE iwadco_cons.class=iwadco_onexcss.typeid and iwadco_cons.id = '" & frmViewRecords.RecList.SelectedItem.Text & "'"
    Per_Cubic_Meter = rs(0)
    
    'get min_rate
    rsCheck
    'rs.Open "SELECT min_rate FROM iwadco_typcon,iwadco_cons WHERE iwadco_cons.class=iwadco_typcon.id AND iwadco_cons.id = '" & frmViewRecords.RecList.SelectedItem.Text & "'"
    rs.Open "SELECT min_rate,iwadco_typcon.id FROM iwadco_typcon,iwadco_cons WHERE iwadco_cons.class=iwadco_typcon.id AND iwadco_cons.id = '" & frmViewRecords.RecList.SelectedItem.Text & "'"
    min_rate = rs(0)
    conTypeID = rs("id")
    'find id
    
    rsCheck
    rs.Open "SELECT * FROM iwadco_readings WHERE id = '" & frmViewRecords.RecList.SelectedItem.SubItems(1) & "'"
        rs("previous_reading") = txtprevious.Text
        rs("present_reading") = txtpresent.Text
        rs("arrears") = Format(txtarrears.Text, "##0.00")
        rs("consume") = CDbl(txtpresent.Text) - CDbl(txtprevious.Text)
        rs("status") = "I"                              '  default(1) (-10) to get the excess
        'EXCESS = (CDbl(txtReading.Text) - CDbl(txtPrevReading.Text)) - default(1)
        EXCESS = 0
        Exc4 = 0
        Amt_EXCESS = 0
        
        EXCESS = (CDbl(txtpresent.Text) - CDbl(txtprevious.Text))
            If conTypeID >= 2 And conTypeID <= 5 Then
                EXCESS = EXCESS - default(1)
                If EXCESS > 9 Then
                    If EXCESS >= 11 And EXCESS <= 20 Then
                        Amt_EXCESS = 10 * on_EXCESS(1)
                    ElseIf EXCESS >= 21 And EXCESS <= 30 Then
                        Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(2))
                    ElseIf EXCESS >= 31 And EXCESS <= 40 Then
                        Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(3))
                    ElseIf EXCESS >= 1 And EXCESS <= 10 Then
                    Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(1))
                    ElseIf EXCESS >= 41 Then
                    Amt_EXCESS = 10 * on_EXCESS(1)
                    End If
                ElseIf EXCESS >= 1 And EXCESS <= 10 Then
                    Amt_EXCESS = EXCESS * on_EXCESS(1)
                'ElseIf EXCESS <= 0 Then
               '     Amt_EXCESS = min_rate
                End If
                
                If EXCESS > 0 Then
                E = EXCESS
                End If
    
                 EXCESS = EXCESS - default(1)
               
                 If EXCESS > 0 And EXCESS <= 9 Then
                    If EXCESS >= 1 And EXCESS <= 10 Then
                    Amt_EXCESS = Amt_EXCESS + (EXCESS * on_EXCESS(2))
                     End If
                End If
                
                If EXCESS >= 10 Then
                    If EXCESS >= 11 And EXCESS <= 20 Then
                        Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(1))
                    ElseIf EXCESS >= 21 And EXCESS <= 30 Then
                        Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(2))
                    ElseIf EXCESS >= 31 And EXCESS <= 40 Then
                        Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(3))
                    ElseIf EXCESS <= 0 Then
                    Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(1))
                    ElseIf EXCESS >= 1 And EXCESS <= 10 Then
                    Amt_EXCESS = Amt_EXCESS + (EXCESS * on_EXCESS(2))
                    ElseIf EXCESS >= 41 Then
                    Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(2))
                    End If
                End If
                
                If EXCESS > 0 Then
                E = EXCESS
                End If
                EXCESS = EXCESS - default(1)
                
                If EXCESS > 0 And EXCESS <= 9 Then
                     If EXCESS >= 1 And EXCESS <= 10 Then
                     Amt_EXCESS = Amt_EXCESS + (EXCESS * on_EXCESS(3))
                     End If
                End If
                
                If EXCESS >= 10 Then
                    If EXCESS >= 11 And EXCESS <= 20 Then
                        Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(1))
                    ElseIf EXCESS >= 21 And EXCESS <= 30 Then
                        Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(2))
                    ElseIf EXCESS >= 31 And EXCESS <= 40 Then
                        Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(3))
                    ElseIf EXCESS >= 41 Then
                    Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(3))
                          MsgBox Amt_EXCESS
                    ElseIf EXCESS <= 0 Then
                    Amt_EXCESS = Amt_EXCESS + (10 * on_EXCESS(1))
                    ElseIf EXCESS >= 1 And EXCESS <= 10 Then
                    Amt_EXCESS = Amt_EXCESS + (EXCESS * on_EXCESS(1))
                    End If
                End If
                
                If EXCESS > 0 Then
                E = EXCESS
                End If
                
                EXCESS = EXCESS - default(1)
                
                   If EXCESS > 0 And EXCESS <= 9 Then
                    If EXCESS >= 1 And EXCESS <= 10 Then
                    Amt_EXCESS = Amt_EXCESS + (EXCESS * on_EXCESS(4))
                     End If
                End If
                
                If EXCESS >= 10 Then
                     If EXCESS >= 11 And EXCESS <= 20 Then
                        Amt_EXCESS = Amt_EXCESS + (EXCESS * on_EXCESS(4))
                    ElseIf EXCESS >= 21 And EXCESS <= 30 Then
                        Amt_EXCESS = Amt_EXCESS + (EXCESS * on_EXCESS(4))
                    ElseIf EXCESS >= 31 And EXCESS <= 40 Then
                     Amt_EXCESS = Amt_EXCESS + (EXCESS * on_EXCESS(4))
                    ElseIf EXCESS >= 41 Then
                    Amt_EXCESS = Amt_EXCESS + (EXCESS * on_EXCESS(4))
              
                    ElseIf EXCESS >= 1 And EXCESS <= 10 Then
                    Amt_EXCESS = Amt_EXCESS + (EXCESS * on_EXCESS(4))
          
                    End If
                End If
                 'Special Account
        ElseIf conTypeID = 6 Then
            EXCESS = EXCESS - default(1)
             If EXCESS < 0 Then
                EXCESS = 0
                Else
                Amt_EXCESS = EXCESS * Per_Cubic_Meter
             End If
            Else
            'FOR NONE COMMERCIAL USER'S
             If conTypeID = 1 Then
            EXCESS = EXCESS - default(1)
            End If
            'Flat Rate
            If conTypeID = 7 Then
                EXCESS = EXCESS - default(1)
            End If
           'GET RESIDENTIAL II AND RESIDENTIAL III
           If conTypeID = 8 Then
            EXCESS = EXCESS - 60
           End If
           
           If conTypeID = 9 Then
           EXCESS = EXCESS - 30
           End If
           
            If EXCESS < 0 Then
            EXCESS = 0
            Else
            Amt_EXCESS = EXCESS * Per_Cubic_Meter
            End If

        
            rs("excess") = EXCESS
            rs("amount_excess") = Format(Amt_EXCESS + min_rate, "###,##0.00")
            rs("wtax") = Format(((Amt_EXCESS + min_rate) * TAX(2)), "###,##0.00")
            rs("total_amount") = Format(((Amt_EXCESS + min_rate)) + ((Amt_EXCESS + min_rate) * TAX(2)), "###,##0.00")
            rs("trxdate") = Format(Now, "mm/dd/yyyy")
            rs("emp_inchage") = empID
            rs.Update
        End If
      
        If EXCESS > 0 Then
        E = EXCESS
        End If
        EXCESS = E
        If EXCESS < 0 Then
        EXCESS = 0
        End If
    
            rs("excess") = EXCESS
            rs("amount_excess") = Format(Amt_EXCESS + min_rate, "###,##0.00")
            rs("wtax") = Format(((Amt_EXCESS + min_rate) * TAX(2)), "###,##0.00")
            rs("total_amount") = Format(((Amt_EXCESS + min_rate)) + ((Amt_EXCESS + min_rate) * TAX(2)) + Val(txtarrears.Text), "###,##0.00")
            rs("trxdate") = Format(Now, "mm/dd/yyyy")
            rs("emp_inchage") = empID
            rs.Update
Else
rsCheck
rs.Open "SELECT * FROM iwadco_readings WHERE id = '" & frmViewRecords.RecList.SelectedItem.SubItems(1) & "'"
        rs("previous_reading") = 0
        rs("present_reading") = 0
        rs("arrears") = Format(txtarrears.Text, "##0.00")
        rs("consume") = 0
        rs("excess") = 0
        rs("amount_excess") = 0
        rs("wtax") = 0
        rs("total_amount") = Format(txtarrears.Text, "##0.00")
        rs("trxdate") = Format(Now, "mm/dd/yyyy")
        rs("emp_inchage") = empID
        rs.Update
End If
      
    MsgBox "Record has been succesfully updated", vbInformation, Me.Caption
    'refresh the data in frmviewrecords
    sql = "SELECT Account_no as 'Account Number',id as 'Record ID', billto as 'Billing Date',due_date as 'Due Date',previous_reading as 'Previous Reading',present_reading as 'Present Readings',consume as 'Consume' ,excess as 'Excess',amount_excess as 'Amount Excess',arrears as 'Arrears',total_amount as 'Total Amount',amountpaid as 'Amount Paid' FROM iwadco_readings WHERE account_no = '" & frmViewRecords.RecList.SelectedItem.Text & "' AND deletedby=0 ORDER BY id"
    lstview.lstDatabase sql, frmViewRecords.RecList, 1
    Unload Me
End Sub

Private Sub Form_Load()
Dim lstCnt, lstCnt2, lstCnt3 As Double
rsCheck
rs.Open "SELECT * FROM iwadco_default ORDER BY id"
ReDim default(rs.RecordCount) As Double
For lstCnt = 0 To rs.RecordCount - 1
   default(lstCnt) = rs(2)
   rs.MoveNext
Next

'get ONEXCESS amount
rsCheck
rs.Open "SELECT * FROM iwadco_onexcss"
ReDim on_EXCESS(rs.RecordCount) As Double
For lstCnt2 = 0 To rs.RecordCount - 1
    on_EXCESS(lstCnt2) = rs(2)
    rs.MoveNext
Next

'GET DEFAULT amount
rsCheck
rs.Open "SELECT * FROM iwadco_default"
ReDim TAX(rs.RecordCount) As Double
For lstCnt3 = 0 To rs.RecordCount - 1
    TAX(lstCnt3) = rs("_value")
    rs.MoveNext
Next



'--PAYMENTS LIST
sql = "SELECT invoice as 'Invoice Number',amountpayed as 'Amount Paid',change as 'Amount Change',dateofpayment as 'Date Of Payment',remarks as 'Remarks' FROM iwadco_payments WHERE id ='" & frmViewRecords.RecList.SelectedItem.SubItems(1) & "'"
lstview.lstDatabase sql, lstpayments, 1



End Sub


Private Sub lstpayments_Click()
On Error GoTo err
frm_EDITPAYMENTS.txtinvoice.Text = lstpayments.SelectedItem.Text
frm_EDITPAYMENTS.txtAP.Text = lstpayments.SelectedItem.SubItems(1)
frm_EDITPAYMENTS.txtchange.Text = lstpayments.SelectedItem.SubItems(2)
frm_EDITPAYMENTS.txtDOP.Value = lstpayments.SelectedItem.SubItems(3)
frm_EDITPAYMENTS.txtremarks = lstpayments.SelectedItem.SubItems(4)
frm_EDITPAYMENTS.Show 1
err:
Exit Sub
End Sub

Private Sub txtarrears_Change()
txtarrears.Text = str_Filter(txtarrears, 48, 57, 46)
End Sub

Private Sub txtpresent_Change()
txtpresent.Text = str_Filter(txtpresent, 48, 57, 0)
End Sub

Private Sub txtprevious_Change()
txtprevious.Text = str_Filter(txtprevious, 48, 57, 0)
End Sub
