VERSION 5.00
Begin VB.Form frmTappingFee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tapping Fee Payment"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   6135
      Begin VB.Frame Frame4 
         Caption         =   "Types Of Payments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   2295
         Begin VB.OptionButton Option4 
            Caption         =   "Bank"
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
            TabIndex        =   12
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Cash"
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
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
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
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2175
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
            TabIndex        =   9
            Top             =   720
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
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
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
         TabIndex        =   0
         Top             =   2280
         Width           =   2655
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
         TabIndex        =   1
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtCCA 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   2
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
         TabIndex        =   3
         Top             =   3720
         Width           =   5655
      End
      Begin VB.Label Label8 
         BackColor       =   &H00BB5900&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Payment :"
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
         TabIndex        =   19
         Top             =   2280
         Width           =   2415
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
         TabIndex        =   18
         Top             =   1920
         Width           =   1395
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
         TabIndex        =   17
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   240
         X2              =   5880
         Y1              =   1800
         Y2              =   1800
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
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
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
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackColor       =   &H00BB5900&
         BackStyle       =   0  'Transparent
         Caption         =   "Change Cash Amount :"
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
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   3000
         Width           =   2175
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
         TabIndex        =   13
         Top             =   3480
         Width           =   5415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   240
         X2              =   5880
         Y1              =   3480
         Y2              =   3480
      End
   End
   Begin Project1.isButton cmdclose 
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "frmTappingFee.frx":0000
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
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "frmTappingFee.frx":1709A
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
   Begin VB.Image Image1 
      Height          =   720
      Index           =   1
      Left            =   120
      Picture         =   "frmTappingFee.frx":1D8FC
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tapping Fee Payments"
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
      TabIndex        =   20
      Top             =   360
      Width           =   2505
   End
End
Attribute VB_Name = "frmTappingFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstCnt As Double
Dim default() As Double
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmsave_Click()
Dim ReceiptPrint As Variant
If Val(txtCP.Text) > 0 And Val(txtCA.Text) > 0 And Val(txtCCA.Text) >= 0 Then
    rsCheck
    rs.Open "SELECT * FROM iwadco_tappingfee"
    'if not
    rs.AddNew
    rs("account_no") = frmpayments.txtacctno.Text
    rs("invoice") = Format(Now, "mmyyddhmmss")
    rs("amount_payed") = txtCP.Text
    rs("change") = txtCCA.Text
    rs("dateofpayment") = Format(Now, "mm/dd/yyyy")
    rs("remarks") = txtremarks.Text
    rs("lastedit") = Format(Now, "mm/dd/yyyy")
    rs("lasteditby") = empID
    rs.Update
    MsgBox "New Tapping New Added", vbInformation, Me.Caption

    'query another
    rsCheck
    rs.Open "SELECT * FROM iwadco_cons WHERE id = '" & frmpayments.txtacctno.Text & "'"
    If rs.RecordCount > 0 Then
        rs("amountpaid") = rs("amountpaid").value + Val(txtCP.Text)
        rs.Update
    End If
    'end query
    'f out
    'FIND IF TOPPING FEE IS COMPLETE IF TAPPING FEE = 3500
    rsCheck
    'rMBR amountpaid is for iwadco_cons
    'amount_payed is for tbltappingfee
    rs.Open "SELECT * FROM iwadco_cons WHERE id ='" & frmpayments.txtacctno.Text & "'"
                                        'default(2) tappping fee
                If rs("amountpaid") >= default(2) Then
                rs("tappingstatus") = "C"
                rs.Update
                End If
Unload Me
Set frmTappingFee = Nothing
Else
MsgBox "Cannot process payment some of fields may contain no Value or Negative value", vbExclamation, Me.Caption
End If

End Sub

Private Sub Form_Load()
Dim TappingFee As Double
'find connection type and amount to paid
rsCheck
rs.Open "SELECT _value FROM iwadco_default WHERE type='Tapping Fee'"
TappingFee = rs(0)
rsCheck
rs.Open "SELECT * FROM iwadco_cons WHERE id= '" & frmpayments.txtacctno.Text & "'"
If rs.RecordCount > 0 Then
lblamount.Caption = TappingFee - rs("amountpaid").value
Else
lblamount.Caption = TappingFee
End If
Option1.value = True
rsCheck
rs.Open "SELECT * FROM iwadco_default ORDER BY id"
ReDim default(rs.RecordCount) As Double
For lstCnt = 0 To rs.RecordCount - 1
   default(lstCnt) = rs(2)
   rs.MoveNext
Next
End Sub

Private Sub isButton1_Click()
Unload Me
End Sub

Private Sub Option1_Click()
txtCP.Text = lblamount.Caption
txtCP.Locked = True
txtremarks.Text = "Full Payments"
End Sub

Private Sub Option2_Click()
txtCP.Text = ""
txtCA.Text = ""
txtCCA.Text = ""
txtremarks.Text = "Partial Payments"
txtCP.Locked = False
End Sub

Private Sub txtCA_Change()
txtCA.Text = str_Filter(txtCA, 48, 57, 46)
If Val(txtCA.Text) >= 1 Then
    If Val(txtCP.Text) >= 1 Then
        txtCA.Locked = False
        txtCCA.Text = txtCA.Text - txtCP.Text
        ElseIf Val(txtCP.Text) <= 0 Then
        txtCA.Locked = True
    End If
End If
End Sub

Private Sub txtCP_Change()
txtCP.Text = str_Filter(txtCP, 48, 57, 46)
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

