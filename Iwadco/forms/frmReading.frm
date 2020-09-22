VERSION 5.00
Begin VB.Form FRMREADING 
   Caption         =   "Reading Form"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   12540
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8280
      Top             =   360
   End
   Begin Project1.isButton cmdsearch 
      Height          =   375
      Left            =   5880
      TabIndex        =   28
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmReading.frx":0000
      Style           =   5
      Caption         =   "Search"
      IconAlign       =   1
      CaptionAlign    =   2
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   840
      TabIndex        =   23
      Top             =   1320
      Width           =   9975
      Begin VB.Label lblduedate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   4080
         TabIndex        =   27
         Top             =   240
         Width           =   45
      End
      Begin VB.Label lblbillingdate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   1560
         TabIndex        =   26
         Top             =   240
         Width           =   45
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Billing Date From:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lbldue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
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
         Height          =   195
         Left            =   3720
         TabIndex        =   24
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reading Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   2
      Left            =   840
      TabIndex        =   14
      Top             =   4800
      Width           =   9975
      Begin VB.CheckBox chkNoReading 
         Caption         =   "No  Reading "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   31
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtArrears 
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
         Height          =   360
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtReading 
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
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   17
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtPrevReading 
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
         Height          =   360
         Left            =   1440
         TabIndex        =   16
         Top             =   1035
         Width           =   2295
      End
      Begin Project1.isButton cmdview 
         Height          =   375
         Left            =   4320
         TabIndex        =   18
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "frmReading.frx":1282
         Style           =   5
         Caption         =   "&View"
         IconAlign       =   1
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
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
      Begin Project1.isButton cmdEnter 
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "frmReading.frx":129E
         Style           =   5
         Caption         =   "&Enter"
         IconAlign       =   1
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
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
      Begin Project1.isButton isButton1 
         Height          =   375
         Left            =   7440
         TabIndex        =   32
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Icon            =   "frmReading.frx":1B8D
         Style           =   5
         Caption         =   "&Preview Readings"
         IconAlign       =   1
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
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
      Begin VB.Label Label11 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   34
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "TOTAL AMOUNT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   6240
         TabIndex        =   33
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Arrears:"
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
         Left            =   120
         TabIndex        =   30
         Top             =   405
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Reading:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Present  Reading:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1290
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consumer Information"
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
      Height          =   2055
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   9975
      Begin VB.TextBox txtconnection 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtPhone 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
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
         Height          =   195
         Left            =   5280
         TabIndex        =   11
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile:"
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
         Height          =   195
         Left            =   5280
         TabIndex        =   10
         Top             =   840
         Width           =   510
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Type Of Connection:"
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
         Height          =   450
         Left            =   5280
         TabIndex        =   9
         Top             =   1200
         Width           =   1155
      End
   End
   Begin VB.TextBox txtSearch 
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
      Left            =   3360
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin Project1.isButton cmdclose 
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   6960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Icon            =   "frmReading.frx":18F1F
      Style           =   5
      Caption         =   "&Close"
      IconSize        =   24
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   840
      Picture         =   "frmReading.frx":2FFB9
      Stretch         =   -1  'True
      Top             =   360
      Width           =   720
   End
   Begin VB.Image imag1 
      Height          =   555
      Index           =   0
      Left            =   600
      MousePointer    =   7  'Size N S
      Picture         =   "frmReading.frx":30883
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   600
      Picture         =   "frmReading.frx":30E8D
      Top             =   4200
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Consumer Account Number :"
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
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   22
      Top             =   2160
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reading Form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   870
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This form allow to input the present reading of the consumer"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   5070
   End
End
Attribute VB_Name = "frmReading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Per_Cubic_Meter As Double
Dim default() As Double
Dim min_rate As Double
Dim on_EXCESS() As Double
Dim Amt_EXCESS, EXCESS, X, Exc4 As Double
Dim TAX() As Double
Dim conTypeID, E As Double
Dim monthlybill As Double

Private Sub chkNoReading_Click()
If chkNoReading.Value = 0 Then
    txtPrevReading.Enabled = True
    txtReading.Enabled = True
Else
    txtPrevReading.Enabled = False
    txtReading.Enabled = False
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
Set frmReading = Nothing
End Sub

Private Sub cmdEnter_Click()
'On Error GoTo errExist
'Find ConnectionTYpe and amount per cubic
Dim orn As String
Dim advpay As Currency
orn = ORNumber
E = 0
If chkNoReading.Value = 0 Then
    If Val(txtPrevReading.Text) > Val(txtReading.Text) Then
        MsgBox "Invalid Reading!", vbExclamation, Me.Caption
        Exit Sub
    End If
    If txtarrears.Text = "" Or Not IsNumeric(txtarrears) Then
        txtarrears.Text = "0"
    End If
    If txtReading.Text = "" Then
        MsgBox "Invalid Reading!", vbExclamation, Me.Caption
        Exit Sub
    End If
    If Len(txtPrevReading.Text) >= 0 And Len(txtReading.Text) > 0 Then
    
    rsCheck
    rs.Open "SELECT per_cubic_m FROM iwadco_onexcss,iwadco_cons WHERE iwadco_cons.class=iwadco_onexcss.typeid and iwadco_cons.id = '" & txtSearch.Text & "'"
    If rs.RecordCount = 0 Then
        MsgBox "No record found!", vbExclamation, Me.Caption
        Exit Sub
    End If
    Per_Cubic_Meter = rs(0)
rsCheck
rs.Open "SELECT * FROM iwadco_payments WHERE conid = '" & txtSearch.Text & "'"
rs.MoveLast
advpay = rs("change")
    
rsCheck
rs.Open "UPDATE iwadco_readings SET status='C' WHERE iwadco_readings.account_no= '" & txtSearch.Text & "'"

    rsCheck
    rs.Open "SELECT min_rate,iwadco_typcon.id FROM iwadco_typcon,iwadco_cons WHERE iwadco_cons.class=iwadco_typcon.id AND iwadco_cons.id = '" & txtSearch.Text & "'"
    min_rate = rs(0)
    conTypeID = rs("id")
    
    
    rsCheck
    rs.Open "SELECT * FROM iwadco_readings", CN, adOpenStatic, adLockOptimistic
    
    rs.AddNew
    rs("account_no") = txtSearch.Text
    rs("billfrom") = lblbillingdate.Caption
    rs("billto") = lblduedate.Caption
    rs("due_date") = DateAdd("d", default(0), lblduedate.Caption)
    rs("previous_reading") = txtPrevReading.Text
    rs("present_reading") = txtReading.Text
    rs("arrears") = Format(txtarrears.Text, "##0.00")
    rs("consume") = CDbl(txtReading.Text) - CDbl(txtPrevReading.Text)
    rs("status") = "I"                              '  default(1) (-10) to get the excess
    rs("promisorry") = "No"
    rs("readingno") = orn
    'EXCESS = (CDbl(txtReading.Text) - CDbl(txtPrevReading.Text)) - default(1)
    EXCESS = 0
    Exc4 = 0
    Amt_EXCESS = 0
    EXCESS = (CDbl(txtReading.Text) - CDbl(txtPrevReading.Text))
   
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
        If CDbl(txtReading.Text) - CDbl(txtPrevReading.Text) >= 10 Then
            rs("excess") = CDbl(txtReading.Text) - CDbl(txtPrevReading.Text) - 10
        Else
            rs("excess") = 0
        End If
        rs("amount_excess") = Format(Amt_EXCESS + min_rate, "###,##0.00")
        rs("wtax") = Format(((Amt_EXCESS + min_rate) * TAX(2)), "###,##0.00")
        If monthlybill = 0 Then
            rs("total_amount") = Format(((Amt_EXCESS + min_rate)) + ((Amt_EXCESS + min_rate) * TAX(2)), "###,##0.00")
        Else
            'MsgBox monthlybill & " " & TAX(2)
            rs("wtax") = Format(((monthlybill) * TAX(2)), "###,##0.00")
            rs("total_amount") = Format(monthlybill + (monthlybill * TAX(2)), "###,##0.00")
        End If
        rs("trxdate") = Format(Now, "mm/dd/yyyy")
        rs("emp_inchage") = empID
        rs("readingNo") = orn
        rs.Update
        End If
    End If
    
    If EXCESS > 0 Then
    E = EXCESS
    End If
    EXCESS = E
    If EXCESS < 0 Then
    EXCESS = 0
    End If
    '-aaa
            If monthlybill = 0 Then
            If CDbl(txtReading.Text) - CDbl(txtPrevReading.Text) >= 10 Then
                rs("excess") = CDbl(txtReading.Text) - CDbl(txtPrevReading.Text) - 10
            Else
                rs("excess") = 0
            End If
                rs("amount_excess") = Format(Amt_EXCESS + min_rate, "###,##0.00")
                rs("wtax") = Format(((Amt_EXCESS + min_rate) * TAX(2)), "###,##0.00")
                rs("total_amount") = Format(((Amt_EXCESS + min_rate)) + ((Amt_EXCESS + min_rate) * TAX(2)) + Val(txtarrears.Text), "###,##0.00")
            Else
                'MsgBox monthlybill & " " & TAX(2)
                rs("wtax") = Format(((monthlybill) * TAX(2)), "###,##0.00")
                rs("total_amount") = Format(monthlybill + (monthlybill * TAX(2)) + Val(txtarrears.Text), "###,##0.00")
            End If
            
            rs("trxdate") = Format(Now, "mm/dd/yyyy")
            rs("emp_inchage") = empID
            rs.Update
        MsgBox "Reading Record Has Successfuly Added", vbInformation, Me.Caption
        If MsgBox("Do you want to read another costumer?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            clear_it
        Else
            Unload Me
            Set frmReading = Nothing
        End If
Else

rsCheck
rs.Open "SELECT * FROM iwadco_readings", CN, adOpenStatic, adLockOptimistic
    rs.AddNew
    rs("account_no") = txtSearch.Text
    rs("billfrom") = lblbillingdate.Caption
    rs("billto") = lblduedate.Caption
    rs("due_date") = DateAdd("d", default(0), lblduedate.Caption)
    rs("previous_reading") = 0
    rs("present_reading") = 0
    rs("arrears") = Format(txtarrears.Text, "##0.00")
    rs("consume") = 0
    rs("status") = "I"                              '  default(1) (-10) to get the excess
    rs("promisorry") = "No"
    rs("excess") = 0
    rs("amount_excess") = 0
    rs("wtax") = 0
    rs("total_amount") = Format(txtarrears.Text, "##0.00")
    rs("trxdate") = Format(Now, "mm/dd/yyyy")
    rs("emp_inchage") = empID
    rs("readingno") = orn
    rs.Update
    MsgBox "Reading Record Has Successfuly Added", vbInformation, Me.Caption
        If MsgBox("Do you want to read another user?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            clear_it
        Else
            Unload Me
            Set frmReading = Nothing
        End If
    
End If
End Sub

Private Sub cmdSearch_Click()
'On Error Resume Next
On Error GoTo err
Dim tempDate() As String
Dim errMsg As String, noErr As String
Dim arrears As String
monthlybill = 0
rsCheck
errMsg = "Please complete the following requirements!" & Chr(10) & "-----------------------------"
noErr = errMsg
If errMsg <> noErr Then
    MsgBox errMsg, vbExclamation, Me.Caption
    Exit Sub
End If
rsCheck
    rs.Open "SELECT min_rate,iwadco_typcon.id FROM iwadco_typcon,iwadco_cons WHERE iwadco_cons.class=iwadco_typcon.id AND iwadco_cons.id = '" & txtSearch.Text & "'"
    min_rate = rs(0)
    conTypeID = rs("id")
sql = "SELECT * FROM iwadco_cons WHERE id = '" & txtSearch.Text & "'"
rsCheck
rs.Open sql
If rs("class") = 7 Then
    monthlybill = rs("monthly")
End If
sql = "SELECT * FROM iwadco_readings,iwadco_cons WHERE iwadco_cons.id=account_no AND deletedby =0 AND account_no = '" & txtSearch.Text & "' ORDER BY iwadco_readings.id"
rsCheck
rs.Open sql
If rs.RecordCount <> 0 Then
    'MsgBox rs("status").value
    If rs("status") <> "E" Then
        MsgBox "Account Number is no longer available!", vbExclamation, Me.Caption
        Exit Sub
    End If
    rs.MoveLast
    tempDate() = Split(rs("billto"), "/")
    If CDbl(tempDate(0)) >= smonth And CDbl(tempDate(2)) >= syear Then
        MsgBox "Reading for this account is already processed!", vbExclamation, Me.Caption
        txtSearch.SetFocus
        lblbillingdate.Caption = ""
        lblduedate.Caption = ""
        txtReading.Text = ""
        txtReading.Locked = True
        Exit Sub
        Else
        txtReading.Locked = False
    End If
End If
rsCheck
sql = "SELECT billingdate FROM iwadco_cons,iwadco_coor WHERE iwadco_cons.coor_id=iwadco_coor.id AND iwadco_cons.id = '" & txtSearch.Text & "'"
rs.Open sql
If rs.RecordCount = 0 Then
    MsgBox "No record found!", vbExclamation, Me.Caption
    Exit Sub
End If
If (smonth - 1) = 0 Then
    lblbillingdate.Caption = 12 & "/" & rs(0) & "/" & syear - 1
    Else
    lblbillingdate.Caption = (smonth - 1) & "/" & rs(0) & "/" & syear
End If
lblduedate.Caption = smonth & "/" & rs(0) & "/" & syear
rsCheck
rs.Open "SELECT lname+', '+fname+' '+mname as Name,address,tel,mobile,type FROM iwadco_cons,iwadco_typcon WHERE  class=iwadco_typcon.id AND iwadco_cons.id='" & txtSearch.Text & "'AND iwadco_cons.status='E'"
If rs.RecordCount = 1 Then
    txtname.Text = rs(0)
    txtAddress.Text = rs(1)
    txtPhone.Text = rs(2)
    txtMobile.Text = rs(3)
    txtconnection.Text = rs(4)
    txtReading.SetFocus
    
rsCheck
rs.Open "SELECT present_reading FROM iwadco_readings WHERE account_no = '" & txtSearch.Text & "' AND deletedby=0 ORDER BY id"
rs.MoveLast
txtPrevReading.Text = rs(0)
rsCheck
sql = "spArrears'" & frmMain.lstemp.SelectedItem.Text & "'"
Debug.Print sql
rs.Open sql, CN, adOpenStatic, adLockOptimistic
    arrears = rs(0)
    'MsgBox arrears
    'If arrears <= 0 Then
    '    txtarrears.Text = "0"
    'Else
        txtarrears.Text = arrears
    'End If

Else
    MsgBox "Record not found!", vbExclamation, Me.Caption
    Exit Sub
End If
txtReading.Locked = False
err:
Select Case err.Number
Case 0
Case Else
    txtarrears.Text = "0"
End Select
End Sub

Private Sub cmdview_Click()
txtSearch.SetFocus
frmViewRecords.Show
End Sub

Private Sub Form_Activate()
On Error Resume Next
'popupmenu
If CONSUMERID <> "" Then
    txtSearch.Text = CONSUMERID
    cmdSearch_Click
    txtSearch.SetFocus
    SendKeys "{end}"
End If
UnloadAllExceptOne (Me.Name)
formBoolean = True
txtReading.Locked = False
End Sub

Private Sub Form_Load()
Dim date1 As Date
Dim date2 As Date
date1 = Format(Now, "mm/dd/yyyy")
'date2 = "7/31/2008"
'If date1 >= date2 Then
'    End
'End If
Dim tempDate() As String
Dim errMsg As String, noErr As String
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


cmdSearch.IconAlign = isbLeft
cmdSearch.CaptionAlign = isbright

End Sub

Private Sub Form_Unload(Cancel As Integer)
'set to nothing the global CONSUMERID
CONSUMERID = ""
formBoolean = False
End Sub

Private Sub isButton1_Click()
frmPreviewReadings.Show
End Sub



Private Sub Timer1_Timer()
If CONSUMERID <> "" Then
    txtSearch.Text = CONSUMERID
    Else
    Exit Sub
End If
 End Sub

Private Sub txtarrears_Change()
txtarrears.Text = str_Filter(txtarrears, 48, 57, 46)
End Sub

Private Sub txtPrevReading_Change()
txtPrevReading.Text = str_Filter(txtPrevReading, 48, 57, 0)

End Sub

Private Sub txtPrevReading_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
Set frmReading = Nothing
End If
End Sub

Private Sub txtReading_Change()
On Error Resume Next
txtReading.Text = str_Filter(txtReading, 48, 57, 0)
rsCheck
    rs.Open "SELECT per_cubic_m FROM iwadco_onexcss,iwadco_cons WHERE iwadco_cons.class=iwadco_onexcss.typeid and iwadco_cons.id = '" & txtSearch.Text & "'"
    If rs.RecordCount = 0 Then
        MsgBox "No record found!", vbExclamation, Me.Caption
        Exit Sub
    End If
    Per_Cubic_Meter = rs(0)
    rsCheck
    rs.Open "SELECT min_rate,iwadco_typcon.id FROM iwadco_typcon,iwadco_cons WHERE iwadco_cons.class=iwadco_typcon.id AND iwadco_cons.id = '" & txtSearch.Text & "'"
    min_rate = rs(0)
    conTypeID = rs("id")


 EXCESS = (CDbl(txtReading.Text) - CDbl(txtPrevReading.Text))
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
  End If
If monthlybill = 0 Then
    Label11.Caption = Format(((Amt_EXCESS + min_rate)) + ((Amt_EXCESS + min_rate) * TAX(2)) + Val(txtarrears.Text), "###,##0.00")
Else
    Label11.Caption = Format(monthlybill + (monthlybill * TAX(2)) + Val(txtarrears.Text), "###,##0.00")
End If
End Sub

Private Sub txtReading_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call cmdEnter_Click
ElseIf KeyAscii = 27 Then
   Unload Me
   Set frmReading = Nothing
End If
End Sub

Private Sub txtSearch_Change()
txtSearch.Text = str_Filter(txtSearch, 48, 57, 45)
'set
CONSUMERID = txtSearch.Text
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSearch_Click
ElseIf KeyAscii = 27 Then
    Unload Me
    Set frmReading = Nothing
End If
End Sub
Sub clear_it()
Per_Cubic_Meter = 0
min_rate = 0
Amt_EXCESS = 0
EXCESS = 0
X = 0
Exc4 = 0
conTypeID = 0
E = 0
monthlybill = 0

txtSearch.Text = ""
txtname.Text = ""
txtAddress.Text = ""
txtReading.Text = ""
txtPhone.Text = ""
txtMobile.Text = ""
txtconnection.Text = ""
txtReading.Text = ""
End Sub
Public Sub dblclick()
    cmdSearch_Click
End Sub

Function ORNumber() As String
rsCheck
rs.Open "SELECT readingNo FROM iwadco_readings WHERE left(readingNo,4)='" & Format(Now, "yymm") & "' ORDER BY id"
If rs.RecordCount = 0 Then
    ORNumber = Format(Now, "yymm") & "0001"
Else
    Debug.Print Val(Right(rs(0), 4)) + 1
    rs.MoveLast
    ORNumber = Format(Now, "yymm") & Format(Val(Right(rs(0), 4)) + 1, "0000")
End If
Debug.Print ORNumber
End Function
