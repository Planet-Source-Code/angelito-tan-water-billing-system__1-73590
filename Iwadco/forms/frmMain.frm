VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmmain 
   BackColor       =   &H00808080&
   Caption         =   "INPART WATERWORKS & DEVELOPMENT COPRP..(IWADCO) SYSTEM"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14730
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   4080
      Top             =   1440
   End
   Begin VB.PictureBox Picture5 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   155
      Left            =   0
      ScaleHeight     =   150
      ScaleWidth      =   18960
      TabIndex        =   20
      Top             =   930
      Width           =   18960
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8700
      Left            =   0
      ScaleHeight     =   8700
      ScaleWidth      =   3615
      TabIndex        =   16
      Top             =   1080
      Width           =   3615
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   3375
      End
      Begin Project1.isButton isButton1 
         Height          =   375
         Left            =   2280
         TabIndex        =   23
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "frmMain.frx":08CA
         Style           =   5
         Caption         =   "&Search"
         IconAlign       =   1
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
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
      Begin MSComctlLib.StatusBar StatusBar2 
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   8280
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   1
               Bevel           =   2
               Object.Width           =   3334
               Text            =   "Total Consumer"
               TextSave        =   "Total Consumer"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imageEmp 
         Left            =   2880
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":08E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1A9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":82FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8BD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":F438
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture6 
         Height          =   1815
         Left            =   3480
         MousePointer    =   1  'Arrow
         ScaleHeight     =   1755
         ScaleWidth      =   75
         TabIndex        =   19
         Top             =   3240
         Width           =   135
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
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   2055
      End
      Begin MSComctlLib.ListView lstemp 
         Height          =   5655
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   9975
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imageEmp"
         SmallIcons      =   "imageEmp"
         ForeColor       =   0
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Select Area:"
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
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Select Coordinator:"
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
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   "Search by Account number Last name, First name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   18960
      TabIndex        =   2
      Top             =   9780
      Visible         =   0   'False
      Width           =   18960
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   3720
         ScaleHeight     =   855
         ScaleWidth      =   8055
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   8055
         Begin Project1.isButton cmdSearch 
            Height          =   855
            Left            =   5280
            TabIndex        =   4
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1508
            Icon            =   "frmMain.frx":15C9A
            Style           =   8
            Caption         =   "&Search"
            IconSize        =   36
            IconAlign       =   1
            CaptionAlign    =   4
            iNonThemeStyle  =   4
            HighlightColor  =   -2147483629
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
         Begin Project1.isButton cmdRefresh 
            Height          =   855
            Left            =   3960
            TabIndex        =   5
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1508
            Icon            =   "frmMain.frx":2AE0C
            Style           =   8
            Caption         =   "&Refresh"
            IconSize        =   36
            IconAlign       =   1
            CaptionAlign    =   4
            iNonThemeStyle  =   4
            HighlightColor  =   -2147483629
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
         Begin Project1.isButton cmdDelete 
            Height          =   855
            Left            =   2640
            TabIndex        =   6
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1508
            Icon            =   "frmMain.frx":30F96
            Style           =   8
            Caption         =   "&Delete"
            IconSize        =   36
            IconAlign       =   1
            CaptionAlign    =   4
            iNonThemeStyle  =   4
            HighlightColor  =   -2147483629
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
         Begin Project1.isButton cmdEdit 
            Height          =   855
            Left            =   1320
            TabIndex        =   7
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1508
            Icon            =   "frmMain.frx":46108
            Style           =   8
            Caption         =   "&Edit"
            IconSize        =   36
            IconAlign       =   3
            CaptionAlign    =   4
            iNonThemeStyle  =   4
            HighlightColor  =   -2147483629
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
         Begin Project1.isButton cmdNew 
            Height          =   855
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1508
            Icon            =   "frmMain.frx":5B27A
            Style           =   8
            Caption         =   "&Create New"
            IconSize        =   36
            IconAlign       =   3
            CaptionAlign    =   4
            iNonThemeStyle  =   4
            BackColor       =   16777215
            HighlightColor  =   -2147483629
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
         Begin Project1.isButton cmdCancel 
            Height          =   855
            Left            =   6600
            TabIndex        =   9
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1508
            Icon            =   "frmMain.frx":61ADC
            Style           =   8
            Caption         =   "&Close"
            IconSize        =   36
            IconAlign       =   1
            CaptionAlign    =   4
            iNonThemeStyle  =   4
            HighlightColor  =   -2147483629
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
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   925
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   18960
      TabIndex        =   1
      Top             =   0
      Width           =   18960
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   50
         ScaleHeight     =   855
         ScaleWidth      =   7695
         TabIndex        =   10
         Top             =   50
         Width           =   7695
         Begin Project1.isButton cmdDis 
            Height          =   855
            Left            =   3600
            TabIndex        =   13
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1508
            Icon            =   "frmMain.frx":62068
            Style           =   5
            Caption         =   "Disconnect"
            IconSize        =   36
            IconAlign       =   1
            CaptionAlign    =   4
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
         Begin Project1.isButton cmdUtilities 
            Height          =   855
            Left            =   4800
            TabIndex        =   11
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1508
            Icon            =   "frmMain.frx":62A51
            Style           =   5
            Caption         =   "Utilities"
            IconSize        =   36
            IconAlign       =   1
            CaptionAlign    =   4
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
         Begin Project1.isButton cmdReports 
            Height          =   855
            Left            =   2400
            TabIndex        =   12
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1508
            Icon            =   "frmMain.frx":6342B
            Style           =   5
            Caption         =   "Reports"
            IconSize        =   36
            IconAlign       =   1
            CaptionAlign    =   4
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
         Begin Project1.isButton cmdPayments 
            Height          =   855
            Left            =   1200
            TabIndex        =   14
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1508
            Icon            =   "frmMain.frx":63E14
            Style           =   5
            Caption         =   "Payments"
            IconSize        =   36
            IconAlign       =   3
            CaptionAlign    =   4
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
         Begin Project1.isButton cmdReading 
            Height          =   855
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1508
            Icon            =   "frmMain.frx":646B9
            Style           =   5
            Caption         =   "Reading"
            IconSize        =   36
            IconAlign       =   3
            CaptionAlign    =   4
            BackColor       =   16777215
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
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10755
      Width           =   18960
      _ExtentX        =   33443
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   14
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   442
            MinWidth        =   442
            Picture         =   "frmMain.frx":64F1E
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "User Name:"
            TextSave        =   "User Name:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "c"
            TextSave        =   "c"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "frmMain.frx":654B8
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Consumer Name :"
            TextSave        =   "Consumer Name :"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "frmMain.frx":65A52
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Today:"
            TextSave        =   "Today:"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11/1/2010"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12:34 AM"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel14 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
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
      Left            =   1200
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6777E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6845A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B77E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D110
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6EAA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F77C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":70456
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71130
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71E0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":72AE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":733C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":740A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":74D7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7633C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":77018
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":778F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":785D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":79F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B8F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   360
      Top             =   4470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C1D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7DB66
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F4F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":80E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8281C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":841AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":85B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":874D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":88E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8A7F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8B4D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8BDB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8CA90
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D76C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E448
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8F124
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8FE00
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   360
      Top             =   4470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":906DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9206E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":93A00
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":95392
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":96D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":986B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A048
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9B9DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9D36C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9ED00
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F9DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A02BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A0F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A1C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A2950
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A362C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A4308
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogout 
         Caption         =   "&Logout User"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit System"
      End
   End
   Begin VB.Menu mnuTransac 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuProf 
         Caption         =   "Profile Manager"
         Begin VB.Menu mnuCoor 
            Caption         =   "Coordinators Profiles"
         End
         Begin VB.Menu mnuCost 
            Caption         =   "Costumer's Profiles"
         End
         Begin VB.Menu l2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuArea 
            Caption         =   "Add Area"
         End
         Begin VB.Menu mnuSysUser 
            Caption         =   "System User Profiles"
         End
      End
      Begin VB.Menu mnuread 
         Caption         =   "Reading"
      End
      Begin VB.Menu mnupay 
         Caption         =   "Payments"
      End
      Begin VB.Menu mnudis 
         Caption         =   "Disconnect"
      End
   End
   Begin VB.Menu mnurep 
      Caption         =   "Reports"
      Begin VB.Menu mnuStatAcc 
         Caption         =   "Statement of Account"
         Begin VB.Menu mnuRepIndi 
            Caption         =   "Individual"
         End
         Begin VB.Menu mnuRepAll 
            Caption         =   "All"
         End
         Begin VB.Menu mnuSum 
            Caption         =   "Summary"
         End
      End
      Begin VB.Menu mnuAging 
         Caption         =   "Aging of Accounts Receivable"
         Begin VB.Menu mnuDetailed 
            Caption         =   "Detailed"
         End
      End
      Begin VB.Menu mnuCommisions 
         Caption         =   "Commisions"
         Begin VB.Menu mnuCom 
            Caption         =   "Coordinators Commisions"
         End
         Begin VB.Menu mnuSumCom 
            Caption         =   "Commision Summary"
         End
      End
      Begin VB.Menu mnuProfRep 
         Caption         =   "Profiles"
         Begin VB.Menu mnuRepCons 
            Caption         =   "Consumer"
         End
         Begin VB.Menu mnuRepCoor 
            Caption         =   "Coordinators"
         End
      End
      Begin VB.Menu mnuRForm 
         Caption         =   "Reading Form"
      End
      Begin VB.Menu mnuReadingsRep 
         Caption         =   "Readings Reports"
      End
      Begin VB.Menu mnuPaymentsRep 
         Caption         =   "Payments Reports"
      End
      Begin VB.Menu mnuTap 
         Caption         =   "Tapping Fees"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "Promisoy History"
      End
      Begin VB.Menu mnuSumRep 
         Caption         =   "Summary Reports"
      End
   End
   Begin VB.Menu mnuNEW 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuCREATE 
         Caption         =   "Create New"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnudel 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnusearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnul2 
         Caption         =   "-"
      End
      Begin VB.Menu mnurefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "Utilities"
      Begin VB.Menu mnuDef 
         Caption         =   "Default Settings"
         Begin VB.Menu mnuULedit 
            Caption         =   "Minimum Rate And Per Cubic Meter Settings"
         End
         Begin VB.Menu mnuds 
            Caption         =   "Default Settings"
         End
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup Database"
      End
      Begin VB.Menu mnuCleanUp 
         Caption         =   "Cleanup Database"
      End
   End
   Begin VB.Menu mnuGroup 
      Caption         =   "popup2"
      Visible         =   0   'False
      Begin VB.Menu mnureading 
         Caption         =   "Reading"
      End
      Begin VB.Menu mnupayments 
         Caption         =   "Payments"
      End
      Begin VB.Menu mnuVP 
         Caption         =   "View Profile"
      End
      Begin VB.Menu mnuvbr 
         Caption         =   "View Billing Records"
      End
      Begin VB.Menu mnuline 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefresh2 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnupopup3 
      Caption         =   "popup3"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit2 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnudel2 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnulin3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclose2 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuregister 
         Caption         =   "Register"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frm As Form
Dim EmpLogName As String
Dim IV As Integer
Dim dateMOVE As Date
Dim dateSTARTED As Date
Dim dateEXPIRED As Date
Dim X As Integer
Dim getDate     As Date
Dim regDecrypt As String
Dim frmLoginSHOW As Boolean
Dim coor_ As Integer
Dim area_ As Integer

Private Sub cmdcancel_Click()
On Error Resume Next
ActiveForm.clse
End Sub

Private Sub cmddelete_Click()
On Error Resume Next
ActiveForm.del
End Sub

Private Sub cmdDis_Click()
mnudis_Click
End Sub

Private Sub cmdedit_Click()
On Error Resume Next
ActiveForm.edit
End Sub

Private Sub cmdNew_Click()
On Error Resume Next
ActiveForm.ADD
End Sub

Private Sub cmdPayments_Click()
mnuPay_Click
End Sub

Private Sub cmdReading_Click()
mnuRead_Click
'frmReading.Show
End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
ActiveForm.Ref
End Sub


Private Sub cmdReports_Click()
PopupMenu mnurep
End Sub

Private Sub cmdSearch_Click()
mnusearch_Click
End Sub

Private Sub cmdUtilities_Click()
PopupMenu mnuUtil
End Sub

Private Sub Combo1_Click()
sql = "SELECT id FROM iwadco_area WHERE area = '" & Combo1.Text & "'"
rsCheck
rs.Open sql, CN, adOpenStatic, adLockOptimistic
If rs.RecordCount = 1 Then
    area_ = rs(0)
Else
    area_ = rs(0)
End If
End Sub

Private Sub Combo3_Click()
Combo1.Clear
sql = "SELECT id FROM iwadco_coor WHERE lname+', '+fname+' '+mname = '" & Combo3.Text & "'"
rsCheck
rs.Open sql, CN, adOpenStatic, adLockOptimistic
If rs.RecordCount = 1 Then
    coor_ = rs(0)
Else
    coor_ = 0
End If
LoadCombo "SELECT area FROM iwadco_area WHERE coor_id=" & coor_, Combo1
End Sub

Private Sub isButton1_Click()
sql = "SELECT id as 'Account No.',lname+', '+fname+' '+mname as 'Full Name' FROM iwadco_cons WHERE lname+', '+fname+' '+mname LIKE '" & Replace(txtSearch.Text, "'", "`") & "%' AND status<>'D'"
If Combo3.Text <> "" Then
    sql = sql & " AND coor_id=" & coor_
End If
If Combo1.Text <> "" Then
    sql = sql & " AND area_id = " & area_
End If
sql = sql & " ORDER BY lname+', '+fname+' '+mname"
StatusBar2.Panels.Item(2).Text = Format(lstview.lstDatabase(sql, lstemp, 1), "###,###,###")
End Sub

Private Sub lstemp_Click()

StatusBar1.Panels.Item(6).Text = " " & lstemp.SelectedItem.SubItems(1)
CONSUMERID = lstemp.SelectedItem.Text
End Sub

Private Sub lstemp_DblClick()
On Error Resume Next

If formBoolean = True Then
    ActiveForm.dblclick
End If
End Sub

Private Sub lstemp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstemp.ListItems.Count < 1 Then Exit Sub
If Button = 2 Then PopupMenu mnuGroup
End Sub


Private Sub MDIForm_Activate()
'--GET NAME OF USER WHO LOG
rsCheck
rs.Open "SELECT username from iwadco_user WHERE id ='" & empID & "'", CN, adOpenStatic, adLockOptimistic
EmpLogName = rs("username")
StatusBar1.Panels.Item(3).Text = " " & EmpLogName
'count record
rsCheck
rs.Open "SELECT * FROM iwadco_cons WHERE status<>'D'"
StatusBar2.Panels.Item(2).Text = Format(rs.RecordCount, "###,###,###")
End Sub

Private Sub MDIForm_Load()
'----REGISTRATION
Dim date1 As Date
Dim date2 As Date
date1 = Format(Now, "mm/dd/yyyy")
'date2 = "8/31/2008"
'If date1 >= date2 Then
'    End
'End If
If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "Reg") = "Error" Then
    If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "DateStart") = "Error" Then
         CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft")
         SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "DateStart", Date
         Else
         'condition if someone has adjust the time
         If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "DateMove") = "Error" Then
            SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "DateMove", Date
            Else
                getDate = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "DateMove")
                If Date - getDate >= 0 Then
                SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "DateMove", Date
                X = X + 1
                Else
                MsgBox "Adjusting Time Will Expired Your Trial", vbCritical
                FRMREG.Show 1
                Exit Sub
                End If
        End If
      End If
      If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "DateExpired") = "Error" Then
         CreateKey ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft")
         SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "DateExpired", Date + 30
         X = X + 1

         Else
         X = X + 1
      End If
      
     If X >= 2 Then
     dateMOVE = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "datemove")
     dateEXPIRED = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "DateExpired")
     Else
     dateMOVE = Date
     dateEXPIRED = Date + 30
     End If
     
        If Int(dateEXPIRED - dateMOVE) > 0 And Int(dateEXPIRED - dateMOVE) <= 30 Then
        Else
           If MsgBox("Your trial version has been expired!" & Chr(10) & "Do you want to register your product?", vbYesNo + vbExclamation, "Registration") = vbYes Then
           frmLoginSHOW = True
           Else
           frmLoginSHOW = False
           Unload Me
           End If
        End If
Else
    If Split(Decrypt(Trim(GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "Reg"))), "angelitojason")(0) = "" Then
    Else
        If MsgBox("You Registration Has been corrupt do you want to register again?", vbYesNo + vbExclamation) = vbYes Then
        FRMREG.Show
        Unload Me
        Else
        Unload Me
        End If
    End If
End If
'--------------------------------------------
If frmLoginSHOW = True Then
FRMREG.Show 1
'---LOAD CONSUMERS LIST
sql = "SELECT id as 'Account Number',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name'  FROM iwadco_cons WHERE status<>'D'  ORDER BY lname+', '+fname+' '+mname"

lstview.lstDatabase sql, lstemp, 2
'-----
cmdNew.CaptionAlign = isbBottom
cmdedit.CaptionAlign = isbBottom
cmdDelete.CaptionAlign = isbBottom
cmdSearch.CaptionAlign = isbBottom
cmdRefresh.CaptionAlign = isbBottom
cmdcancel.CaptionAlign = isbBottom

cmdReading.CaptionAlign = isbBottom
cmdPayments.CaptionAlign = isbBottom
cmdReports.CaptionAlign = isbBottom
cmdUtilities.CaptionAlign = isbBottom
cmdDis.CaptionAlign = isbBottom

cmdNew.IconAlign = isbTop
cmdedit.IconAlign = isbTop
cmdDelete.IconAlign = isbTop
cmdSearch.IconAlign = isbTop
cmdRefresh.IconAlign = isbTop
cmdcancel.IconAlign = isbTop

cmdReading.IconAlign = isbTop
cmdPayments.IconAlign = isbTop
cmdDis.IconAlign = isbTop
cmdReports.IconAlign = isbTop
cmdUtilities.IconAlign = isbTop
frmShortCuts.Show

Else
frmLogin.Show 1
'---LOAD CONSUMERS LIST
sql = "SELECT id as 'Account Number',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name'  FROM iwadco_cons WHERE status <> 'D' ORDER BY lname+', '+fname+' '+mname"
'MsgBox sql
lstview.lstDatabase sql, lstemp, 2
'-----
cmdNew.CaptionAlign = isbBottom
cmdedit.CaptionAlign = isbBottom
cmdDelete.CaptionAlign = isbBottom
cmdSearch.CaptionAlign = isbBottom
cmdRefresh.CaptionAlign = isbBottom
cmdcancel.CaptionAlign = isbBottom
cmdDis.IconAlign = isbBottom
cmdReading.CaptionAlign = isbBottom
cmdPayments.CaptionAlign = isbBottom
cmdReports.CaptionAlign = isbBottom
cmdUtilities.CaptionAlign = isbBottom
cmdDis.CaptionAlign = isbBottom
cmdNew.IconAlign = isbTop
cmdedit.IconAlign = isbTop
cmdDelete.IconAlign = isbTop
cmdSearch.IconAlign = isbTop
cmdRefresh.IconAlign = isbTop
cmdcancel.IconAlign = isbTop

cmdReading.IconAlign = isbTop
cmdPayments.IconAlign = isbTop
cmdReports.IconAlign = isbTop
cmdUtilities.IconAlign = isbTop
cmdDis.IconAlign = isbTop
frmShortCuts.Show
End If
LoadCombo "SELECT lname+', '+fname+' '+mname as coor_name FROM iwadco_coor", Combo3
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
Me.lstemp.Height = Me.Height - Picture2.Height - StatusBar2.Height - 3500
StatusBar2.Top = lstemp.Top + lstemp.Height
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If MsgBox("Are you sure you want to terminate the system?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
    End
Else
    Cancel = 1
    frmShortCuts.Show
End If
End Sub

Private Sub mnuabout_Click()
frmShortCuts.Image1.Visible = False
frmAbout.Show 1
End Sub

Private Sub mnuArea_Click()
frmAddArea.Show 1
End Sub

Private Sub mnuBackup_Click()
Shell App.Path & "\backup.exe", vbNormalFocus
End
End Sub

Private Sub mnuCleanUp_Click()
If empType > 1 Then MsgBox "Access Denied!", vbExclamation, Me.Caption: Exit Sub
frmCleanUp.Show 1
End Sub

Private Sub mnuclose2_Click()
ActiveForm.clse
End Sub

Private Sub mnuCom_Click()
ChngPrinterOrientationLandscape Me
frmDate.Show 1
End Sub

Private Sub mnuCoor_Click()
frmCoordinator.Show
End Sub

Private Sub mnuCost_Click()
frmConsumer.Show
End Sub

Private Sub mnuCREATE_Click()
ActiveForm.ADD
End Sub

Private Sub mnudel_Click()
ActiveForm.del
End Sub


Private Sub mnuDetailed_Click()
ChngPrinterOrientationLandscape Me
frmDate2.Caption = "Aging"
    
frmDate2.Show 1
End Sub

Private Sub mnudis_Click()
frmDisconnect.Show
End Sub

Private Sub mnuedit_Click()
ActiveForm.edit
End Sub

Private Sub mnuEdit2_Click()
'Call edit
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHistory_Click()
ChngPrinterOrientationPortrait Me
Form3.Show 1
End Sub

Private Sub mnuLogout_Click()
Dim logoff As Variant
Dim oFrm As Form
logoff = "Are you sure you want to log-off " & EmpLogName & "?"
    If (MsgBox(logoff, vbExclamation + vbYesNo, "Log - Off User")) = vbYes Then
    Load frmLogin
    frmLogin.Show 1
    End If
End Sub

Private Sub mnuPay_Click()
Load frmpayments
frmpayments.Show
End Sub

Private Sub mnupayments_Click()
Load frmpayments
frmpayments.Show
End Sub

Private Sub mnuPaymentsRep_Click()
ChngPrinterOrientationLandscape Me
frmDate2.Show 1
End Sub

Private Sub mnuRead_Click()
frmBillingDate.Show 1
End Sub

Private Sub mnureading_Click()
Unload frmBillingDate
frmBillingDate.Show 1
End Sub

Private Sub mnuReadingsRep_Click()
ChngPrinterOrientationLandscape Me
Load frmSelCoorRep
frmSelCoorRep.Caption = "Reading"
frmSelCoorRep.Show 1
If aa1 = "" Or aa2 = "" Then Exit Sub
rptReadings.DataMember = "cmdReadingsTots"
rptReadings.Sections("Section4").Controls("lblDate").Caption = aa1 & " " & aa2
rptReadings.Sections("Section4").Controls("lblcoorname").Caption = frmSelCoorRep.Lstcoor.SelectedItem.SubItems(1) & "( Area - " & asdasd & " ) "
aa1 = ""
aa2 = ""
aa3 = -1
End Sub

Private Sub mnurefresh_Click()
ActiveForm.Ref
End Sub

Private Sub MnuRefresh2_Click()
sql = "SELECT id as 'Account Number',iwadco_cons.lname+', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname as 'Name'  FROM iwadco_cons WHERE status <>'D'  ORDER BY lname+', '+fname+' '+mname"
StatusBar2.Panels(2).Text = lstview.lstDatabase(sql, lstemp, 2)
End Sub

Private Sub mnuregister_Click()
FRMREG.Show 1
End Sub

Private Sub mnuRepAll_Click()
ChngPrinterOrientationPortrait Me
Load frmSelCoorRep
frmSelCoorRep.Caption = "SOA - All"
frmSelCoorRep.Show 1
End Sub

Private Sub mnuRepCons_Click()
frmDate4.Show 1
End Sub

Private Sub mnuRepCoor_Click()
ChngPrinterOrientationPortrait Me
rptCoor.Show
End Sub

Private Sub mnuRepIndi_Click()
ChngPrinterOrientationPortrait Me
frmSelAccNum.Show 1
End Sub

Private Sub mnuRForm_Click()
ChngPrinterOrientationPortrait Me
Load frmSelCoorRep
frmSelCoorRep.Caption = "Reading Form"
frmSelCoorRep.Show 1
'.Show
End Sub

Private Sub mnusearch_Click()
FRMsearch.Show 1
End Sub

Private Sub mnuSum_Click()
'sql = "SELECT iwadco_cons.lname + ', ' + iwadco_cons.fname + ' ' + iwadco_cons.mname AS 'Account Name', iwadco_cons.id AS 'Account No', iwadco_cons.address AS Address, iwadco_coor.lname + ', ' + iwadco_coor.fname + ' ' + iwadco_coor.mname AS Coordinato, iwadco_typcon.type AS 'Type Of Connection', iwadco_readings.billfrom, iwadco_readings.billto, iwadco_readings.due_date AS 'Due Date', iwadco_readings.previous_reading AS Previous, iwadco_readings.present_reading AS Present, iwadco_readings.consume AS 'Total Used', iwadco_readings.excess AS Excess, iwadco_readings.amount_excess AS 'Amount Excess',iwadco_typcon.min_rate, iwadco_readings.arrears AS Arrears, iwadco_readings.total_amount AS 'Total Amount Due', iwadco_readings.wtax,iwadco_typcon.min_rate+iwadco_readings.amount_excess as 'Total Amount' " & _
'"FROM iwadco_cons INNER JOIN iwadco_coor ON iwadco_cons.coor_id = iwadco_coor.id INNER JOIN iwadco_typcon ON iwadco_cons.class = iwadco_typcon.id INNER JOIN iwadco_readings ON iwadco_cons.id = iwadco_readings.account_no group by iwadco_coor.lname,iwadco_coor.fname,iwadco_coor.mname"
'DataEnvironment1.rscmdSOA.Open sql
Load frmSelCoorRep
frmSelCoorRep.Caption = "SOA Summary"
frmSelCoorRep.Show 1
End Sub

Private Sub mnuSumCom_Click()
ChngPrinterOrientationLandscape Me
frmDate6.Show 1
End Sub

Private Sub mnuSumRep_Click()
frmdate5.Show 1
End Sub

Private Sub mnuSysUser_Click()
frmSysUser.Show 1
End Sub

Private Sub mnuTap_Click()
frmdate3.Show 1
End Sub

Private Sub mnuULedit_Click()
If user_priv("update") = True Then
    FRM_UTL_CON_MIN_RATE.Show 1
Else
    MsgBox "Access Denied!", vbExclamation, "Security"
    Exit Sub
End If
End Sub

Private Sub mnuPCM_Click()
If user_priv("update") = True Then
    FRM_UTL_ONEXCESS_MIN.Show 1
Else
    MsgBox "Access Denied!", vbExclamation, "Security"
    Exit Sub
End If
End Sub

Private Sub mnuDS_Click()
If user_priv("update") = True Then
    FRM_ULT_DEFAULT.Show 1
Else
    MsgBox "Access Denied!", vbExclamation, "Security"
    Exit Sub
End If
End Sub

Private Sub mnuvbr_Click()
Load frmViewRecords
frmViewRecords.Show
End Sub

Private Sub mnuVP_Click()
Unload frmAddConsumer
FormShow frmAddConsumer, True
End Sub

Private Sub Picture1_Click()
'rptReceipt.Show
End Sub

Private Sub Timer1_Timer()
IV = IV + 1
If IV = 1 Then
StatusBar1.Panels.Item(3).Text = ""
ElseIf IV = 2 Then
StatusBar1.Panels.Item(3).Text = " " & EmpLogName
IV = 0
End If
End Sub

Private Sub txtSearch_Change()
    'txtSearch.Text = str_Filter(txtSearch, 48, 57, 45)

End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call isButton1_Click
End If
End Sub
