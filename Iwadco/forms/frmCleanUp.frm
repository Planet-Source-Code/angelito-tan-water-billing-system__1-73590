VERSION 5.00
Begin VB.Form frmCleanUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clean-up Database"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4575
      Begin VB.CheckBox Check7 
         Caption         =   "Tapping Fees"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Coordinators Area Record"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Coordinators Record"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Commisions"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Readings"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Consumers Record"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin Project1.isButton cmdsave 
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   1800
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
         icon            =   "frmCleanUp.frx":0000
         style           =   5
         caption         =   "&Clean"
         iconsize        =   15
         captionalign    =   2
         iconalign       =   1
         inonthemestyle  =   0
         tooltiptitle    =   ""
         tooltipicon     =   0
         tooltiptype     =   0
         ttforecolor     =   0
         font            =   "frmCleanUp.frx":6864
      End
      Begin Project1.isButton cmdclose 
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   1800
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
         icon            =   "frmCleanUp.frx":688C
         style           =   5
         caption         =   "&Close"
         iconsize        =   18
         captionalign    =   2
         iconalign       =   1
         inonthemestyle  =   0
         tooltiptitle    =   ""
         tooltipicon     =   0
         tooltiptype     =   0
         ttforecolor     =   0
         font            =   "frmCleanUp.frx":1D928
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clean-up Database"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   3825
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Clean Transactions And Consumer"
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
      Height          =   690
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   120
      Picture         =   "frmCleanUp.frx":1D950
      Top             =   120
      Width           =   585
   End
End
Attribute VB_Name = "frmCleanUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
On Error GoTo err
Dim msg As String
msg = ""
If Check3.Value = 1 Then
    rsCheck
    rs.Open "DELETE FROM iwadco_payments", CN, adOpenStatic, adLockOptimistic
    msg = Chr(10) & "Payments"
End If
If Check4.Value = 1 Then
    rsCheck
    rs.Open "DELETE FROM iwadco_commisions", CN, adOpenStatic, adLockOptimistic
    msg = msg & Chr(10) & "Commisions"
End If
If Check2.Value = 1 Then
    rsCheck
    rs.Open "DELETE FROM iwadco_readings", CN, adOpenStatic, adLockOptimistic
    msg = msg & Chr(10) & "Readings"
End If
If Check7.Value = 1 Then
    rsCheck
    rs.Open "DELETE FROM iwadco_tappingfee", CN, adOpenStatic, adLockOptimistic
    msg = msg & Chr(10) & "Tapping Fee"
End If
If Check1.Value = 1 Then
    rsCheck
    rs.Open "DELETE FROM iwadco_cons", CN, adOpenStatic, adLockOptimistic
    msg = msg & Chr(10) & "Consumer"
End If
If Check6.Value = 1 Then
    rsCheck
    rs.Open "DELETE FROM iwadco_area", CN, adOpenStatic, adLockOptimistic
    msg = msg & Chr(10) & "Area"
End If
If Check5.Value = 1 Then
    rsCheck
    rs.Open "DELETE FROM iwadco_coor", CN, adOpenStatic, adLockOptimistic
    msg = msg & Chr(10) & "Coordinator"
End If
If msg <> "" Then
MsgBox "The following tables has been clean!" & msg
End If
err:
Select Case err
Case 0
Case Else
    MsgBox err.Description, vbCritical, err.Number
    Exit Sub
End Select
End Sub

