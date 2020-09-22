VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   -60
   ClientTop       =   30
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "RECEIVED from"
      Height          =   435
      Left            =   3480
      TabIndex        =   5
      Top             =   2040
      Width           =   6675
   End
   Begin VB.Label Label5 
      Caption         =   "RECEIVED from"
      Height          =   195
      Left            =   3480
      TabIndex        =   4
      Top             =   1800
      Width           =   6675
   End
   Begin VB.Label Label4 
      Caption         =   "RECEIVED from"
      Height          =   195
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   6675
   End
   Begin VB.Label Label3 
      Caption         =   "RECEIVED from"
      Height          =   195
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   6675
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   9600
      TabIndex        =   0
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.PrintForm
End Sub
