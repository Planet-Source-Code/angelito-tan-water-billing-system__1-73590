VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "JTGroup"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Enter"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5400
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5400
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Serial Number:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Trial Version has expired!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
