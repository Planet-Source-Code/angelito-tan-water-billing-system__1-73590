VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   18
      Left            =   5640
      Top             =   2520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Decrypt"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Encrypt"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Decrypt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Encrypt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Text2.Text = ENCRYPT(Text1.Text)

End Sub

Private Sub Command2_Click()
Text3.Text = Decrypt(Text2.Text)
End Sub

Private Sub Command3_Click()
Timer1.Enabled = True
End Sub

Private Sub Text2_Change()
Dim x, y, z As Double
Dim tmparray() As String


 tmparray() = Split(Text2.Text, ":")
 
For x = 1 To UBound(tmparray)
    If x > 0 Then
    Debug.Print tmparray(x)
    End If
Next

End Sub

Private Sub Timer1_Timer()
Call Command1_Click
End Sub
