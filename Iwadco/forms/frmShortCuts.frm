VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShortCuts 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Shortcuts"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   11295
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortCuts.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   600
      Picture         =   "frmShortCuts.frx":52D2
      Top             =   2640
      Width           =   7500
   End
End
Attribute VB_Name = "frmShortCuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
Me.WindowState = 2
End Sub

Private Sub form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Beep: Cancel = 1
End Sub

Private Sub Form_Resize()
Image1.Left = (Me.Width / 2) - (Image1.Width / 2)
Exit Sub
End Sub

