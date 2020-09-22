VERSION 5.00
Begin VB.Form FRMREG 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Register"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3825
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdreg 
      Caption         =   "Register"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtreg 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "FRMREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DecVar As Variant
Private Sub cmdreg_Click()
'MsgBox Decrypt(GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Exp", "Reg"))
If Len(txtreg.Text) > 0 Then
    If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "Reg") = "ERROR" Then
        If Decrypt(txtreg.Text) = "angelitojason" Then
        CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft"
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "Reg", Trim(txtreg.Text)
        MsgBox "You registration is complete", vbExclamation
        Unload Me
        frmLogin.Show 1
        Exit Sub
        Else
        MsgBox "Invalid Registration Code", vbExclamation
        End If
    Else
        If Decrypt(txtreg.Text) = "angelitojason" Then
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Kernel32Soft", "Reg", Trim(txtreg.Text)
        MsgBox "You registration is complete", vbExclamation
        Unload Me
        frmLogin.Show 1
        Exit Sub
        Else
        MsgBox "Invalid Registration Code", vbExclamation
        End If
    End If
End If
''@º¹³É¶‚ú—¢’
'-ñ=¤‚6¯ÀâßÅ2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Set FRMREG = Nothing
End Sub
