VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Monthly Bill"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtminAmount 
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
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin Project1.isButton cmdsave 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      icon            =   "Form2.frx":0000
      style           =   5
      caption         =   "&Save"
      captionalign    =   2
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "Form2.frx":15174
   End
   Begin Project1.isButton cmdcancel 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      icon            =   "Form2.frx":1519C
      style           =   5
      caption         =   "&Cancel"
      iconsize        =   20
      iconalign       =   1
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      ttforecolor     =   0
      font            =   "Form2.frx":2C238
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Monthly Bill:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

