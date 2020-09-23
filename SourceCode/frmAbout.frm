VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "INI OCX for WSH, VBA, VB by Mike Werren"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  psAntwort = frmAbout.Label1.Caption
  Unload frmAbout
End Sub


Private Sub Timer1_Timer()
  Call cmdOK_Click
End Sub
