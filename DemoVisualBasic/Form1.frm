VERSION 5.00
Object = "{C1F0FE8A-7875-4451-9ED3-A1A669AE5AF0}#1.0#0"; "WSHWIControl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   360
      Width           =   4455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete a section"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Delete key"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Keys of a section"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "All sections"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Write INI"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read INI"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   4695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   960
      Width           =   2535
   End
   Begin WSHWIControl.WSHExt WSHExt1 
      Height          =   495
      Left            =   2760
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Input/Output"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "c:\temp\test.ini"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objINIFile
Sub readthefile()
  Dim liFileNr As Integer
  Dim lsLine As String

  Text1.Text = ""
  On Error GoTo NoFile
    liFileNr = FreeFile
    Open "c:\temp\test.ini" For Input As #liFileNr
    Do While Not EOF(liFileNr)
      Line Input #liFileNr, lsLine
      Text1.Text = Text1.Text + lsLine + vbCrLf
    Loop
    Close #liFileNr
  GoTo Cont
NoFile:
  MsgBox "Missing c:\temp\test.ini File, please copy file to this folder"
  GoTo Cont
Cont:
  On Error GoTo 0

End Sub

'Read from a file
Private Sub Command1_Click()
  Set objINIFile = CreateObject("WSHWIControl.WSHExt")
    Text2.Text = objINIFile.WSHGetINI("Test", "Name5", "c:\temp\test.ini")
  Set objINIFile = Nothing
  
  readthefile
End Sub

' Write to a File
Private Sub Command2_Click()
  Set objINIFile = CreateObject("WSHWIControl.WSHExt")
    objINIFile.WSHWriteINI "Test", "Name5", Text2.Text, "c:\temp\test.ini"
  Set objINIFile = Nothing
  
  readthefile
End Sub

' Show all Sections of a file
Private Sub Command3_Click()
  Set objINIFile = CreateObject("WSHWIControl.WSHExt")
    Text2.Text = objINIFile.WSHGetSections("c:\temp\test.ini")
  Set objINIFile = Nothing

  readthefile
End Sub

Private Sub Command4_Click()
  Set objINIFile = CreateObject("WSHWIControl.WSHExt")
    Text2.Text = objINIFile.WSHGetKeys("Test", "c:\temp\test.ini")
  Set objINIFile = Nothing

  readthefile
End Sub

Private Sub Command5_Click()
  Set objINIFile = CreateObject("WSHWIControl.WSHExt")
    objINIFile.WSHINIDelKey "Test", "Name2", "c:\temp\test.ini"
  Set objINIFile = Nothing

  readthefile
End Sub

Private Sub Command6_Click()
  Set objINIFile = CreateObject("WSHWIControl.WSHExt")
    objINIFile.WSHINIDelSection "Test2", "c:\temp\test.ini"
  Set objINIFile = Nothing

  readthefile
End Sub

Private Sub Form_Load()
  readthefile
End Sub
