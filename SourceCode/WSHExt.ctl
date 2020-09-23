VERSION 5.00
Begin VB.UserControl WSHExt 
   CanGetFocus     =   0   'False
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   Picture         =   "WSHExt.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   495
End
Attribute VB_Name = "WSHExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function WSHAbout(lsMeldung As String, lsTitel As String) As String
  frmAbout.Label1.Caption = lsMeldung
  frmAbout.Caption = lsTitel
  frmAbout.Show vbModal
  WSHAbout = psAntwort
End Function

' Liest einen Eintrag einer INI
Public Function WSHGetINI(lsSection As String, lsKey As String, lsDatei As String) As String
  GetINI lsSection, lsKey, lsDatei
  WSHGetINI = psAntwort
End Function

' Ermittelt die Sektionen einer INI
Public Function WSHGetSections(lsDatei As String) As String
  GetSections lsDatei
  WSHGetSections = psAntwort
End Function

' Ermittelt die Keys einer Sektion
Public Function WSHGetKeys(lsSection As String, lsDatei As String) As String
  GetKeys lsSection, lsDatei
  WSHGetKeys = psAntwort
End Function

' Schreibt ein Eintrag in die INI Datei
Public Sub WSHWriteINI(lsSection As String, lsKey As String, lsWert As String, lsDatei As String)
  WriteINI lsSection, lsKey, lsWert, lsDatei
End Sub

' Löscht einen Key in der INI Datei
Public Sub WSHINIDelKey(lsSection As String, lsKey As String, lsDatei As String)
  INIDelKey lsSection, lsKey, lsDatei
End Sub

' Löscht ein Kapitel in der INI Datei
Public Sub WSHINIDelSection(lsSection As String, lsDatei As String)
  INIDelSection lsSection, lsDatei
End Sub

' Weist die Applikation an liTime  (Millisekunden) zu warten
Public Sub WSHSleep(liTime As Long)
  AppWait liTime
End Sub

