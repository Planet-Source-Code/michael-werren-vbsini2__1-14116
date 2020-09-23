Attribute VB_Name = "Basic"
Public psAntwort As String

Sub GetINI(Sektion As String, Key As String, Datei As String)
  Dim lsAntwort As String
  Dim liX As Integer
  lsAntwort = String(255, " ")
  liX = GetPrivateProfileString(Sektion, Key, "", lsAntwort, 255, Datei)
  lsAntwort = Trim(Left(lsAntwort, liX))
  psAntwort = lsAntwort
End Sub

Sub GetSections(Datei As String)
  Dim lsAntwort As String
  Dim lsReturn As String
  Dim lsZeichen As String
  
  Dim liX As Integer
  Dim liY As Integer
  lsReturn = ""
  lsAntwort = String(4096, " ")
  liX = GetPrivateProfileString(0&, 0&, "", lsAntwort, 4096, Datei)
  For liY = 1 To liX
    lsZeichen = Mid(lsAntwort, liY, 1)
    If Asc(lsZeichen) = 0 Then
      lsReturn = lsReturn + ","
    Else
      lsReturn = lsReturn + lsZeichen
    End If
  Next liY
  If Right(lsReturn, 1) = "," Then lsReturn = Left(lsReturn, Len(lsReturn) - 1)
  psAntwort = Trim(lsReturn)
End Sub

Sub GetKeys(Sektion As String, Datei As String)
  Dim lsAntwort As String
  Dim lsReturn As String
  Dim liX As Integer
  
  lsReturn = ""
  lsAntwort = String(4096, " ")
  liX = GetPrivateProfileString(Sektion, 0&, "", lsAntwort, 4096, Datei)
  For liY = 1 To liX
    lsZeichen = Mid(lsAntwort, liY, 1)
    If Asc(lsZeichen) = 0 Then
      lsReturn = lsReturn + ","
    Else
      lsReturn = lsReturn + lsZeichen
    End If
  Next liY
  If Right(lsReturn, 1) = "," Then lsReturn = Left(lsReturn, Len(lsReturn) - 1)
  psAntwort = Trim(lsReturn)
  
End Sub

Sub WriteINI(Sektion As String, Key As String, Wert As String, Datei As String)
  Dim liX As Integer
  liX = WritePrivateProfileString(Sektion, Key, Wert, Datei)
End Sub

Sub INIDelKey(Sektion As String, Key As String, Datei As String)
  Dim liX As Integer
  liX = WritePrivateProfileString(Sektion, Key, 0&, Datei)
End Sub

Sub INIDelSection(Sektion As String, Datei As String)
  Dim liX As Integer
  liX = WritePrivateProfileString(Sektion, 0&, 0&, Datei)
End Sub

Sub AppWait(liTime As Long)
  Sleep liTime
End Sub
