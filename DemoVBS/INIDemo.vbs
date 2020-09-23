' This Demo shows the how to access INI Files with the WSHWIContol.ocx
' The WSHWIContol.ocx is FreeWare for all WSH (vbscript and jscript) users
' Coded by Michael Werren, mike@werren.com

' Note:
' 1. Start the registerocx.bat for registration the ocx file
' 2. For testing  the script copy the file test.ini to c:\temp, have a look to the 
'    File before, during and after executing the Code

' Waranty
' It's your own risk to work with this module


option Explicit

Dim ObjAdr
Dim Antwort

' Declaration of the ActiveX Control
  set ObjAdr = WScript.CreateObject("WSHWIControl.WSHExt")


antwort = objadr.wshabout("The Demo with WSHWIControl","Have fun")

  
' Read a key from the INI file
  Antwort = objadr.WSHGetINI("Test1","Name5","c:\temp\test.ini")
  WScript.Echo "I have read this key form the INI file:" + chr(13) + Antwort

' Write to the INI file
  objadr.WSHWriteINI "Test","Name5","Mr. Thomas Miller","c:\temp\test.ini"
  WSCript.Echo "I have write 'Mr. Thomas Miller' to the INI File"
 
' Show all sections of the INI file
  Antwort = objAdr.WSHGetSections("c:\temp\test.ini")
  WScript.Echo "Those sections are in the INI file:" + chr(13) + Antwort

' Take a sleep about 1 second
  objAdr.WSHSleep 1000

' Show all keys of a section
  Antwort = objAdr.WSHGetKeys("Test","c:\temp\test.ini")
  WScript.Echo "This are the keys of the section Test:" + chr(13) + Antwort

' Clear a key in a section of a INI file
  objadr.WSHINIDelKey "Test","Name2","c:\temp\test.ini"

' Clear a section in a INI file
  objadr.WSHINIDelSection "Test2","c:\temp\test.ini"

' Clear the memory of the Object and exit
  set ObjAdr = Nothing
  WScript.Quit()