 ' *  VBScript to roll back the IE 
 ' *  Source: http://www.itninja.com/question/un-install-internet-explerer10-silently-and-have-it-default-back-to-ie-9-or-ie-8 
 ' *  /Björn J.
 ' *  2013-05-03
 
 Option Explicit
  ' *  
 
 Dim updateID
 
 updateID = "proxy.trafalgar.jm:8080"
 
 ' *  Application Title
 
 Dim Title
 Title = "IE Changer"
 
 ' *  Welcome Message
 
 Dim Welcome_Text
 Welcome_Text = "Do you really want toremove" & updateID & "?"
 
  Call Welcome()
 
 Sub Welcome()
     Dim GO
     GO = MsgBox(Welcome_Text, 36, Title)
     If GO = 7 Then
         WScript.Quit
     End If
 End Sub
 
  Dim WSHShell
  Set WSHShell = WScript.CreateObject("WScript.Shell")
 
 Dim confirm
 confirm = MsgBox("ie version .....  Enjoy.", 64, Title)
 
 ' *  Open IE
 
 WSHShell.run "iexplore.exe http://www.google.com"