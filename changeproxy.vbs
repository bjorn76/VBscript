' *  VBScript to change the IE Proxy Server, by directly editing the registry
 ' *  Version:  2.1
 ' *  
 ' *  Enjoy!
 
 Option Explicit
 
 ' *  Proxy Server
 
 Dim proxy
 'proxy = "proxy.csd.royston.jm:8080"
 proxy = "proxy.trafalgar.jm:8080"
 
 ' *  Application Title
 
 Dim Title
 Title = "Los Proxy Changer"
 
 ' *  Welcome Message
 
 Dim Welcome_Text
 Welcome_Text = "Do you really want to change the Proxy Server to " & proxy & "?"
 
 Call Welcome()
 
 Sub Welcome()
     Dim GO
     GO = MsgBox(Welcome_Text, 36, Title)
     If GO = 7 Then
         WScript.Quit
     End If
 End Sub
 
 ' *  Warning Message
 
 Dim Warning_Text
 Warning_Text = "Warning:  All current Internet Explorer and Windows Explorer windows will be closed." & Chr(10) _
                 & Chr(10) _
                 & "Do you still wish to continue?"
 
 Call Warning()
 
 Sub Warning()
     Dim GO
     GO = MsgBox(Warning_Text, 36, Title)
     If GO = 7 Then
         WScript.Quit
     End If
 End Sub
 
 ' *  WSHShell
 
 Dim WSHShell
 Set WSHShell = WScript.CreateObject("WScript.Shell")
 
 ' *  Kill IE (and consequently WE as well)
 
 Call Kill_IE()
 
 Sub Kill_IE()
     While WSHShell.AppActivate("Internet Explorer")
         WSHShell.SendKeys "%{F4}"
         WScript.Sleep 500    ' Impedes a traffic jam
     WEnd
 End Sub
 
 ' *  Regedits
 
 WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"
 WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", proxy
 WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride", ""
 WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\DisablePasswordCaching", 1, "REG_DWORD"
 
 ' *  Confirmation
 
 Dim confirm
 confirm = MsgBox("Proxy Server has been changed to " & proxy & ".  Enjoy.", 64, Title)
 
 ' *  Open IE
 
 WSHShell.run "iexplore.exe http://www.google.com"