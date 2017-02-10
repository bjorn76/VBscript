Option Explicit

'Input var
Dim strComputer
Dim strTargetAddress
Dim arrDNSServers

'Output var
Dim strNetConnectionID
Dim colItems

'Temp var
Dim objWMIService
Dim arrIPAddresses
Dim strAddress
Dim objItem
Dim strMACAddress
Dim strTargetIP




'Code
'arrDNSServers = Array("192.168.1.100", "192.168.1.200")
arrDNSServers = Array("172.30.4.89", "172.30.4.90")

'WScript.Echo ListIPAddresses()
strTargetIP = ListIPAddresses()
strTargetIP = InputBox("Enter IP","Change DNS Search order",strTargetIP)
'WScript.Echo ChangeDNSorder(strTargetIP)
strMACAddress = ChangeDNSorder(strTargetIP)
wscript.Echo "Netoconid(MAC): " & NetConID(strMACAddress) & " (" &strMACAddress  & ")"

Function ChangeDNSorder(strTargetIP)

'Init inputs
strComputer = "."
'strTargetAddress = "192.168.29.78"

If strTargetIP = "" Then 
  Exit Function
End If
  
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objItem in colItems
    arrIPAddresses = objItem.IPAddress
    For Each strAddress in arrIPAddresses
        If strAddress = strTargetIP Then
            'obj.Get
            objItem.SetDNSServerSearchOrder(arrDNSServers)
            strMACAddress = objItem.MacAddress
            ChangeDNSorder = strMACAddress
        End If
    Next
Next

End Function



Function NetConID (strMACAddress)

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapter Where MACAddress = '" & strMACAddress & "'")

For Each objItem in colItems
    If Not IsNull(objItem.NetConnectionID) Then
        'Wscript.Echo objItem.NetConnectionID
        NetConID = objItem.NetConnectionID
    End If
Next
End Function



Function ListIPAddresses
' List IP Addresses for a Computer

' Windows Server 2003 : Yes
' Windows XP : Yes
' Windows 2000 : Yes
' Windows NT 4.0 : Yes
' Windows 98 : Yes


Dim objWMIService, IPConfigSet, IPConfig, i

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set IPConfigSet = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
 
For Each IPConfig in IPConfigSet
    If Not IsNull(IPConfig.IPAddress) Then 
        For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
            'WScript.Echo IPConfig.IPAddress(i)
            ListIPAddresses = ListIPAddresses & IPConfig.IPAddress(i) & ";"
        Next
    End If
Next
End Function



