strComputer = "."
strTargetAddress = "192.168.29.78"

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objItem in colItems
    arrIPAddresses = objItem.IPAddress
    For Each strAddress in arrIPAddresses
        If strAddress = strTargetAddress Then
            strMACAddress = objItem.MacAddress
        End If
    Next
Next

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapter Where MACAddress = '" & strMACAddress & "'")

For Each objItem in colItems
    If Not IsNull(objItem.NetConnectionID) Then
        Wscript.Echo objItem.NetConnectionID
    End If
Next

