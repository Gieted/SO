If Not Wscript.Arguments.Named.Exists("elevate") Then
  Call CreateObject("Shell.Application").ShellExecute(Wscript.FullName, """" & Wscript.ScriptFullName & """ /elevate", "", "runas", 1)
  Wscript.Quit()
End If


strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")                                                                      
Set networkAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter")

list = "Network Adapters:" + vbCrlf

For Each adapter in networkAdapters
  id = adapter.Index + 1
	list = list & vbCrlf & id & ". " & adapter.name
Next

selectedAdapterId = InputBox(list & vbCrlf & vbCrlf & "Please select a network adapter you want to modify:")

If IsEmpty(selectedAdapterId) Then
  Wscript.Quit()
End If

selectedAdapterIndex = selectedAdapterId - 1
selectedAdapterName = networkAdapters.ItemIndex(selectedAdapterIndex).name
Set adapterConfiguration = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration where Index=" & selectedAdapterIndex).ItemIndex(0)

noneString = "None"

ip = null
If NOT IsNull(adapterConfiguration.IPAddress) Then 
  ip = adapterConfiguration.IPAddress(0) 
Else
  ip = noneString
End If

subnetMask = null
If NOT IsNull(adapterConfiguration.IPSubNet) Then 
  subnetMask = adapterConfiguration.IPSubNet(0) 
Else
  subnetMask = noneString
End If

defaultGateway = null
If NOT IsNull(adapterConfiguration.DefaultIpGateway) Then 
  defaultGateway = adapterConfiguration.DefaultIpGateway(0) 
Else
  defaultGateway = noneString
End If

dns = null
If NOT IsNull(adapterConfiguration.DNSServerSearchOrder) Then 
  dns = adapterConfiguration.DNSServerSearchOrder(0) 
Else
  dns = noneString
End If

Function handleResult(result)
  If result = 0 Then
    MsgBox("The operation completed successfully.")
  Else
    MsgBox("The operation failed with the following error: " & result)
  End If
End Function

choice = InputBox(selectedAdapterName & vbCrlf & vbCrlf _
  & "IP Address: " & ip & vbCrlf _
  & "Subnet mask: " & subnetMask & vbCrlf _
  & "Default gateway: " & defaultGateway & vbCrlf _
  & "DNS: " & DNS & vbCrlf & vbCrlf _
  & "Select what you want to change:" & vbCrlf _
  & "1. IP Address" & vbCrlf _
  & "2. Subnet mask" & vbCrlf _
  & "3. Default gateway" & vbCrlf _
  & "4. DNS")

Select Case choice
  Case 1
    newIp = InputBox("Please enter the IP address you want to set:")

    If IsEmpty(newIp) Then
      Wscript.Quit()
    End If

    result = adapterConfiguration.EnableStatic(Array(newIp), Array(subnetMask))
    handleResult(result)

  Case 2
    newSubnetMask = InputBox("Please enter the subnet mask you want to set:")

    If IsEmpty(newSubnetMask) Then
      Wscript.Quit()
    End If

    result = adapterConfiguration.EnableStatic(Array(ip), Array(newSubnetMask))
    handleResult(result)

  Case 3
    newDefaultGateway = InputBox("Please enter the default gateway you want to set:")

    If IsEmpty(newDefaultGateway) Then
      Wscript.Quit()
    End If

    gatewayMetric = Array(1)
    result = adapterConfiguration.SetGateways(Array(newDefaultGateway), gatewayMetric)
    handleResult(result)
  
  Case 4
    newDns = InputBox("Please enter the DNS you want to set:")

    If IsEmpty(newDns) Then
      Wscript.Quit()
    End If

    result = adapterConfiguration.SetDNSServerSearchOrder(Array(newDns))
    handleResult(result)
End Select
