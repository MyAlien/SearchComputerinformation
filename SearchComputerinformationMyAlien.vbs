on error resume next

' old strComputer = "HOSTNAME.domain.tld"

strComputer = InputBox("Conti Help -> Enter Computer name", _
"Conti Help - Search Computer Information -", strComputer)

Info1 = ""
Info2 = ""
Info3 = ""
Info4 = ""

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colBIOS = objWMIService.ExecQuery ("Select * from Win32_BIOS")

For each objBIOS in colBIOS
    Info1 =  "Manufacturer: " & objBIOS.Manufacturer
    Info2 =  "Serial Number: " & objBIOS.SerialNumber
Next

Set colSystem = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")

For each objItem in colSystem
    'Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Info3 =  "Model: " &  objItem.Model
Next


Set colAdapters = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
 
 
For Each objAdapter in colAdapters
   
   Info4 = Info4 &  "Hostname:" & objAdapter.DNSHostName & ", "
   Info4 = Info4 &  "MAC:" & objAdapter.MACAddress & ", "
   Info6 = Info6 &  "MAC:" & objProcess.GetOwner & ", "

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
   Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='explorer.exe'")
   For Each objProcess in colProcessList
  objProcess.GetOwner strNameOfUser, strUserDomain
  strOwner = strUserDomain & "\" & strNameOfUser
   Info7 = Info7 &  "UID :" & strOwner & ", "
    Next

       Info5 = Info5 &  "Latest Update 10/01/2020 MyAlien"
 
   If Not IsNull(objAdapter.IPAddress) Then
	  Nb=0
      For i = 0 To UBound(objAdapter.IPAddress)
        if Nb = 0 then 
			Info4 = Info4 &  "IP address: " & objAdapter.IPAddress(i)
			nb = nb + 1
		end if
      Next
   End If
   Info4 = Info4 & chr(13)
next


' Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
' Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='explorer.exe'")

' For Each objProcess in colProcessList
'  objProcess.GetOwner strNameOfUser, strUserDomain
'   WScript.Echo "Latest Users LogOn" & vbNewLine
'  strOwner = strUserDomain & "\" & strNameOfUser
'  WScript.Echo strOwner
' Next



'Msgbox Info1 & chr(13) & Info2 & chr(13) & Info3 & chr(13) & chr(13) & Info4 & chr(13) & Info5
r= MsgBox(Info1 & chr(13) & Info3 & chr(13) & Info2 & chr(13) & chr(13) & Info4 & chr(13) & Info7 & chr(13) & Info5, vbOKOnly + vbInformation, "System Informations |") 
