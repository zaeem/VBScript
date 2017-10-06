
strComputer = "." 
GB = 1024 *1024 * 1024
     Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
     Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
     For Each objItem in colItems
          Manufacturer = objItem.Manufacturer
          Model = objItem.Model
          PhysicalMemory =  Round(objItem.TotalPhysicalMemory/GB,3) 
     next


Set objWMIService = GetObject("winmgmts:" _ 
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 

Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
  
Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem") 
For Each objOS in colOSes 
  ComputerName =  objOS.CSName 
  OSname = objOS.Caption 'Name 
  OSVersion = objOS.Version 
  OSSerialno = objOS.SerialNumber
  dtmConvertedDate.Value = objOS.InstallDate
    dtmInstallDate = dtmConvertedDate.GetVarDate
    
Next 

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
              Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk WHERE DriveType=3")
        For Each objItem in colItems
 
             DIM pctFreeSpace,strFreeSpace,strusedSpace
 
             pctFreeSpace = INT((objItem.FreeSpace / objItem.Size) * 1000)/10 
       strDiskSize = Int(objItem.Size /1073741824) & "Gb"
       strFreeSpace = Int(objItem.FreeSpace /1073741824) & "Gb"
       strUsedSpace = Int((objItem.Size-objItem.FreeSpace)/1073741824) & "Gb"

Next

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")

For Each objItem in colItems
ProcessorManufacturer =  objItem.Manufacturer
ProcessorID = objItem.ProcessorId
NoOfProcessor =  objItem.NumberOfCores
ProcessorVersion = objItem.Version
Next


Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
For Each objItem in colItems
MACAddress =  objItem.MACAddress
Next

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
     Set colItems = objWMIService.ExecQuery("Select * from Win32_BIOS")
     For Each objItem in colItems
          pc_serial_no = objItem.SerialNumber 
     next

dim NIC1, Nic, StrIP, CompName

Set NIC1 =     GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

For Each Nic in NIC1

    if Nic.IPEnabled then
        StrIP = Nic.IPAddress(i)
    End if
Next


Dim objXmlHttpMain , URL


strJSONToSend = "displayName="&ComputerName&"&secretKey="&secret_id&"&macAddress="&MACAddress&"&assetType=Desktop&processorCores="&NoOfProcessor&"&pcSerialNo="&pc_serial_no&"&processorName="&ProcessorManufacturer&"&memory="&PhysicalMemory&"&osName="&OSName&"&osVersion="&OSVersion&"&osSerialNo="&OSSerialno&"&windowsInstallationDate="&dtmInstallDate&"&ipAddress="&StrIP&"&diskSpace="&strDiskSize&"&manufacturer="&Manufacturer&"&Model="&Model



URL="http://secure.nulodgic-staging.com/asset_collections/host_asset" 
Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP") 
on error resume next 
objXmlHttpMain.open "POST",URL, False 
objXmlHttpMain.setRequestHeader "Authorization", "Basic"
objXmlHttpMain.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"


objXmlHttpMain.send (strJSONToSend)
set objJSONDoc = nothing 
set objResult = nothing


WScript.Echo "Data Sent"
