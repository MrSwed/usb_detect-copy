

Option Explicit

Const wbemFlagReturnImmediately = &H00000010
Const wbemFlagForwardOnly       = &H00000020

Dim objSWbemServicesEx
Dim collSWbemObjectSet
Dim objSWbemObjectEx_Win32_DiskDrive, objSWbemObjectEx_Win32_DiskPartition, objSWbemObjectEx_Win32_LogicalDisk
Dim oShell,oExec,DiskDrive

Set objSWbemServicesEx = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'Set collSWbemObjectSet = objSWbemServicesEx.ExecNotificationQuery( _
'   "SELECT * FROM __InstanceOperationEvent WITHIN 5 WHERE TargetInstance ISA 'Win32_PnPEntity'")

'Set collSWbemObjectSet = objSWbemServicesEx.ExecNotificationQuery( _
'     "SELECT * FROM __InstanceCreationEvent WITHIN 5 WHERE " & _
'    "TargetInstance ISA 'Win32_LogicalDisk'" & _
'" AND TargetInstance.DriveType = 2")

Set collSWbemObjectSet = objSWbemServicesEx.ExecNotificationQuery( _
    "SELECT * FROM __InstanceOperationEvent WITHIN 5 WHERE TargetInstance ISA 'Win32_PnPEntity' AND " & _
    "TargetInstance.Caption like '%USB Device'")

'Set collSWbemObjectSet = objSWbemServicesEx.ExecNotificationQuery( _
'    "SELECT * FROM __InstanceOperationEvent WITHIN 5 WHERE TargetInstance ISA 'Win32_PnPEntity' AND " & _
'    "TargetInstance.Description = 'Дисковый накопитель'")


WScript.Echo ""
WScript.Echo "Usb Flash detecter and data copy."
WScript.Echo "Press CTRL-C to exit."
WScript.Echo ""
WScript.Echo "Ready. Plugin an USB flash"

Set oShell = WScript.CreateObject ("WSCript.shell")

'WScript.Quit

Do
 With collSWbemObjectSet.NextEvent
  Select Case .Path_.Class
   Case "__InstanceCreationEvent"
    With .TargetInstance
     WScript.Echo "Plugged device:       ", .Name
     WScript.Echo "DeviceID:             ", .DeviceID
   
     For Each objSWbemObjectEx_Win32_DiskDrive In .Associators_("", "Win32_DiskDrive",,,,,,, wbemFlagReturnImmediately + wbemFlagForwardOnly)
'      WScript.Echo "  Physical disk drive:", objSWbemObjectEx_Win32_DiskDrive.Name
    
      If objSWbemObjectEx_Win32_DiskDrive.Partitions > 0 Then
       For Each objSWbemObjectEx_Win32_DiskPartition In objSWbemObjectEx_Win32_DiskDrive.Associators_("", "Win32_DiskPartition",,,,,,, wbemFlagReturnImmediately + wbemFlagForwardOnly)
'        WScript.Echo "    Disk partition:   ", objSWbemObjectEx_Win32_DiskPartition.Name
    
        For Each objSWbemObjectEx_Win32_LogicalDisk In objSWbemObjectEx_Win32_DiskPartition.Associators_("", "Win32_LogicalDisk",,,,,,, wbemFlagReturnImmediately + wbemFlagForwardOnly)
          DiskDrive = objSWbemObjectEx_Win32_LogicalDisk.Name
          WScript.Echo "Copy data to ", DiskDrive

'         With objSWbemObjectEx_Win32_LogicalDisk
'          WScript.Echo "Copy data to ", .Name
     
     '    WScript.Echo "      Logical disk:   ", .Name
     '    WScript.Echo "      Volume Name:   ", .VolumeName
     '    WScript.Echo "      File System:   ",  .FileSystem
'         End With
        Next
       Next
      End If
     Next

'     Set oExec = oShell.Exec("deveject.exe -EjectId:""" & .DeviceID & """" )
     Set oExec = oShell.Exec("cmd /c start_copy.bat """ & DiskDrive & """ """ & .DeviceID & """" )
     WScript.Echo ""
'     WScript.Echo oExec.Status 
     WScript.Echo oExec.StdOut.ReadAll()

     WScript.Echo
    End With
   Case "__InstanceDeletionEvent"
    With .TargetInstance
     WScript.Echo "Unplugged device:     ", .Name
     WScript.Echo "DeviceID:             ", .DeviceID
     WScript.Echo ""
     WScript.Echo "Ready. Plugin another USB flash"

    End With
   Case Else
  End Select
 End With
Loop

Set collSWbemObjectSet = Nothing
Set objSWbemServicesEx = Nothing

WScript.Quit 0



Wscript.Quit


strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set objEvents = objWMIService.ExecNotificationQuery _
("SELECT * FROM __InstanceCreationEvent WITHIN 5 WHERE " & _
"TargetInstance ISA 'Win32_LogicalDisk'" & _
" AND TargetInstance.DriveType = 2")
Wscript.Echo "Ожидаем события ..."
Do While(True)
 Set objReceivedEvent = objEvents.NextEvent
 Wscript.Echo "Name: " & objReceivedEvent.TargetInstance.Name
 Wscript.Echo "Caption: " & objReceivedEvent.TargetInstance.Caption
 Wscript.Echo "FileSystem: " & objReceivedEvent.TargetInstance.FileSystem
 Wscript.Echo "Description: " & objReceivedEvent.TargetInstance.Description
 Wscript.Echo "DeviceID: " & objReceivedEvent.TargetInstance.DeviceID
 Wscript.Echo "PNPDeviceID: " & objReceivedEvent.TargetInstance.PNPDeviceID
 Wscript.Echo "ProviderName: " & objReceivedEvent.TargetInstance.ProviderName
 Wscript.Echo "Purpose: " & objReceivedEvent.TargetInstance.Purpose
 Wscript.Echo "SystemCreationClassName: " & objReceivedEvent.TargetInstance.SystemCreationClassName
 Wscript.Echo "VolumeSerialNumber: " & objReceivedEvent.TargetInstance.VolumeSerialNumber
 Wscript.Echo "VolumeName: " & objReceivedEvent.TargetInstance.VolumeName
Loop