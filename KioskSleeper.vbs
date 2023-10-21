Set objArgs 			= Wscript.Arguments

Dim strBattery

strComputer 			= "."
strLatestBatteryPrev 	= 0
count 					= 0

'Define these vars only
minsBetweenChecks 		= 4		' Set the check frequency
checkInstances 			= 3		' How many runs at below desired battery level
batteryLow				= 25	' On battery sleep %
dcBatterythreshold		= 25	' The difference between the battery % and the brightness % on DC
acBatterythreshold		= 15	' The difference between the battery % and the brightness % on AC
OpenHour 				= 9		' When does normal operation start
CloseHour				= 17	' When does the operation stop. Beyond minute 55 the screen will dim.


If objArgs.Count 		= 0 Then
	OpenHour 			= OpenHour
	CloseHour			= CloseHour
Else
	OpenHour			= objArgs(0)
	CloseHour			= objArgs(1)
End If



'WScript.Sleep 60000

Set objShell = CreateObject("WScript.Shell") 

objShell.LogEvent 4, "Kiosk sleeper is starting"

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
objWMIService.Security_.ImpersonationLevel = 3
objWMIService.Security_.privileges.addasstring "SeDebugPrivilege", True

Function GetInfo(strLatestBattery, strBatteryStatusText)
	Set colBatteryStatus = objWMIService.ExecQuery("Select EstimatedChargeRemaining, BatteryStatus FROM Win32_Battery")
	For Each objItem in colBatteryStatus 
		strLatestBattery = objItem.EstimatedChargeRemaining
		strBattery = objItem.BatteryStatus
	Next

	Select Case strBattery
		Case "1"
			strBatteryStatusText = "Discharging"
		Case "2"
			strBatteryStatusText = "Charging"
		Case "3"
			strBatteryStatusText = "Fully Charged"
		Case "4"
			strBatteryStatusText = "Low"
		Case "5"
			strBatteryStatusText = "Critical"
		Case "6" 
			strBatteryStatusText = "Charging"
		Case "7" 
			strBatteryStatusText = "Charging and High"
		Case "8" 
			strBatteryStatusText = "Charging and Low"
		Case "9"
			strBatteryStatusText = "Charging and Critical"
		Case "10" 
			strBatteryStatusText = "Undefined"
		Case "11"
			strBatteryStatusText = "Partially Charged"
		Case Else
			strBatteryStatusText = strBattery
	End Select


End Function

Do
	GetInfo strLatestBattery, strBattery

		
	If strBattery = "Discharging" AND strLatestBattery < batteryLow Then 
		count = count +1
		objShell.LogEvent 2, "This system is discharging. The number of occurences is: " & count
	Else
		count = 0
	End If
	'WScript.Echo count & " " & checkInst

	If count >= checkInstances AND strLatestBattery < batteryLow Then
		WScript.Echo "Going to sleep"
		it = objShell.Popup("This system has no power. Going to sleep",30,"Kiosk Sleeper", 048)
		objShell.LogEvent 1, "The system is not being powered and a sleep will be initiated."
		WScript.Sleep 30000
		objShell.Run("rundll32.exe powrprof.dll,SetSuspendState")
		count = 0
	End If	
	
	If strLatestBattery -acBatterythreshold >= 0 Then ACPercent = strLatestBattery -acBatterythreshold End If
	If strLatestBattery -dcBatterythreshold >= 0 Then DCPercent = strLatestBattery -dcBatterythreshold End If
	If ACPercent <= acBatterythreshold Then ACPercent = 0 End If
	If DCPercent <= dcBatterythreshold Then DCPercent = 0 End If	
	If Hour(Now()) < OpenHour AND Minute(Now()) >= 55 OR Hour(Now()) >= CloseHour Then
		ACPercent = 0
		DCPercent = 0
		Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name Like '%.scr'")
		For Each objProcess in colProcessList
			objProcess.Terminate()
		Next
		objShell.SendKeys("{F15}") 
	End If
	If strLatestBattery <= 50 Then 
		Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name Like '%.scr'")
		For Each objProcess in colProcessList
			objProcess.Terminate()
		Next
		'Return = objShell.Run ("cmd.exe /c taskkill /IM NameOfScreensaver.scr", 0, FALSE)
		objShell.SendKeys("{F15}") 
	End If
	Return = objShell.Run ("Brightness.cmd " & ACPercent & " " & DCPercent, 0, FALSE)
	'WScript.Echo Return
	WScript.Echo Time() & " Battery Charge is " & strLatestBattery  & "% Battery Status is " & strBattery
	WScript.Echo Time() & " AC % is " & ACPercent & " DC % is " & DCPercent 

	WScript.Sleep 60000 * minsBetweenChecks
	strLatestBatteryPrev = strLatestBattery

Loop