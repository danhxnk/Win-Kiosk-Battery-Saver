# Win-Kiosk-Battery-Saver
 Change the screen brightness on Windows based on the battery/charge %. 

Store both files in the same directory and run KioskSleeper.vbs as admin, in your kiosk launch script. Some event log information is stored in the Application event logs in Windows.

To test that the brightness is changing, run the following :

powershell -Command "Get-Ciminstance -Namespace root/WMI -ClassName WmiMonitorBrightness | Select -ExpandProperty "CurrentBrightness""
