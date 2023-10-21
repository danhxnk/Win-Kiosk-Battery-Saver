@ECHO OFF
SET ACBriPercent=100
SET DCBriPercent=100

IF "%1" NEQ "" SET ACBriPercent=%1
IF "%2" NEQ "" SET DCBriPercent=%2

::Get Scheme
FOR /F "Tokens=4 Delims=: " %%a in ('powercfg /getactivescheme') Do SET Scheme=%%a

::Get Sub
FOR /F "Tokens=3 Delims=: " %%a in ('powercfg /q ^| findstr Display ^| findstr Subgroup') Do SET SubScheme=%%a

::Get Power Setting
FOR /F "Tokens=4 Delims=: " %%a in ('powercfg /q ^| findstr Display ^| findstr Power') Do SET Setting=%%a


::Config Brightness
powercfg -SetAcValueIndex %Scheme% %SubScheme% %Setting% %ACBriPercent%
powercfg -SetDcValueIndex %Scheme% %SubScheme% %Setting% %DCBriPercent%

::Set Scheme
powercfg /S %Scheme%
