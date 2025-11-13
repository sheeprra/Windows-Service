sc create AutoUPS binPath="%~dp0AutoUPS.exe" start= auto

reg add "HKLM\System\CurrentControlSet\Control\Session Manager\Shutdown" /v "WaitToKillServiceTimeout" /t REG_SZ /d "15000" /f

sc start AutoUPS

reg add "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /v HiberbootEnabled /t REG_DWORD /d 0 /f

pause