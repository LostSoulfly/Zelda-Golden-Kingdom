@echo off
goto check_Permissions

:check_Permissions
    echo Run with administrative permissions required. Detecting permissions...
    
    net session >nul 2>&1
    if %errorLevel% == 0 (
        echo Success: Administrative permissions confirmed.
    ) else (
        echo Failure: Current permissions inadequate.
		pause
		exit
    )
	
start "" "Server\Web Server.exe"
REM Uncomment below to run multiple servers and the Hub.
REM start "" "Server\Hub.exe"
REM start "" "Server\Server.exe" -troll 0 -port 4001 -hub 1 -lock 1 -name Nayru
start "" "Server\Server.exe" -troll 0 -port 4000 -hub 0 -lock 1 -name Saria