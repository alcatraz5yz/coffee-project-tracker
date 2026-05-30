@echo off
rem Stop the PCS Dashboard cleanly by killing only the process listening on port 8090.
rem Safer than `taskkill /IM node.exe` which would kill every node process on the machine.
for /f "tokens=5" %%a in ('netstat -aon ^| findstr "LISTENING" ^| findstr ":8090"') do taskkill /F /PID %%a
