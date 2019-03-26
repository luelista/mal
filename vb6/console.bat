set BAT_HOME=%~dp0
echo %BAT_HOME%
cd %BAT_HOME%


"C:\Programme\Microsoft Visual Studio\VB98\LINK.EXE" /EDIT /SUBSYSTEM:CONSOLE Mal.exe
Mal.exe %1
pause
