@echo off
set filename=ClearCard.dll
%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe /u "%~dp0\"\%filename%
pause