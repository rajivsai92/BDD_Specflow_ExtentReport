@echo off
Set Projectpath=%1
set projectdll=%2
set OutputPath=%3
@echo on
livingdoc test-assembly %Projectpath%%projectdll% -t %Projectpath%TestExecution.json --output %OutputPath%
pause



