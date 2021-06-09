@echo off
Set Projectpath=C:\Users\rajivsa_90\source\repos\TestProject1\TestProject1\bin\Debug\netcoreapp3.1\
set projectdll=TestProject1.dll
set OutputPath=C:\Users\rajivsa_90\source\repos\TestProject1\TestProject1\TestReports\
@echo on
dotnet test %Projectpath%%projectdll%

timeout 5
cd %OutputPath%
timeout 5

for /f "delims=" %%d in ('dir /A:D /O:-D /b') do (
livingdoc test-assembly %Projectpath%%projectdll% -t %Projectpath%TestExecution.json --output %OutputPath%%%d\%%d.html
pause
goto :break
)
:break

pause


