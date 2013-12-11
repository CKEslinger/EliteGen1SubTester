@echo off
ECHO Use this batch file to make a new cab file. Press CTRL-C to cancel, or
pause
IF EXIST "C:\Program Files\Microsoft Visual Studio\VB98\Wizards\PDWizard\MAKECAB.EXE" GOTO makecab_VS
IF EXIST "C:\Program Files\Microsoft Visual Studio 6.0\VB98\Wizards\PDWizard\MAKECAB.EXE" GOTO makecab_VS6
ECHO Could not find MAKECAB.EXE
GOTO end
:makecab_VS6
"C:\Program Files\Microsoft Visual Studio 6.0\VB98\Wizards\PDWizard\MAKECAB.EXE" /f "PassTimeTester.DDF"
GOTO end
:makecab_VS
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\PDWizard\MAKECAB.EXE" /f "PassTimeTester.DDF"
:end
pause
