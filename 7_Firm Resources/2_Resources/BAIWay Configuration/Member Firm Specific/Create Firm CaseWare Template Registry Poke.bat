@Echo off
REM ======================================================================================================================================================================
REM
REM Company:	BAI Bragonier & Associates Inc
REM Version:	rev.23.07
REM Date:	2023-07-27
REM 
REM Purpose:  
REM		Create the CaseWare Templates registry poke to be deployed to all users
REM		based on the CaseWare templates configured on the current workstation
REM
REM Instructions:
REM		First configure the workstation for the CaseWare templates to be deployed using File > Templates within CaseWare Working Papers
REM		Second run this script entering, at the prompt, the CaseWare application version for which templates have been configured
REM		Once the script completes the registry poke will be created in the following location for deployment (YYYY replaced with the value entered):
REM			K:\7_Firm Resources\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare YYYY Working Papers Templates.reg
REM
REM ======================================================================================================================================================================

TITLE Exporting the Configured CaseWare Templates

REM Make sure that text is treated with the correct character code page (also useful when exporting french characters)
	CHCP 1252

REM Allow expansion of variables
	setlocal EnableDelayedExpansion

REM ============================================================================================
REM Firm CaseWare Template Custodian - Create CaseWare Template registry poke
REM ============================================================================================
		:UserChoice
		cls
		ECHO.
		ECHO ============================================================================================
		ECHO.
		ECHO	  Firm CaseWare Template Custodian
		ECHO	  Input the CaseWare version below to deploy CaseWare Templates for all staff
		ECHO.      
		ECHO	  e.g.  For CaseWare Working Papers 2024... type  2024
		ECHO	        For CaseWare Working Papers 2023... type  2023
		ECHO.
		ECHO ============================================================================================
		ECHO.
		echo.

		echo Please type the CaseWare Working Papers version number, then press the ENTER key
		echo To exit press the ENTER key
		echo.
		
		REM Get the users choice
			set /p choice=Version:
		
		REM assess the users choice
			if [%choice%] EQU [] EXIT /B
			set rsGEQ=0
			set rsLEQ=0
			if %choice% GEQ 2019 set rsGEQ=1
			if %choice% LEQ 2027 set rsLEQ=1
			if %rsGEQ% NEQ 1 goto UserChoice
			if %rsLEQ% NEQ 1 goto UserChoice

		REM Verify software has been installed, hiding all output
			reg query "HKLM\SOFTWARE\CaseWare\Working Papers\%choice%.00\Templates" >nul 2>nul

		REM	If the key does not exist, ask the user to try again
			if %ERRORLEVEL% EQU 1 (
				echo.
				echo The version of CaseWare you input does not appear to be installed
				echo.
				pause
				goto UserChoice
				)
		REM	If the key exists configure the application
			if %ERRORLEVEL% EQU 0 (call :runCWTemplate %choice% cwReturn)
			cls

		REM	If the script was successful
			if "%cwReturn%" EQU "Yes" (
				ECHO.
				ECHO ============================================================================================
				ECHO.
				ECHO Firm CaseWare %choice% Working Papers Templates registry poke successfully created
				ECHO Users will need to refresh their BAIWay configuration to receive the changes
				ECHO.
				ECHO ============================================================================================
				ECHO.
				pause
				)

		REM	If the script was not successful
			if "%cwReturn%" EQU "No" (
				ECHO.
				ECHO ============================================================================================
				ECHO.
				ECHO An error occured, please try again
				ECHO Make sure you are a member of the security group 'ClientDocs Firm Resource Managers'
				ECHO Contact your BAICoach for further assistance
				ECHO.
				ECHO ============================================================================================
				ECHO.
				pause
				)

REM ============================================================================================	
REM Script complete
Exit /B 



REM ============================================================================================
REM Function to run the script to create the Firm CaseWare Template registry poke
REM ============================================================================================
:runCWTemplate
REM Change the YYYY after the equals sign to match the deployed version of CaseWare Working Papers
	set cwVer=%1

REM Sets temporary output locations for the registry exports
	set xTemp=%temp%\temp.txt
	set xTemp32=%temp%\temp32.txt
	set xTemp6432=%temp%\temp6432.txt

REM Sets the destination path of the registry poke for deployment to all users
	set KDriveReg=K:\7_Firm Resources\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare %cwVer% Working Papers Templates.reg
	set CDriveReg=C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare %cwVer% Working Papers Templates.reg
											
REM Begin creating the registry poke
	echo Windows Registry Editor Version 5.00 > %xTemp%
	echo. >>%xTemp%
	echo ;Removing existing CaseWare templates >>%xTemp%
	echo [-HKEY_LOCAL_MACHINE\SOFTWARE\CaseWare\Working Papers\%cwVer%.00\Templates]>>%xTemp%
REM	echo [-HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CaseWare\Working Papers\%cwVer%.00\Templates]>>%xTemp%
	echo. >>%xTemp%

REM Exporting the template configurations
	reg export "HKEY_LOCAL_MACHINE\SOFTWARE\CaseWare\Working Papers\%cwVer%.00\Templates" "%xTemp32%" /y
REM	reg export "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CaseWare\Working Papers\%cwVer%.00\Templates" "%xTemp6432%" /y

REM Remove the extra title from the second registry export and Merge the two exports together
REM (Important - if you do not have CMD /U /C in front the combined text will have incorrect characters)
	CMD /A /C Type "%xTemp32%" | findstr /v "Windows Registry Editor Version 5.00" >> "%xTemp%"
REM	CMD /A /C Type "%xTemp6432%" | findstr /v "Windows Registry Editor Version 5.00" >> "%xTemp%"

REM Copy to the final destinations
	copy /y "%xTemp%" "%CDriveReg%" >nul
	copy /y "%xTemp%" "%KDriveReg%" >nul

REM If the poke cannot be copied then the user is not a ClientDocs Firm Resource Manager
	if %ERRORLEVEL% EQU 0 (set %2=Yes)
	if %ERRORLEVEL% EQU 1 (set %2=No)

REM Remove the temporary files
	del /q "%xTemp%" "%xTemp32%" "%xTemp6432%" >nul 2>nul

REM subroutine complete, return to script
	EXIT /B

REM End of BAIWay CaseWare Template Deployment Script  v23.07

