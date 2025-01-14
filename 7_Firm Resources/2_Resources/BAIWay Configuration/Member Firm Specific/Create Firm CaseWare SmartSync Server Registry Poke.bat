@Echo off
REM ======================================================================================================================================================================
REM
REM Company:	BAI Bragonier & Associates Inc
REM Version:	rev.23.07
REM Date:	2023-07-27
REM 
REM Purpose:  
REM		Create the CaseWare SmartSync Server registry poke to be deployed to all users
REM		based on the CaseWare SmartSync Server configuration on the current workstation for the current user
REM
REM Instructions:
REM		First configure the workstation for the CaseWare SmartSync Server sites to be deployed using 'File - SmartSync Server - Manage Servers' within CaseWare Working Papers
REM		Second run this script entering, at the prompt, the CaseWare application version for which you are configuring
REM		Once the script completes the registry poke will be created in the following location for deployment (YYYY replaced with the value entered):
REM			K:\7_Firm Resources\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare YYYY SmartSync Server Registry Poke.reg
REM
REM ======================================================================================================================================================================

TITLE Exporting the CaseWare SmartSync Server Configuration

REM Make sure that text is treated with the correct character code page (also useful when exporting french characters)
	CHCP 1252

REM Allow expansion of variables
	setlocal EnableDelayedExpansion

REM ============================================================================================
REM Firm CaseWare SmartSync Server Custodian - Create CaseWare SmartSync Server registry poke
REM ============================================================================================
		:UserChoice
		cls
		ECHO.
		ECHO ============================================================================================
		ECHO.
		ECHO	  Firm CaseWare SmartSync Server Custodian
		ECHO	  Input the CaseWare version below to deploy CaseWare SmartSync Server for all staff
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
			if %choice% LEQ 2025 set rsLEQ=1
			if %rsGEQ% NEQ 1 goto UserChoice
			if %rsLEQ% NEQ 1 goto UserChoice

		REM Verify software has been installed, hiding all output
			reg Query "HKCU\SOFTWARE\CaseWare International\Working Papers\%choice%.00\SyncServer" /v PublishLocation >nul 2>nul

		REM	If the key does not exist, ask the user to try again
			if %ERRORLEVEL% EQU 1 (
				echo.
				echo CaseWare SmartSync Server for the version you input does not appear to be configured
				echo.
				pause
				goto UserChoice
				)
		REM	If the key exists configure the application
			if %ERRORLEVEL% EQU 0 (call :runCWSmartSync %choice% cwReturn)
			cls

		REM	If the script was successful
			if "%cwReturn%" EQU "Yes" (
				ECHO.
				ECHO ============================================================================================
				ECHO.
				ECHO Firm CaseWare %choice% CaseWare SmartSync Server registry poke successfully created
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
REM Function to run the script to create the Firm CaseWare SmartSync Server registry poke
REM ============================================================================================
:runCWSmartSync
REM Change the YYYY after the equals sign to match the deployed version of CaseWare Working Papers
	set cwVer=%1

REM Sets temporary output locations for the registry exports
	set xTemp=%temp%\temp.txt
	set xTemp32=%temp%\temp32.txt
	set xTempFinal=%temp%\tempFinal.txt

REM Sets the destination path of the registry poke for deployment to all users
	set KDriveReg=K:\7_Firm Resources\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare %cwVer% SmartSync Server Registry Poke.reg
	set CDriveReg=C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare %cwVer% SmartSync Server Registry Poke.reg

REM Begin creating the registry poke
	echo Windows Registry Editor Version 5.00 > %xTemp%
	echo. >>%xTemp%
	echo ;Removing existing CaseWare SmartSync Server configuration >>%xTemp%
	echo [-HKCU\SOFTWARE\CaseWare International\Working Papers\%cwVer%.00\SyncServer]>>%xTemp%
	echo. >>%xTemp%

REM Exporting the CaseWare SmartSync Server configuration
	reg export "HKCU\SOFTWARE\CaseWare International\Working Papers\%cwVer%.00\SyncServer" "%xTemp32%" /y

REM Remove the extra title from the second registry export and Merge the two exports together
REM (Important - if you do not have CMD /U /C in front the combined text will have incorrect characters)
	CMD /A /C Type "%xTemp32%" | findstr /v "Windows Registry Editor Version 5.00" >> "%xTemp%"

REM Correct for a bug in the SmartSync key that exports export Groups with extra line breaks that prevents importing the groups
	set txt1=[char]13,[char]13,[char]10,[char]34
	set txt2=[char]34
	powershell -command "& {$txt1=-join(%txt1%);$txt2=%txt2%;[System.IO.File]::ReadAllText('%xTemp%').Replace($txt1,$txt2)}">"%xTempFinal%"


REM Copy to the final destinations
	copy /y "%xTempFinal%" "%CDriveReg%" >nul
	copy /y "%xTempFinal%" "%KDriveReg%" >nul

REM If the poke cannot be copied then the user is not a ClientDocs Firm Resource Manager
	if %ERRORLEVEL% EQU 0 (set %2=Yes)
	if %ERRORLEVEL% EQU 1 (set %2=No)

REM Remove the temporary files
	del /q "%xTemp%" "%xTemp32%" "%xTempFinal%" >nul 2>nul

REM subroutine complete, return to script
	EXIT /B

REM End of BAIWay CaseWare SmartSync Server Deployment Script  v23.07
