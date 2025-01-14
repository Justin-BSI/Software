@Echo off
chcp 1252
REM ======================================================================================================================================================================
REM
REM Company:	BAI Bragonier & Associates Inc
REM Version:	rev.005
REM Date:	2023-07-27
REM 
REM Purpose:  
REM		Appy default BAIWay ClientDocs settings for CaseWare Working Papers
REM
REM Instructions:
REM		This script can be called from another script (eg the users logon script)
REM 		or may be run manually
REM
REM ======================================================================================================================================================================

	Title ClientDocs CaseWare Working Papers configuration
	
REM ======================================================================================================================================================================
	Echo Retrieving SmartSync variables
		set cwFiles=%userprofile%\CW Files
		set cwSyncServer=Local
			if exist "%appdata%\BAIWay\SmartSyncPath.txt" (set /p cwFiles=<"%appdata%\BAIWay\SmartSyncPath.txt")
			if exist "%appdata%\BAIWay\SmartSyncServer.txt" (set /p cwSyncServer=<"%appdata%\BAIWay\SmartSyncServer.txt")
	
	Echo Set variables for SmartSync Server and CaseWare Cloud
		if cwSyncServer==Local set configSSS=1
		if cwSyncServer==Hybrid set configSSS=1 && configCWC=1
		if cwSyncServer==CWCloud set configCWC=1
	
	
	Echo Configuring CaseWare Working Papers and CaseView User and Template settings

		reg import "C:\2_Resources\BAIWay Configuration\Registry Pokes\BAIWay CaseWare 2022 Working Papers + CaseView Registry Poke.reg" >nul 2>nul
		reg import "C:\2_Resources\BAIWay Configuration\Registry Pokes\BAIWay CaseWare 2022 Working Papers Templates.reg" >nul 2>nul
		reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare 2022 Working Papers Registry Poke.reg" >nul 2>nul
		reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare 2022 Working Papers Templates.reg" >nul 2>nul

		reg import "C:\2_Resources\BAIWay Configuration\Registry Pokes\BAIWay CaseWare 2023 Working Papers + CaseView Registry Poke.reg" >nul 2>nul
		reg import "C:\2_Resources\BAIWay Configuration\Registry Pokes\BAIWay CaseWare 2023 Working Papers Templates.reg" >nul 2>nul
		reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare 2023 Working Papers Registry Poke.reg" >nul 2>nul
		reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare 2023 Working Papers Templates.reg" >nul 2>nul

		reg import "C:\2_Resources\BAIWay Configuration\Registry Pokes\BAIWay CaseWare 2024 Working Papers + CaseView Registry Poke.reg" >nul 2>nul
		reg import "C:\2_Resources\BAIWay Configuration\Registry Pokes\BAIWay CaseWare 2024 Working Papers Templates.reg" >nul 2>nul
		reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare 2024 Working Papers Registry Poke.reg" >nul 2>nul
		reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare 2024 Working Papers Templates.reg" >nul 2>nul


	Echo Configuring CaseWare Working Papers SyncPath and SignOutPath

		reg add "HKCU\Software\CaseWare International\Working Papers\2022.00\Settings" /f /v SyncPath /t REG_SZ /d "%cwFiles%" >nul 2>nul
		reg add "HKCU\Software\CaseWare International\Working Papers\2022.00\Settings" /f /v SignOutPath /t REG_SZ /d "%cwFiles%" >nul 2>nul
		reg add "HKCU\Software\CaseWare International\Working Papers\2022.00\Settings" /f /v LastBackupPath /t REG_SZ /d "%cwFiles%\_Backup" >nul 2>nul
		reg add "HKCU\Software\CaseWare International\Working Papers\2022.00\Settings" /f /v LastConversionBackupPath /t REG_SZ /d "%cwFiles%\_Backup" >nul 2>nul

		reg add "HKCU\Software\CaseWare International\Working Papers\2023.00\Settings" /f /v SyncPath /t REG_SZ /d "%cwFiles%" >nul 2>nul
		reg add "HKCU\Software\CaseWare International\Working Papers\2023.00\Settings" /f /v SignOutPath /t REG_SZ /d "%cwFiles%" >nul 2>nul
		reg add "HKCU\Software\CaseWare International\Working Papers\2023.00\Settings" /f /v LastBackupPath /t REG_SZ /d "%cwFiles%\_Backup" >nul 2>nul
		reg add "HKCU\Software\CaseWare International\Working Papers\2023.00\Settings" /f /v LastConversionBackupPath /t REG_SZ /d "%cwFiles%\_Backup" >nul 2>nul

		reg add "HKCU\Software\CaseWare International\Working Papers\2024.00\Settings" /f /v SyncPath /t REG_SZ /d "%cwFiles%" >nul 2>nul
		reg add "HKCU\Software\CaseWare International\Working Papers\2024.00\Settings" /f /v SignOutPath /t REG_SZ /d "%cwFiles%" >nul 2>nul
		reg add "HKCU\Software\CaseWare International\Working Papers\2024.00\Settings" /f /v LastBackupPath /t REG_SZ /d "%cwFiles%\_Backup" >nul 2>nul
		reg add "HKCU\Software\CaseWare International\Working Papers\2024.00\Settings" /f /v LastConversionBackupPath /t REG_SZ /d "%cwFiles%\_Backup" >nul 2>nul

	Echo Configuring CaseView Dictionary
		robocopy /W:5 /R:3 "C:\2_Resources\BAIWay Configuration\Files to Push\CaseView" "C:\ProgramData\CaseWare\Working Papers\2022.00\CaseView" cvwin.ini >nul 2>nul
		robocopy /W:5 /R:3 "C:\2_Resources\BAIWay Configuration\Files to Push\CaseView" "C:\ProgramData\CaseWare\Working Papers\2023.00\CaseView" cvwin.ini >nul 2>nul
		robocopy /W:5 /R:3 "C:\2_Resources\BAIWay Configuration\Files to Push\CaseView" "C:\ProgramData\CaseWare\Working Papers\2024.00\CaseView" cvwin.ini >nul 2>nul

	Echo Configuring QAT and Ribbon Settings
		REM Check to see if the script was called with a variable to skip configuring toolbar [0], if not default to configCWQAT=1
			set configCWQAT=1
			if [%1] EQU [0] set configCWQAT=0

		REM Configure QAT if configCWQAT equals 1
			if %configCWQAT%==1 (
				robocopy /W:5 /R:3 "C:\2_Resources\BAIWay Configuration\Files to Push\CaseView" "%LOCALAPPDATA%\CaseWare\Working Papers\2022.00\CaseView" CvRibbon.xml >nul 2>nul	
				robocopy /W:5 /R:3 "C:\2_Resources\BAIWay Configuration\Files to Push\CaseView" "%LOCALAPPDATA%\CaseWare\Working Papers\2023.00\CaseView" CvRibbon.xml >nul 2>nul
				robocopy /W:5 /R:3 "C:\2_Resources\BAIWay Configuration\Files to Push\CaseView" "%LOCALAPPDATA%\CaseWare\Working Papers\2024.00\CaseView" CvRibbon.xml >nul 2>nul

				reg import "C:\2_Resources\BAIWay Configuration\Registry Pokes\BAIWay CaseWare 2022 Working Papers + CaseView QAT Registry Poke.reg" >nul 2>nul
				reg import "C:\2_Resources\BAIWay Configuration\Registry Pokes\BAIWay CaseWare 2023 Working Papers + CaseView QAT Registry Poke.reg" >nul 2>nul
				reg import "C:\2_Resources\BAIWay Configuration\Registry Pokes\BAIWay CaseWare 2024 Working Papers + CaseView QAT Registry Poke.reg" >nul 2>nul
				)

	Echo Configuring CaseWare SmartSync Server
		if configSSS==1 ( 

			reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare 2022 SmartSync Server Registry Poke.reg" >nul 2>nul	
			reg add "HKCU\Software\CaseWare International\Working Papers\2022.00\Settings" /f /v DefaultPath /t REG_SZ /d "%cwFiles%" >nul 2>nul
			reg add "HKCU\Software\CaseWare International\Working Papers\2022.00\Settings" /f /v YECPath /t REG_SZ /d "%cwFiles%" >nul 2>nul

			reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare 2023 SmartSync Server Registry Poke.reg" >nul 2>nul
			reg add "HKCU\Software\CaseWare International\Working Papers\2023.00\Settings" /f /v DefaultPath /t REG_SZ /d "%cwFiles%" >nul 2>nul
			reg add "HKCU\Software\CaseWare International\Working Papers\2023.00\Settings" /f /v YECPath /t REG_SZ /d "%cwFiles%" >nul 2>nul


			reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare 2024 SmartSync Server Registry Poke.reg" >nul 2>nul
			reg add "HKCU\Software\CaseWare International\Working Papers\2024.00\Settings" /f /v DefaultPath /t REG_SZ /d "%cwFiles%" >nul 2>nul
			reg add "HKCU\Software\CaseWare International\Working Papers\2024.00\Settings" /f /v YECPath /t REG_SZ /d "%cwFiles%" >nul 2>nul
			)
		
	Echo Configuring CaseWare Cloud
		if configCWC==1 ( 
			reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare 2022 CaseWare Cloud Registry Poke" >nul 2>nul
			reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare 2023 CaseWare Cloud Registry Poke" >nul 2>nul
			reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare 2024 CaseWare Cloud Registry Poke" >nul 2>nul	
			)

	Echo Configuring CaseWare Data Store
		reg import "C:\2_Resources\BAIWay Configuration\Member Firm Specific\Firm CaseWare Data Store Registry Poke.reg" >nul 2>nul	