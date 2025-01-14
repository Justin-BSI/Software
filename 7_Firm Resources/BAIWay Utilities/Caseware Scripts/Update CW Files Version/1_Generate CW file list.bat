@echo off
chcp 1252
REM ======================================================================================================================================================================
REM
REM Company:	BAI Bragonier & Associates Inc
REM Version:	rev.001
REM Date:		2023-04-18
REM 
REM Purpose:  
REM		Generate listing of CW files and store listing in current directory
REM
REM Instructions:
REM		This script creates the file 'cwupdatelisting.txt' in the current directory
REM		The listing is ready to used by a variety of different scripts
REM		You can change the path scanned as needed
REM
REM ======================================================================================================================================================================

REM ======================================================================================================================================================================
	TITLE	Please wait, do not close this window... a list of CW files is being created
	CLS
	ECHO	Please wait, do not close this window... a list of CW files is being created. . .
	ECHO.
REM ======================================================================================================================================================================

::  Set the path for the location to scan for CW files
	set cwPath=K:\1_Clients\Clients

::  Get the relative path of the script
	set "scriptPath=%~dp0"

::  Create the path to output the list and report
	md "%scriptPath%Output"
	
	
::  Get the listing and output to the relative path
	cd /d %cwPath%
	dir /b /s *.ac *.ac_>"%scriptPath%Output\BAIWay_Convert_List.txt"