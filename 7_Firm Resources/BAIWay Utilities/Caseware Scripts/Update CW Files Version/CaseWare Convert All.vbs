Option explicit

Dim scriptdir
	scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

'Include the tools
	includeFile "BAI_Toolbox.txt"
'=========================================================================================================================================================================================
'----------------------------------------------------------------------     Variables to manually change go here     ------------------------------------------------------------------

'Set the location of the logging file
	Dim logFile:	logFile = "CWConvertLogFile.txt"

	logFile = scriptdir & "\" & logFile
	
	
'Set the location of the file listing the CW files to update	
	Dim inFile:		inFile  = scriptdir & "\" & "Output\BAIWay_Convert_List.txt"

'Uncompress files before converting (True or False)
	dim uCompress
		uCompress = True
		
'Recompress at the end (True or False)
	Dim rCompress
		rCompress = True
		
'Compress All at the end (True or False)
	Dim aCompress
		aCompress = False
		

'---------------------------------------------------------------------- End of Variables to manually change go here     ------------------------------------------------------------------
'=========================================================================================================================================================================================
'*****************************************************************************************************************************************************************************************
'
'Author:			  					  Gerry Wilton, CPA, CA
'Company:					     BAI Bragonier & Associates Inc
'Version:    				 					        rev.003
'Date: 					         					 2023-11-23
'
'Purpose:      Reads listing of CW Files and converts them to current version. 
'                By default will not uncompress any listed files that are compressed before converting, and will not convert them
'                     You can change this by changing the last line in the previous section to uCompress=True 
'                By default will not recompress after converting
'                     You can change this by changing the last line in the previous section to rCompress=True 
'                By default will not compress all files after converting
'                     You can change this by changing the last line in the previous section to aCompress=True 

'
'*****************************************************************************************************************************************************************************************
'*****************************************************************************************************************************************************************************************

'Open the selected file
		Dim msgNote			'String for the progress update

	'Declare and set the items needed to read the listing of CaseWare files to update
		Dim cList			'array listing files to process
			cList = getFileList(inFile)
			
	'Establish the number of files to process
		Dim fileCount	
			fileCount = UBound(cList)
	'Generic counter
		dim k
		
	'Type tracker
		Dim sourceType

' Heading creation tracker
	Dim mHeading
		mHeading = False
		
	'Loop through the listing
			For k = 0 To fileCount
				Do ' Keep processing the item unless we hit a roadblock

					logMe("Start")

					'show the progress once every 20
						if (k mod 20) = 0 then 
							msgNote = "Processing " &(k+1) & " of " & (fileCount + 1) & " items listed" & vbCrLf & cList(k)
							popMSG msgNote,1
						End If

					'Check for that the file exists
						If cwExists(cList(k)) Then
							logMe("Success - File exists")
						Else
							logMe("Failure - File not exist")
							Exit Do
						End If

					'Verify source file type
						If cwTypeComp(cList(k)) Then
							sourceType=False
							logMe("Source State - Compressed")
							'Exit if we are not uncompressing
								If Not(uCompress) Then
										logMe("Ignored - Compressed")
										Exit Do
								End If
						ElseIf cwType(cList(k)) Then
							sourceType=True
							logMe("Source State - Uncompressed")
						Else
							logMe("Failure - Invalid file type")
							Exit Do
						End If
						
					'Check that the file version needs to be updated
						If cwCheck(cList(k)) Then
							 logMe("Already Converted")
							 Exit Do
						End If

					'Uncompress source if required
						On Error Resume Next
						If Not(sourceType) Then
									cwUncompress(cList(k))
									cList(k)=Left(cList(k),Len(cList(k))-1)
						End If


						'Check to see if an error occured during uncompress
							If Err.Number <> 0 Then
									'Disable error handling
										On Error goto 0 
										logMe("Error - Uncompress failed")
										Exit Do
								Else
										logMe("Source - Uncompressed")
							End If

					'Attempt the update
						 cwConvert(cList(k))
						'Check to see if an error occured during conversion
							If Err.Number <> 0 Then
									'Disable error handling
										On Error goto 0 
										logMe("Error - Conversion failed")
										Exit Do
								Else
									logMe("Converted")
							End If
					
					'Attempt to compress if required
						 If (rCompress And Not(sourceType)) Or aCompress Then
						 	cwCompress(cList(k))
							'Check to see if an error occured during compress
							If Err.Number <> 0 Then
									'Disable error handling
										On Error goto 0 
										logMe("Error - Compress failed") 
										Exit Do
								Else
									logMe("Compressed")
							End If
						 End If
						

				Loop While False ' end of null loop
				logMe("End")
			Next



MsgBox "Done!"

'=========================================================================================================================================================================================
'----------------------------------------------------------------------     Tool for loading external vbs files     ------------------------------------------------------------------

'Standard file system object for getting files


' Includes a file in the global namespace of the current script.
' The file can contain any VBScript source code.
' The path of the file name must be specified relative to the
' directory of the main script file.
	Private Sub IncludeFile (ByVal RelativeFileName)
	   Dim fso
		set fso = CreateObject("Scripting.FileSystemObject")
	   Dim ScriptDir
		ScriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
	   Dim FileName
		FileName = fso.BuildPath(ScriptDir,RelativeFileName)
	   IncludeFileAbs FileName
	   End Sub

' Includes a file in the global namespace of the current script.
' The file can contain any VBScript source code.
' The path of the file name must be specified absolute (or
' relative to the current directory).

	Private Sub IncludeFileAbs (ByVal FileName)
		Dim fso
			set fso = CreateObject("Scripting.FileSystemObject")
		Const ForReading = 1
	    Dim f
			set f = fso.OpenTextFile(FileName,ForReading)
	    Dim s
			s = f.ReadAll()
	    ExecuteGlobal s
	   End Sub

'----------------------------------------------------------------------  End of Tool for loading external vbs files     ------------------------------------------------------------------
'=========================================================================================================================================================================================


'--------------------------------------------------------------------------------------------
'Function to check if the file type appears to be '.ac'
'Returns True or False
'--------------------------------------------------------------------------------------------
	Function logMe(logMsg)
		If k = 0  And Not(mHeading) Then
			logger(logfile).write "Progress " & vbTab & "Message" & vbTab & "Filename" & vbCrLf
			mHeading = True
		End If

		logger(logfile).write "Processing " &(k+1) & " of " & (fileCount + 1) & vbTab & logMsg & vbTab & cList(k) & vbCrLf
	

	
	End Function