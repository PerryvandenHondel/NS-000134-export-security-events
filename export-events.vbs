''
''	SCRIPT:
''		export-events.vbs
''
''
''	SCRIPT_ID:
''		000134
''
'' 
''	DESCRIPTION:
''		This scrip performs the following actions
''		1) Export Security Event Logs to LPR using LogParser
''		2) Convert LPR to SKV
''		3) Transfer the LPR and the SKV files to the COLLECTOR DC
''		4) The collector DC transfers the PSV and SKV files to the Splunk server.
''
''
''	VERSION:
''		02	2015-04-14	Modification:
''						1) Do not export computer accounts, use LPR2SKV.EXE options --skip-computer-account
''		01	2015-04-08	First version
'' 
''	SUBS AND FUNCTIONs:
''		Function ConvertUsingLpr2skv
''		Function DoesShareExits
''		Function GetRandomCharString
''		Function ExportEventsUsingLogparser
''		Function GetComputerName
''		Function GetFileSize
''		Function GetPathLastRun
''		Function GetProgramPath
''		Function GetScriptPath
''		Function GetUniqueFileName
''		Function IsThisScriptAlreadyRunning
''		Function LastRunGet
''		Function LastRunPut
''		Function NumberAlign
''		Function ProperDateFs
''		Function ProperDateTime
''		Function RunCommand
''		Sub MakeFolder
''		Sub MoveCollectorToSplunkServer
''		Sub MoveExportFolderToCollectorDc
''		Sub ProcessEventLog
''		Sub ScriptDone
''		Sub ScriptInit
''		Sub ScriptRun
'' 



Option Explicit



''	-----------------------------------------------------------------------------------------------
''	GLOBAL CONSTANTS
''	-----------------------------------------------------------------------------------------------


Const 	FOR_READING =				1
Const 	FOR_WRITING =				2
Const	FOR_APPENDING =				8
Const	EXTENSION_LPR = 			".lpr"
Const	EXTENSION_SKV = 			".skv"
Const	LOGPARSER_OK = 				0
Const	LOGPARSER_FAIL = 			1
Const	LOGPARSER_SEPARATOR	=		"|"
Const	SHARE_SKV = 				"\\vm70as006.rec.nsint\000134-SKV"
Const	SHARE_LPR = 				"\\vm70as006.rec.nsint\000134-LPR"
Const	COLLECTOR_DC = 				"NS00DC011"
Const	COLLECTOR_SHARE = 			"000134-COLLECTOR"

''	-----------------------------------------------------------------------------------------------
''	GLOBAL VARIABLES
''	-----------------------------------------------------------------------------------------------




Dim		gobjFso
Dim		gstrComputerName
Dim		gdtmInitLastRun




''	-----------------------------------------------------------------------------------------------
''	FUNCTIONS
''	-----------------------------------------------------------------------------------------------

Function GetRandomCharString(ByVal intLen)
	''
	''	Returns a string of random chars of intLen length.
	'' 	                           12345678901234567890
	''	GetRandomCharString(20) >> 12ghyUjHsdbeH5fDsYt6
	''
	
	Dim		strValidChars	'' String with valid chars
	Dim		i				'' Random position of strValidChars
	Dim		r				'' Function return value
	Dim		x				'' Loop counter
	
	strValidChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyz"
	
	For x = 1 to intLen
		Randomize
		'' intNumber = Int(1 - Len(strValidChars) * Rnd + 1)
		i = Int(Len(strValidChars) * Rnd + 1)
		
		r = r & Mid(strValidChars, i, 1)
	Next
	GetRandomCharString = r
End Function '' of Function GetRandomCharString


Function GetUniqueFileName(ByVal dt)
	''
	''	Generate a file name of hex numbers of length 32 chars.
	''
	''	First 12 digits are the current year, month, day, hour, minutes and seconds 
	''	Extra unique chars to fill-up to 32 chars length.
	'' 
	'' 	Function does not add a extension to the file name.
	''
	''				12345678901234567890123456789012
	''	strPrefix =	YYYYMMDDHHMMSS-xxxxxxxxxxxxxxxxx
	''
	''	dt:	Date Time of the last batch is done
	''
	
	Const	UNIQUE_LEN =		32
	Const	HEX_LEN	=			8
	Const	NUM_LOW	=			0
	Const	NUM_HIGH =			15
	
	Dim		r
	Dim		strFilename
	Dim		dtmNow
	Dim		strPrefix
	Dim		i
	Dim		intNumber
	Dim		n
	
	'' Get the current date time.
	''dtmNow = Now()
	
	'' Make the prefix as YYYYMMDDHHMMSS.
	strPrefix = Year(dt) & NumberAlign(Month(dt), 2) & NumberAlign(Day(dt), 2) & "-"
	strPrefix = strPrefix & NumberAlign(Hour(dt), 2) & NumberAlign(Minute(dt), 2) & NumberAlign(Second(dt), 2)
	strPrefix = strPrefix & "-"

	'' Generate the fill string up to UNIQUE_LEN chars with a hex number.
	'For i = 1 to UNIQUE_LEN - Len(strPrefix)
	'	Randomize
	'	intNumber = Int((NUM_HIGH - NUM_LOW + 1) * Rnd + NUM_LOW)
	'	strFilename = strFilename & LCase(Hex(intNumber))
	'Next
	
	GetUniqueFileName = strPrefix & GetRandomCharString(UNIQUE_LEN - Len(strPrefix))
End Function ' GetUniqueFileName


Function ProperDateTime(dDateTime)
	'
	'	Convert a system formatted date time to a proper format
	'	Returns the current date time in proper format when no date time
	'	is specified.
	'
	'	15-5-2009 4:51:57  ==>  2009-05-15 04:51:57
	'

	If Len(dDateTime) = 0 Then
		dDateTime = Now()
	End If
	
	ProperDateTime = NumberAlign(Year(dDateTime), 4) & "-" & _
		NumberAlign(Month(dDateTime), 2) & "-" & _
		NumberAlign(Day(dDateTime), 2) & " " & _
		NumberAlign(Hour(dDateTime), 2) & ":" & _
		NumberAlign(Minute(dDateTime), 2) & ":" & _
		NumberAlign(Second(dDateTime), 2)
End Function



Function NumberAlign(ByVal intNumber, ByVal intLen)
	'
	'	Returns a number aligned with zeros to a defined length
	'
	' 	NumberAlign(1234, 6) returns '001234'
	' 
	NumberAlign = Right(String(intLen, "0") & intNumber, intLen)
End Function '' NumberAlign()



Function GetComputerName()
	''
	''     Returns the name of the computer that runs the script.
	''
	''    Input:
	''		None
	''
	''	Output:
	''		A string with the computer name that runs the script
	''
	Dim     oNetwork
	
	Set oNetwork = CreateObject("WScript.Network")
	GetComputerName = oNetwork.ComputerName
	Set oNetwork = Nothing
End Function '' GetComputerName()



Function GetScriptPath()
	''
	''	Returns the path where the script is located.
	''
	''	Output:
	''		A string with the path where the script is run from.
	''
	''		drive:\folder\folder (no last backslash)
	''
	Dim sScriptPath
	Dim sScriptName

	sScriptPath = WScript.ScriptFullName
	sScriptName = WScript.ScriptName
	GetScriptPath = Left(sScriptPath, Len(sScriptPath) - Len(sScriptName) - 1)
End Function '' GetScriptPath()



Function GetProgramPath(sProgName)
	'==
	'==	Locates a command line program in the path of the user,
	'==	or in the current folder where the script is started.
	'==
	'==	Returns:
	'==		Path to program when found
	'==		Blank string when program is not found
	'==
	Dim	oShell
	Dim	sEnvPath
	Dim	oColVar
	Dim	aPath
	Dim	sScriptPath
	Dim	sScriptName
	Dim	x
	Dim	oFso
	Dim	sPath
	Dim	sReturn

	sReturn = ""

	Set oFso = CreateObject("Scripting.FileSystemObject")
	Set oShell = CreateObject("WScript.Shell")
	
	sScriptPath = WScript.ScriptFullName
	sScriptName = WScript.ScriptName

	sScriptPath = Left(sScriptPath, Len(sScriptPath) - Len(sScriptName))
	
	'=
	'=	Build the path string like:
	'=		folder;folder;folder;...
	'=
	'=	Place the current folder first in line. So it will find the file first when
	'=	it is in the same folder as the script.
	'=
	sEnvPath = sScriptPath & ";" & oShell.ExpandEnvironmentStrings("%PATH%")
	
	'WScript.Echo sEnvPath
	aPath = Split(sEnvPath, ";")
	For x = 0 To UBound(aPath)
		If Right(aPath(x), 1) <> "\" Then
			aPath(x) = aPath(x) & "\"
		End If
		
		'WScript.Echo x & ": " & aPath(x)
		sPath = aPath(x) & sProgName
		'WScript.Echo sPath
		If oFso.FileExists(sPath) = True Then
			sReturn = sPath
			Exit For
		End If
	Next
	
	Set oShell = Nothing
	Set oFso = Nothing
	'= Return the string with double quotes enclosed. For paths with spaces.
	'GetProgramPath = Chr(34) & sReturn & Chr(34)
	'= 2011-02-16 Removed the Chr(34); was not working.
	GetProgramPath = sReturn
End Function '' of Function GetProgramPath



Function GetPathLastRun(ByVal strComputerName, ByVal strEventLog)
	''
	''	Returns the path for the lastrun file (scriptpath\release\computer\last-run.txt)
	''	
	Dim		r
	
	r = GetScriptPath() & "\lastrun\" & UCase(strComputerName) & "-" & strEventLog & ".lrdt"
	Call MakeFolder(r)
	GetPathLastRun = r
End Function



Function LastRunGet(ByVal strComputerName, ByVal strEventLog)
	'
	'	Read the last run date time from a text file.
	'	When the file does not exists. Create a file.
	'	Place a date time of previous day in it.
	'
	Dim		objFile
	Dim		r
	Dim		strPath
	
	strPath = GetPathLastRun(strComputerName, strEventLog)
	
	If gobjFso.FileExists(strPath) = True Then
		Set objFile = gobjFso.OpenTextFile(strPath, FOR_READING)
		r = ProperDateTime(objFile.ReadLine)
		objFile.Close
		Set objFile = Nothing
	Else
		'' Can't file a file so this must the first time we use this function.
		'' Create the file and return strLine value which contains a date time 5 mins ago.
		
		Set objFile = gobjFso.OpenTextFile(strPath, FOR_WRITING, True)
		r = ProperDateTime(DateAdd("m", -5, Now()))
		objFile.WriteLine r
		objFile.Close
		Set objFile = Nothing
	End If
	LastRunGet = r
End Function

'	-----------------------------------------------------------------------------------------------

Function LastRunPut(ByVal strComputerName, ByVal strEventLog)
	'
	'	Put the current date time in a the last run file.
	'	Returns the current date time
	'
	Dim		objFile
	Dim		strPath
	Dim		r
	
	strPath = GetPathLastRun(strComputerName, strEventLog)
	
	If gobjFso.FileExists(strPath) = True Then
		'' Write the current date time to the file.
		Set objFile = gobjFso.OpenTextFile(strPath, FOR_WRITING, True)
		r = ProperDateTime("") '' Get the current date time
		objFile.WriteLine r
		objFile.Close
		Set objFile = Nothing
	End If
	LastRunPut = r
End Function



Function RunCommand(sCommandLine)
	''
	''	RunCommand(sCommandLine)
	''
	''	Run a DOS command and wait until execution is finished before the script can commence further.
	''
	''	Input
	''		sCommandLine	Contains the complete command line to execute 
	''
	Dim oShell
	Dim sCommand
	Dim	nReturn

	Set oShell = WScript.CreateObject("WScript.Shell")
	sCommand = "CMD /c " & sCommandLine
	' 0 = Console hidden, 1 = Console visible, 6 = In tool bar only
	'LogWrite "RunCommand(): " & sCommandLine
	nReturn = oShell.Run(sCommand, 6, True)
	Set oShell = Nothing
	RunCommand = nReturn 
End Function '' RunCommand



Function GetFileSize(ByVal sFName)
	'
	'	GetFileSize
	'
	'	Return the length of a file or folder. Returns -1 when file or folder does not exist
	'

	Dim		objFso
	Dim		objFile
	
	GetFileSize = -1
	
	Set objFso = CreateObject("Scripting.FileSystemObject")
	
	If objFso.FileExists(sFName) = True Then
		Set objFile = objFso.GetFile(sFName)
		GetFileSize = objFile.Size
	End If
	
	Set objFile = Nothing
	Set objFso = Nothing
End Function


Function ExportEventsUsingLogparser(ByVal strComputer, ByVal dtmLastRun, ByVal dtmNow, ByVal strPathLogparser)
	''
	''	Export the events to a temp file specified in strPathLogparser
	''
	''	Return
	''		0		Noting was exported
	''		1		Export was successful and export file contains data. (file size > 0)
	''
	Dim		r
	Dim		c
	Dim		strPathExe
	Dim		intReturn
	Dim		intFileSize
	''Call LogWrite("--")
	''Call LogWrite("ExportEventsUsingLogparser()")
	
	''Call LogWrite("  strComputer:      " & strComputer)
	''Call LogWrite("  strPathLogparser: " & strPathLogparser)
	
	r = 0
	intReturn = 0

	'' Make the folder for the export strPathLogparser
	MakeFolder(strPathLogparser)
	
	'strPathExe = GetProgramPath("LogParser.exe")
	
	'TimeGenerated,EventId,REPLACE_STR(Strings,'\u000d\u000a','|') AS Strings FROM \\NS00DC009\Security" -stats:OFF -oSeparator:"~" -formatMsg:OFF
	
	c = "logparser.exe "
	c = c & "-i:EVT "
	c = c & "-o:TSV "
	c = c & Chr(34)
	c = c & "SELECT "
	c = c & "TimeGenerated,EventId,EventType,REPLACE_STR(Strings,'\u000d\u000a','|') AS Strings "
	c = c & "FROM "
	c = c & "\\" & strComputer & "\Security "
	c = c & "WHERE TimeGenerated>'" & dtmLastRun & "' AND TimeGenerated<='" & dtmNow & "'"
	''c = c & "AND EventId IN (" & strEvents & ")"
	c = c & Chr(34)
	c = c & " -stats:OFF"
	c = c & " -oSeparator:" & Chr(34) & LOGPARSER_SEPARATOR & Chr(34)
	c = c & " >" & Chr(34) & strPathLogparser & Chr(34)
	
	WScript.Echo
	WScript.Echo c
	WScript.Echo
	WScript.Echo "Running logparser.exe to export to " & strPathLogparser
	
	r = RunCommand(c)
	If r = 0 Then
		''Call LogWrite("  SUCCESS: Logparser exported success full the events: r=" & r)
		intFileSize = GetFileSize(strPathLogparser)
		If intFileSize > 0 Then
			WScript.Echo "  SUCCESS: File " & strPathLogparser & " contains data, file size is " & intFileSize & " bytes"
			intReturn = LOGPARSER_OK
		Else
			WScript.Echo "  WARNING: File " & strPathLogparser & " contains no data, deleting the this file"
			''Call DeleteFile(strPathLogparser)
			intReturn = LOGPARSER_FAIL
		End If
	Else
		WScript.Echo "  ERROR: Logparser was unable to export the events, r=" & r
		intReturn = LOGPARSER_FAIL
	End If
	ExportEventsUsingLogparser = intReturn
End Function



Function ConvertUsingLpr2skv(ByVal strPathLpr, ByVal strPathSkv)
	Dim		c
	Dim		r
	Dim		el
	Dim		intFileSize
	
	c = "lpr2skv.exe "
	c = c & Chr(34) & strPathLpr & Chr(34) & " "
	c = c & "--skip-computer-account" '' V02: Added --skip-computer-account
	
	WScript.Echo
	WScript.Echo c
	WScript.Echo
	
	
	el = RunCommand(c)
	If el = 0 Then
		intFileSize = GetFileSize(strPathSkv)
		If intFileSize > 0 Then
			WScript.Echo "Converted file "& strPathSkv & " contains " & intFileSize & " bytes"
			r = 0
		Else
			WScript.Echo " ERROR: File conversion " & strPathLpr & " failed with code: " & el
			r = 1
		End If
	Else
		r = 1
	End If
	ConvertUsingLpr2skv = r
End Function



Function DoesShareExist(ByVal strShareName)
	''
	''	Check if a share exists on the computer
	''
	''	Source: http://stackoverflow.com/questions/7980214/check-if-share-exists-if-so-then-continue
	''
	Dim		strComputer
	Dim		objWMIService
	Dim		colShares
	Dim		objShare
	Dim		r

	r = False
	
	strComputer = "."
	
	Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set colShares = objWMIService.ExecQuery("Select * from Win32_Share Where Name = '" & strShareName & "'")

	For each objShare in colShares
		
		If (Err.Number <> 0) Then 
			r = False '' strShareName share does not exists.
		Else 
			r = True '' strShareName share exists.
		End If 
	Next
	DoesShareExist = r
End Function '' of Function DoesShareExist



Function ProperDateFs(ByVal dtmDateTime, ByVal blnFolder3)
	''
	''	Convert a system formatted date time to a proper file system date time
	''
	''	Returns the current date time when no date time is specified by dtmDateTime
	''
	''	Returns a date time in format: YYYY-MM-DD
	''
	''	dtmDateTime 
	''
	''	blnFolder3
	''		True: 	Uses '\' as the separator char in the date: YYYY\MM\DD
	''		False:	Uses '-' as the separator char in the date: YYYY-MM-DD
	''
	Dim		strSeperator
	Dim		strResult


	strResult = ""
	
	If blnFolder3 = True Then
		strSeperator = "\"
	Else
		strSeperator = "-"
	End If
	
	If Len(dtmDateTime) = 0 Then
		dtmDateTime = Now()
	End If
	
	strResult = NumberAlign(Year(dtmDateTime), 4) & strSeperator 
	strResult = strResult & NumberAlign(Month(dtmDateTime), 2) & strSeperator
	strResult = strResult & NumberAlign(Day(dtmDateTime), 2)
	
	ProperDateFs = strResult
End Function '' of Function ProperDateFS




Function IsThisScriptAlreadyRunning()
	'
	'	Check in the process list if there is another instance of this script running.
	'
	'	Returns:
	'		True		Another instance of this script is already running
	'		False		No instance of this script is running on this computer
	'
	Dim		strComputer
	Dim		objWMIService
	Dim		colItems
	Dim		objItem
	Dim		intCount
	Dim		r
	
	r = False
	strComputer = "."
	
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'cscript.exe' OR Name = 'wscript.exe'")
	
	intCount = 0
	For Each objItem in colItems
		' WScript.Echo objItem.CommandLine & " == " & WScript.ScriptName 
		
		If InStr(objItem.CommandLine, WScript.ScriptName) > 0 Then
			intCount = intCount + 1
		End If
		If intCount > 1 Then
			'WScript.Echo "Another instance of " & WScript.ScriptName & " is already running"
			r = True
		End If
	Next
	IsThisScriptAlreadyRunning = r
End Function ' IsThisScriptAlreadyRunning()



''	-----------------------------------------------------------------------------------------------
''	SUBS
''	-----------------------------------------------------------------------------------------------



Sub MakeFolder(ByVal sNewFolder)
	'
	'	Create a folder structure when it doesn't exist.
	'
	'	Parameters:
	'		sNewFolder	Contains the path of the folder structure
	'					e.g. C:\This\Is\A\New\Folder or
	'					\\server\share\folder\folder
	'
	'	Added
	'		When the path contains a file name (d:\folder\file.ext)
	'		It will be deleted first from the sNewFolder.
	'	
	'	Returns:
	'		True		Folder created.
	'		False		Folder could not be created.
	'

	Dim		objFso
	Dim		arrFolder
	Dim		c
	Dim		intCount
	Dim		intRootLen
	Dim		strCreateThis
	Dim		strPathToCreate
	Dim		strRoot
	Dim		x
	Dim		bReturn

	bReturn = False

	Set objFSO = CreateObject("Scripting.FileSystemObject")
		
	'	If the sNewFolder contains a file name (d:\folder\file.ext)
	'	Return only the path and delete file.ext from the sNewFolder.
	If InStrRev(sNewFolder, ".") > 0 Then
		sNewFolder = Left(sNewFolder, InStrRev(sNewFolder, "\") - 1)
	End If
	
	If objFSO.FolderExists(sNewFolder) = False Then
		'	WScript.Echo "Folder " & sNewFolder & " does not exists, creating it."
		If Right(sNewFolder, 1) = "\" Then
			sNewFolder = Left(sNewFolder, Len(sNewFolder) - 1)
		End If
		
		If Mid(sNewFolder, 2, 1) = ":" Then
			'	Path contains a drive letter (e.g. 'D:')
			intRootLen = 2 
			strPathToCreate = Right(sNewFolder, Len(sNewFolder) - intRootLen)
			strRoot = Left(sNewFolder, intRootLen)
		Else
			'	Path contains a share name (e.g. '\\server\share')
			intCount = 0
			intRootLen = 0
			For intRootLen = 1 To Len(sNewFolder)
				c = Mid(sNewFolder, intRootLen, 1)
				If c = "\" Then
					intCount = intCount + 1
				End If
				If intCount = 4 Then
					Exit For
				End If
			Next
			intRootLen = intRootLen - 1
			strPathToCreate = Right(sNewFolder, Len(sNewFolder) - intRootLen)
			strRoot = Left(sNewFolder, intRootLen)
		End If
		arrFolder = Split(strPathToCreate, "\")
		strCreateThis = strRoot
	
		For x = 1 To UBound(arrFolder)
			strCreateThis = strCreateThis & "\" & arrFolder(x)

			'	s = s & "\" & arrFolder(x)
			If Not objFSO.FolderExists(strCreateThis) Then
				On Error Resume Next
				objFSO.CreateFolder strCreateThis
				If Err.Number <> 0 Then
					WScript.Echo "MakeFolder: Error: Can't create " & strCreateThis
				End If
			End If
		Next
	End If
	Set objFSO = Nothing
End Sub '' of Sub MakeFolder


Sub ExtractFilePerEvent(ByVal strFolderExportTo, ByVal strPathLpr)
	''
	''	TEST 
	'' 
	''	Extract a file per event
	''
	Dim		objFile
	Dim		objFileEvent
	Dim		strLine
	Dim		arrLine
	Dim		strEvent
	Dim		strPathEvent
	Dim		x
	
	Set objFile = gobjFso.OpenTextFile(strPathLpr, FOR_READING)
	
	strPathEvent = strFolderExportTo
	
	'strLine = ProperDateTime(objFile.ReadLine)

	Do While objFile.AtEndOfStream = False
		strLine = objFile.ReadLine
		'WScript.Echo strLine
		arrLine = Split(strLine, "|")
		
		strEvent = arrLine(1)
		
		
		
		
		strPathEvent = strFolderExportTo & "\event-" & strEvent & ".lpr"
		If gobjFso.FileExists(strPathEvent) = False Then
			'' Write the event line to a file
			'WScript.Echo "Writing to " & strPathEvent & ": " & strLine
		
			Set objFileEvent = gobjFso.OpenTextFile(strPathEvent, FOR_WRITING, True)
			
			objFileEvent.WriteLine strLine
			
			objFileEvent.WriteLine
			
			For x = 0 To UBound(arrLine)
				'WScript.Echo x & vbTab & arrLine(x)
				objFileEvent.WriteLine x & vbTab & arrLine(x)
			Next
			
			objFileEvent.Close
			Set objFileEvent = Nothing
		End If
		
	Loop
	objFile.Close
	Set objFile = Nothing	
End Sub '' of Sub ExtractFilePerEvent



Sub MoveExportFolderToCollectorDc(ByVal strFolderSource)
	''
	''	Move all files in the local export folder to the collector share.
	''

	Dim		c
	Dim		strFolderDest
	Dim		el				' DOS Error Level
	
	
	strFolderDest = "\\" & COLLECTOR_DC	& "\" & COLLECTOR_SHARE
	
	'' Robocopy.exe options:
	''		/z			Restartable
	''		/mov		move files
	''		/s			Copy subdirectories, but not empty ones.
	''		/r:9 /w:10	Retry and Wait if failure
	''		/np			Do no show process, skrews up your log
	''		/copy:dt	copyflags : D=Data, A=Attributes, T=Timestamps
	
	c = "robocopy.exe "
	c = c & Chr(34) & strFolderSource & Chr(34)
	c = c & " " 
	c = c & Chr(34) & strFolderDest & Chr(34)
	c = c & " "
	c = c & "*.*"
	c = c & " "
	c = c & "/e /z /mov /r:9 /w:10 /np /copy:dt /log:robocopy-collector.log"
	
	WScript.Echo c
	el = RunCommand(c)
	If el <= 8 Then
		WScript.Echo "MoveExportFolderToCollectorDc() SUCCESS"
	Else
		WScript.Echo "MoveExportFolderToCollectorDc() ERROR: " & el
	End If
End Sub '' of Sub MoveExportFolderToCollectorDc



Sub MoveCollectorToSplunkServer(strFolderSource)
	''
	''	Move all files in the Collector share to the Splunk server.
	''	
	''	LPR > \\SPLUNKSERVER\000134-LPR
	''	SKV > \\SPLUNKSERVER\000134-SKV
	''
	Dim		c
	Dim		strFolderDest
	Dim		el				' DOS Error Level
	
	
	'Dim		objFolder
	'Dim		colFiles
	'Dim		objFile
	'Dim		colSubFolders
	'Dim		objSubFolder
	
	'' First move the SKV files to the SKV share on the Splunk server.
	
	WScript.Echo "MoveCollectorToSplunkServer(): Move the " & EXTENSION_SKV
	
	c = "robocopy.exe "
	c = c & Chr(34) & strFolderSource & Chr(34)
	c = c & " " 
	c = c & Chr(34) & SHARE_SKV & "\" & ProperDateFs(Now(), False) & Chr(34) 
	c = c & " "
	c = c & "*" & EXTENSION_SKV
	c = c & " "
	c = c & "/e /z /mov /r:9 /w:10 /np /copy:dt /log:robocopy-skv.log"
	
	WScript.Echo c
	el = RunCommand(c)
	If el <= 8 Then
		WScript.Echo "MoveExportFolderToCollectorDc() SUCCESS"
	Else
		WScript.Echo "MoveExportFolderToCollectorDc() ERROR: " & el
	End If
	
	WScript.Echo "MoveCollectorToSplunkServer(): Move the " & EXTENSION_LPR
	
	c = "robocopy.exe "
	c = c & Chr(34) & strFolderSource & Chr(34)
	c = c & " " 
	c = c & Chr(34) & SHARE_LPR & "\" & ProperDateFs(Now(), False) & Chr(34)
	c = c & " "
	c = c & "*" & EXTENSION_LPR
	c = c & " "
	c = c & "/e /z /mov /r:9 /w:10 /np /copy:dt /log:robocopy-lpr.log"
	
	WScript.Echo c

	el = RunCommand(c)
	If el <= 8 Then
		WScript.Echo "MoveExportFolderToCollectorDc() SUCCESS"
	Else
		WScript.Echo "MoveExportFolderToCollectorDc() ERROR: " & el
	End If
End Sub '' of Sub MoveCollectorToSplunkServer



Sub ProcessEventLog(ByVal strEventLog)
	Dim		dtmLastRun
	Dim		dtmNow
	Dim		strPathExport
	Dim		strPathLpr
	Dim		strPathSkv
	
	dtmLastRun = LastRunGet(gstrComputerName, strEventLog)
	dtmNow = LastRunPut(gstrComputerName, strEventLog)
	
	strPathExport = GetScriptPath() & "\export"
	strPathLpr = strPathExport & "\" & gstrComputerName & "\" & GetUniqueFileName(dtmLastRun) & EXTENSION_LPR

	WScript.Echo "Event Log            : " & strEventLog
	WScript.Echo "Computer             : " & gstrComputerName
	WScript.Echo "Date time - last run : " & dtmLastRun
	WScript.Echo "Date time - now      : " & dtmNow
	WScript.Echo "Path export LPR      : " & strPathLpr
		
	If ExportEventsUsingLogparser(gstrComputerName, dtmLastRun, dtmNow, strPathLpr) = LOGPARSER_OK Then
		strPathSkv = Replace(strPathLpr, EXTENSION_LPR, EXTENSION_SKV)
		If ConvertUsingLpr2skv(strPathLpr, strPathSkv) = 0 Then
		
			'' All DC's need to deliver their exports to the Collector DC
			'If UCase(gstrComputerName) <> UCase(COLLECTOR_DC) Then 
				'WScript.Echo "Move export files to the collector DC"
				'Call MoveExportFolderToCollectorDc(strPathExport)
			'End If
			
			Call ExtractFilePerEvent(strPathExport & "\" & gstrComputerName, strPathLpr)
			
			
			
			If DoesShareExist(COLLECTOR_SHARE) = True Then
				'' Actions that need to be done by the Collector DC
				WScript.Echo "This is the Collector DC, found the " & COLLECTOR_SHARE & " share."
				
				Call MoveCollectorToSplunkServer(strPathExport)
				
			Else
				'' Actions that need to be done by the 
				WScript.Echo "This is not the Collector DC, move files to Collector DC " & COLLECTOR_DC
				Call MoveExportFolderToCollectorDc(strPathExport)
			End If

			' If UCase(gstrComputerName) = UCase(COLLECTOR_DC) Then 
				'WScript.Echo "Move Collector files to REC Splunk server"
				 'Call MoveCollectorToSplunkServer(strPathExport)
			' End If
		Else
			Script.Echo "ERROR conversion!!"
		End If
	Else
		WScript.Echo "No export Logparser"
	End If
	
End Sub '' of Sub ProcessEventLog



Sub ScriptInit()
	If IsThisScriptAlreadyRunning() = True Then
		WScript.Echo "WARNING: Another instance of this script is already running on this computer, stopping this instance!"
		WScript.Quit(0)
	End If

	If Len(GetProgramPath("robocopy.exe")) = 0 Then
		WScript.Echo "WARNING: Could not find robocopy.exe, stopping this instance!"
		WScript.Quit(0)
	End If

	If Len(GetProgramPath("lpr2skv.exe")) = 0 Then
		WScript.Echo "WARNING: Could not find lpr2skv.exe, stopping this instance!"
		WScript.Quit(0)
	End If
	
	Set gobjFso = CreateObject("Scripting.FileSystemObject")
	
	gstrComputerName = GetComputerName()
End Sub '' of Sub ScriptInit



Sub ScriptRun()
	WScript.Echo "export-events.vbs started..."

	'Call ExportEventsUsingLogparser("NS00DC011", "2015-04-08 12:00:00", "2015-04-08 12:30:00", "testfike-NS00DC011.lpr")
	
	Call ProcessEventLog("Security")

	
End Sub '' of Sub ScriptRun



Sub ScriptDone()
	Set gobjFso = Nothing
End Sub '' of Sub ScriptDone



Call ScriptInit()
Call ScriptRun()
Call ScriptDone()
WScript.Quit(0)
