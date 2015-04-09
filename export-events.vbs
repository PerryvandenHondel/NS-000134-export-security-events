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
''		01	2015-04-08	First version
'' 
''	SUBS AND FUNCTIONs:
''		Function ExportEventsUsingLogparser
''		Function GetComputerName
''		Function GetFileSize
''		Function GetPathLastRun
''		Function GetScriptPath
''		Function IsThisScriptAlreadyRunning
''		Function LastRunGet
''		Function LastRunPut
''		Function GetUniqueFileName
''		Function NumberAlign
''		Function ProperDateTime
''		Function RunCommand
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
Const	LOGPARSER_OK = 		0
Const	LOGPARSER_FAIL = 			1
Const	LOGPARSER_SEPARATOR	=		"|"

''	-----------------------------------------------------------------------------------------------
''	GLOBAL VARIABLES
''	-----------------------------------------------------------------------------------------------




Dim		gobjFso
Dim		gstrComputerName
Dim		gdtmInitLastRun




''	-----------------------------------------------------------------------------------------------
''	FUNCTIONS
''	-----------------------------------------------------------------------------------------------


Function GetUniqueFileName()
	'
	'	Generate a file name of hex numbers of length 32 chars.
	'
	'	First 12 digits are the current year, month, day, hour, minutes and seconds 
	'	Extra unique chars to fill-up to 32 chars length.
	'
	'	Function does not add a extension to the file name.
	'
	'	Result: YYYYMMDDHHMMSS
	'
	
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
	
	dtmNow = Now()
	
	''         123456789012345678901234567890123456789
	'' Format: SYSTEMNAME-SEC-YYYYMMDD-HHMMSS-XXXXXXXX
	
	''				12345678901234567890123456789012
	'' strPrefix =	YYYYMMDDHHMMSSxxxxxxxxxxxxxxxxxx
	
	''strPrefix = Left(strEventLogName, 3)
	strPrefix = Year(dtmNow) & NumberAlign(Month(dtmNow), 2) & NumberAlign(Day(dtmNow), 2)
	strPrefix = strPrefix & NumberAlign(Hour(dtmNow), 2) & NumberAlign(Minute(dtmNow), 2) & NumberAlign(Second(dtmNow), 2)
	
	'For i = 1 to HEX_LEN - Len(strPrefix)
	
	For i = 1 to UNIQUE_LEN - Len(strPrefix)
		Randomize
		intNumber = Int((NUM_HIGH - NUM_LOW + 1) * Rnd + NUM_LOW)
		strFilename = strFilename & LCase(Hex(intNumber))
	Next
	
	GetUniqueFileName = strPrefix & strFilename
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



Function GetPathLastRun(ByVal strComputerName, ByVal strEventLog)
	''
	''	Returns the path for the lastrun file (scriptpath\release\computer\last-run.txt)
	''	
	Dim		r
	
	r = GetScriptPath() & "\lastrun\" & UCase(strComputerName) & "-" & strEventLog & ".lrdt"
	Call MakeFolder(r)
	GetPathLastRun = r
End Function

'	-----------------------------------------------------------------------------------------------

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
		If GetFileSize(strPathLogparser) > 0 Then
			WScript.Echo "  SUCCESS: File " & strPathLogparser & " contains data. intReturn=1"
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
End Sub '' of Sub bMakeFolder



Sub ProcessEventLog(ByVal strEventLog)

	Dim		dtmLastRun
	Dim		dtmNow
	Dim		strPathExport
	Dim		strPathLpr
	
	dtmLastRun = LastRunGet(gstrComputerName, strEventLog)
	dtmNow = LastRunPut(gstrComputerName, strEventLog)
	
	strPathExport = GetScriptPath() & "\export"
	strPathLpr = strPathExport & "\" & gstrComputerName & "\" & GetUniqueFileName & EXTENSION_LPR

	WScript.Echo "EventLog             : " & strEventLog
	WScript.Echo "Computer             : " & gstrComputerName
	WScript.Echo "Date time - previous : " & dtmLastRun
	WScript.Echo "Date time - now      : " & dtmNow
	WScript.Echo "Path export LPR      : " & strPathLpr
	
	If ExportEventsUsingLogparser(gstrComputerName, dtmLastRun, dtmNow, strPathLpr) = LOGPARSER_OK Then
		WScript.Echo "Convert the file " & strPathLpr
	Else
		WScript.Echo "No export Logparser"
	End If
	
End Sub '' of Sub ProcessEventLog



Sub ScriptInit()
	If IsThisScriptAlreadyRunning() = True Then
		WScript.Echo "WARNING: Another instance of this script is already running on this computer, stopping this instance!"
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
