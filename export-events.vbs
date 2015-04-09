''
''	export-events.vbs
''
''
''
''
''		Function ExportEventsUsingLogparser
''		Function GetComputerName
''		Function GetFileSize
''		Function GetPathLastRun
''		Function GetScriptPath
''		Function LastRunGet
''		Function LastRunPut
''		Function RunCommand
''		Sub ScriptDone
''		Sub ScriptInit
''		Sub ScriptRun
'' 



Option Explicit





Dim		gobjFso



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



Function GetPathLastRun(ByVal strComputer)
	'
	'	Returns the path for the lastrun file (scriptpath\release\computer\last-run.txt)
	'	
	'	V17: Build the path of the script path with a last run folder and file name.
	'	V18: Add the gstrScriptRelease part.
	'
	'	GetPathLastRun = GetScriptPath() & "\LastRun\" & strComputer & ".txt"
	Dim		p
	
	p = GetScriptPath() & "\export" & "\" & strComputer & "\last-run.txt"
	
	Call MakeFolder(p)
	GetPathLastRun = p
End Function

'	-----------------------------------------------------------------------------------------------

Function LastRunGet(ByVal strHostName)
	'
	'	Read the last run date time from a text file.
	'	When the file does not exists. Create a file.
	'	Place a date time of previous day in it.
	'
	Dim		objFile
	Dim		strFileName
	Dim		strLine
	Dim		dtmYesterday
	Dim		strPath
	
	strPath = GetPathLastRun(strHostName)
	
	If gobjFso.FileExists(strPath) = True Then
		Set objFile = gobjFso.OpenTextFile(strPath, FOR_READING)
		strLine = ProperDateTime(objFile.ReadLine)
		objFile.Close
		Set objFile = Nothing
	Else
		Call MakeFolder(strPath)
		'Wscript.Echo "LastRunGet() WARNING: File " & strPath & " not found, creating one with date " & gdtmInitLastRun
		Set objFile = gobjFso.OpenTextFile(strPath, FOR_WRITING, True)
		strLine = ProperDateTime(gdtmInitLastRun)
		objFile.WriteLine strLine
		objFile.Close
		Set objFile = Nothing
	End If
	LastRunGet = strLine
End Function

'	-----------------------------------------------------------------------------------------------

Function LastRunPut(ByVal strHostName)
	'
	'	Put the current date time in a the last run file.
	'	Returns the current date time
	'
	Dim		objFile
	Dim		strFileName
	Dim		dtmNow
	Dim		strLine
	Dim		strPath
	
	strPath = GetPathLastRun(strHostName)
	
	
	If gobjFso.FileExists(strPath) = True Then
		Set objFile = gobjFso.OpenTextFile(strPath, FOR_WRITING, True)
		
		'	Get the current date time.
		dtmNow = ProperDateTime("")
		
		'' WScript.Echo "LastRunPut(): Writing new date time " & dtmNow & " to file " & strFileName
		objFile.WriteLine ProperDateTime("")
		objFile.Close
		Set objFile = Nothing
	End If
	LastRunPut = dtmNow
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
	'
	'	Export the events to a temp file specified in strPathLogparser
	'
	'	Return
	'		0		Noting was exported
	'		1		Export was successful and export file contains data. (file size > 0)
	'
	
	Const	LOGPARSER_SEPARATOR	= "|"
	Const	LOGPARSER_FAIL	= 1
	Const	LOGPARSER_SUCCESS = 0
	
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
	
	''Call LogWrite("-")
	WScript.Echo c
	'' Call LogWrite("-")
	
	r = RunCommand(c)
	If r = 0 Then
		''Call LogWrite("  SUCCESS: Logparser exported success full the events: r=" & r)
		If GetFileSize(strPathLogparser) > 0 Then
			WScript.Echo "  SUCCESS: File " & strPathLogparser & " contains data. intReturn=1"
			intReturn = LOGPARSER_SUCCESS
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

'	-----------------------------------------------------------------------------------------------




Sub ScriptInit()
	Set gobjFso = CreateObject("Scripting.FileSystemObject")
End Sub '' of Sub ScriptInit



Sub ScriptRun()
	WScript.Echo "export-events.vbs started..."

	Call ExportEventsUsingLogparser("NS00DC011", "2015-04-08 12:00:00", "2015-04-08 12:30:00", "testfike-NS00DC011.lpr")
	

End Sub '' of Sub ScriptRun



Sub ScriptDone()
	Set gobjFso = Nothing
End Sub '' of Sub ScriptDone



Call ScriptInit()
Call ScriptRun()
Call ScriptDone()
WScript.Quit(0)
