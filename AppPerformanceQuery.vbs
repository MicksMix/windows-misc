'
' This script can enable performance monitoring
'   on provided executables and will output a 
'   CSV report
'
' Author: Mick Grove
' Date: Sep 2008
'
'
'Option Explicit
On Error Resume Next
		
Dim objFile		
Dim objFolder
Dim objWMIService
Dim objRefresher
Dim objFSO
Dim objItem
Dim objPageFileItem
Dim objSystemMemItem
Dim objArgs

Dim LogFolder
Dim LogFile
Dim i, j, k, m
Dim strComputer
Dim sAppsArray
Dim WshShell
Dim mytime
Dim colItems
Dim colPageFile
Dim colSystemMem
Dim sCurApp
Dim sAppData
Dim sHeaderLine
Dim sLineTest
Dim strTextToWrite
Dim bLineWritten
Dim sLineToWrite
Dim DataList
Dim iNumberofQueries
Dim iQueryInterval
Dim sPfPercent
Dim sSystemMemUsage

Const adVarChar = 200
Const MaxCharacters = 255 
Const ForWriting = 2
Const ForAppending = 8

strComputer = "."

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set WshShell=CreateObject("WScript.shell")
Set DataList = CreateObject("ADOR.Recordset") 'memory dataset

Set objArgs = WScript.Arguments
iNumberofQueries = objArgs(0)
iQueryInterval = objARgs(1)



'========================================
'========================================
'
' You can add processes to the array below
'
        
sAppsArray = _
	Split("cvpnd|" _
		& "vpngui|" _
		& "vpnclient","|")

if Len(iNumberofQueries) < 1 then
	iNumberofQueries = 60
End If

if Len(iQueryInterval) < 1 then
	iQueryInterval = 1 '1 second
End If
'
'
'========================================
'========================================


WScript.Echo "Monitoring performance data for " _
	& iNumberofQueries & " seconds" & vbcrlf

sAppData = WshShell.ExpandEnvironmentStrings("%appdata%")
mytime = GetFormattedTimeString()

LogFolder = sAppData & "\MyAppPerfMonitoring" 
LogFile = LogFolder & "\_" & mytime & "_perfmon.csv"'

If objFSO.FolderExists(LogFolder) Then
	Set objFile = objFSO.CreateTextFile(LogFile)
	objFile.Close
	Set objFile = objFSO.OpenTextFile(LogFile, ForWriting)  
Else
	Set objFolder = objFSO.CreateFolder(LogFolder)
	Set objFile = objFSO.CreateTextFile(LogFile)
	objFile.Close
	Set objFile = objFSO.OpenTextFile(LogFile, ForWriting)  
End If

Set colItems = objRefresher.AddEnum _
	(objWMIService, "Win32_PerfFormattedData_PerfProc_Process").objectSet
Set colPageFile = objRefresher.AddEnum _
	(objWMIService, "Win32_PerfFormattedData_PerfOS_PagingFile").objectSet
Set colSystemMem = objRefresher.AddEnum _
	(objWMIService, "Win32_PerfFormattedData_PerfOS_Memory").objectSet

For i = 0 To UBound(sAppsArray)
	DataList.Fields.Append sAppsArray(i), adVarChar, MaxCharacters
	if i = 0 then
		sHeaderLine = sAppsArray(i) & " [Proc %], " _
			&  sAppsArray(i) & " [Handles], " _
			& sAppsArray(i) & " [WS Memory]"
	else
		sHeaderLine = sHeaderLine & "," _
			& sAppsArray(i) & " [Proc %], " _
			&  sAppsArray(i) & " [Handles], " _
			& sAppsArray(i) & " [WS Memory]"
	end if
Next

objFile.Write "Event Time, System Memory Usage, System PF Usage," _
	& sHeaderLine & vbcrlf
DataList.Open

For j = 0 to iNumberofQueries
	
	objRefresher.Refresh
	
	For Each objItem in colItems
		bLineWritten = False
		sLineToWrite = ""
		
		For k = 0 To UBound(sAppsArray)
			sCurApp = UCase(sAppsArray(k))			
			if  sCurApp = UCase(objItem.Name) then
				DataList.AddNew
				if k = (UBound(sAppsArray) - 1) then
					DataList(sCurApp) = objItem.PercentProcessorTime _
						& "% ," & objItem.HandleCount & "," _
						& FormatNumber((objItem.WorkingSet / 1024 / 1024),2) _
						& " mb"
				else
					DataList(sCurApp) = objItem.PercentProcessorTime _
						& "% ," & objItem.HandleCount & "," _
						& FormatNumber((objItem.WorkingSet / 1024 / 1024),2) _
						& " mb,"
				end if
				DataList.Update
				bLineWritten = True
			end if
		Next
		
		if bLineWritten = True Then
			DataList.MoveFirst
			Do Until DataList.EOF
				For m = 0 To UBound(sAppsArray)
					strTextToWrite = strTextToWrite _
						& DataList.Fields.Item(sAppsArray(m))
				Next
				
				DataList.MoveNext
			Loop
			
			DataList.MoveFirst
			Do While Not DataList.EOF
				DataList.Delete
				DataList.MoveNext
			Loop
		End If
    Next
	
	sLineTest = Replace(strTextToWrite, ",","")
	
	if Len(sLineTest) > 0 then
	
		For Each objPageFileItem in colPageFile
			if objPageFileItem.Name = "_Total" Then
				sPfPercent = objPageFileItem.PercentUsage				
			end if
		Next
		
		For Each objSystemMemItem in colSystemMem
			sSystemMemUsage = objSystemMemItem.AvailableMBytes				
		Next
		
		objFile.Write Now & "," & sSystemMemUsage _
			& " mb," & sPfPercent & "% ," & strTextToWrite & vbcrlf
	end if
	
	strTextToWrite = ""

    Wscript.Sleep (iQueryInterval * 1000) 'convert to milliseconds
Next

Function GetFormattedTimeString
	Dim curTime
	
	curTime = Now
	curTime = Replace(curTime, "/","-") 
	curTime = Replace(curTime, ":","_") 
	curTime = Replace(curTime, " ","_") 
	
	GetFormattedTimeSTring = curTime
End Function

objFile.Close

WScript.Echo "Done. Log located at: " _
	& vbcrlf & "<" & LogFile & ">"