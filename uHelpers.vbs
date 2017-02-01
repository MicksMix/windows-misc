'
' Basic helper methods I've often used as part of other VBScripts
'

Function GetEnvVariable(sEnv)
	Dim sValue
	Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")	
	sValue = objShell.ExpandEnvironmentStrings("%" & sEnv & "%")

	GetEnvVariable = sValue
End Function

Sub SearchAndBackupAndReplace(sNewFile, sExtension, sName)
	On Error Resume Next
	Dim objWMIService : Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2") 
	Dim objItem
	Dim colItems : Set colItems = objWMIService.ExecQuery( _
		"SELECT Path,Extension,FileName,Name FROM CIM_DataFile Where " _
			"Extension='" & sExtension & _
			"' AND FileName='" & sName & _
			"' AND (Path LIKE '\\Users\\%' OR Path LIKE '\\Program Files%')",,48) 
	For Each objItem in colItems 
		'backup first
		CopyFile objItem.Name,objItem.Name & ".bak"
		'replace
		CopyFile sNewFile,objItem.Name
	Next	
End Sub

Sub CopyFile(SourceFile, DestinationFile)
	On Error Resume Next
    Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")

    'Check to see if the file already exists in the destination folder
    Dim wasReadOnly
    wasReadOnly = False
    If objFSO.FileExists(DestinationFile) Then
        'Check to see if the file is read-only
        If objFSO.GetFile(DestinationFile).Attributes And 1 Then 
            'The file exists and is read-only.
            WScript.Echo "Removing the read-only attribute"
            'Remove the read-only attribute
            objFSO.GetFile(DestinationFile).Attributes = objFSO.GetFile(DestinationFile).Attributes - 1
            wasReadOnly = True
        End If

        WScript.Echo "Deleting the file"
        objFSO.DeleteFile DestinationFile, True
    End If

    'Copy the file
    WScript.Echo "Copying " & SourceFile & " to " & DestinationFile
    objFSO.CopyFile SourceFile, DestinationFile, True

    If wasReadOnly Then
        'Reapply the read-only attribute
        objFSO.GetFile(DestinationFile).Attributes = objFSO.GetFile(DestinationFile).Attributes + 1
    End If

    Set objFSO = Nothing
End Sub

Sub ProcessEveryUser()
	Const HKEY_LOCAL_MACHINE    = &H80000002
	Dim strKeyPath, strValueName, strValue, strSubPath, arrSubKeys 
	Dim objRegistry : Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"

	objRegistry.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys 
	 
	For Each objSubkey In arrSubkeys 
		strValueName = "ProfileImagePath"
		strSubPath = strKeyPath & "\" & objSubkey 
		objRegistry.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue 
		sProfilePath = strValue 
		sCurrentUser = RetrieveUsernameFromPath(strValue) 
	 
		If ((UCase(sCurrentUser) <> "ALL USERS") and _ 
			(UCase(sCurrentUser) <> "LOCALSERVICE") and _ 
			(UCase(sCurrentUser) <> "SYSTEMPROFILE") and _ 
			(UCase(sCurrentUser) <> "NETWORKSERVICE")) then 
				WScript.Echo "Preparing to update the user: " & sCurrentUser
				'Call LoadProfileHive(sPathToDatFile, sCurrentUser, DAT_FILE)
		End If
	Next
End Sub

Function RetrieveUsernameFromPath(sTheProfilePath) 
    On Error Resume Next
    
    Dim lstPath 
    Dim sTmp 
    Dim sUsername 
     
    lstPath = Split(sTheProfilePath,"\") 
    For each sTmp in lstPath 
        sUsername = sTmp 
        'last split is our username 
    Next
     
    RetrieveUsernameFromPath = sUsername 
End Function

Sub LogToEventlog(sMsg, sSource)
	Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")	
	Dim sEventCmd 
	sEventCmd = "eventcreate.exe /L APPLICATION /T INFORMATION /SO " & sSource & " /ID 1844 /D " & chr(34) & sMsg & chr(34)
	objShell.Run sEventCmd,0,True	

End Sub

Sub FindAndDeleteOldFilesByExtension(strPath, strExt, iDays)
    ' Call FindAndDeleteOldFiles("c:\temp", "sql", 30)
    On Error Resume Next
    Dim objFolder, objSubFolder, objFile
    Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.getFolder(strPath)
    
	For Each objFile In objFolder.files
		'If dateDiff("d",objFile.dateLastModified,Now) > 30 Then objFile.Delete
        If dateDiff("d",objFile.dateLastModified,Now) > iDays And right((objFile.Path),4) = "." & strExt Then objFile.Delete
	Next
	For Each objSubFolder In objFolder.SubFolders
		FindAndDeleteOldFilesByExtension(objSubFolder.path, strExt, iDays)
	Next
    
	For Each objSubFolder In objFolder.SubFolders
		If objSubFolder.Files.Count = 0 Then ObjSubFolder.Delete
	Next
End Sub

Function FirstVersionSupOrEqualToSecondVersion(strFirstVersion, strSecondVersion)
	Dim arrFirstVersion,  arrSecondVersion, i, iStop, iMax
	Dim iFirstArraySize, iSecondArraySize
	Dim blnArraySameSize : blnArraySameSize = False
	
	If strFirstVersion = strSecondVersion Then
		FirstVersionSupOrEqualToSecondVersion = True
		Exit Function
	End If
	
	If strFirstVersion = "" Then
		FirstVersionSupOrEqualToSecondVersion = False
		Exit Function
	End If
	If strSecondVersion = "" Then
		FirstVersionSupOrEqualToSecondVersion = True
		Exit Function
	End If

	arrFirstVersion = Split(strFirstVersion, "." )
	arrSecondVersion = Split(strSecondVersion, "." )
	iFirstArraySize = UBound(arrFirstVersion)
	iSecondArraySize = UBound(arrSecondVersion)
	
	If iFirstArraySize = iSecondArraySize Then
		blnArraySameSize = True
		iStop = iFirstArraySize
		For i=0 To iStop
			If CInt(arrFirstVersion(i)) < CInt(arrSecondVersion(i)) Then
				FirstVersionSupOrEqualToSecondVersion = False
				Exit Function
			End If
		Next
		FirstVersionSupOrEqualToSecondVersion = True
	Else
		If iFirstArraySize > iSecondArraySize Then
			iStop = iSecondArraySize
		Else
			iStop = iFirstArraySize
		End If
		For i=0 To iStop
			If CInt(arrFirstVersion(i)) < CInt(arrSecondVersion(i)) Then
				FirstVersionSupOrEqualToSecondVersion = False
				Exit Function
			End If
		Next
		If iFirstArraySize > iSecondArraySize Then
			FirstVersionSupOrEqualToSecondVersion = True
			Exit Function
		Else
			For i=iStop+1 To iSecondArraySize
				If CInt(arrSecondVersion(i)) > 0 Then
					FirstVersionSupOrEqualToSecondVersion = False
					Exit Function
				End If
			Next
			FirstVersionSupOrEqualToSecondVersion = True
		End If
	End If
	
	REM =========================================
	REM =
	REM =		USAGE
	REM =		
	REM =========================================
	REM =
	REM sInstalledVersion = GetInstalledSoftwareVersion("%Cisco%AnyConnect%Secure%Client%")
	REM Dim bInstall
	REM bInstall = False ' default

	REM CONST NEW_VERSION = "3.1.04072"
	REM If sInstalledVersion <> "0.0.0" Then
		REM If FirstVersionSupOrEqualToSecondVersion(sInstalledVersion,NEW_VERSION) Then
			REM WScript.Echo sInstalledVersion & " >= " & NEW_VERSION
			REM ' current or newer version already installed
			REM ' delete scheduled task and quit
			
			REM WScript.Quit
		REM Else
			REM WScript.Echo sInstalledVersion & " < " & NEW_VERSION
			REM 'BeginInstallation()
			REM bInstall = True
		REM End If	
	REM End If
End Function

Function GetInstalledSoftwareVersion(sWqlSoftwareName)
	REM sWqlSoftwareName = "%Cisco%AnyConnect%Secure%Client%"
	Dim strComputer 
	strComputer = "."
	Dim objWMIService : Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
	Dim colItems : Set colItems = objWMIService.ExecQuery( _
		"SELECT Version FROM Win32_Product WHERE Name LIKE '" & sWqlSoftwareName & "'",,48)
 
	Dim sVersion
	Dim objItem
		
	For Each objItem in colItems 
		sVersion = objItem.Version
		If Not IsBlank(sVersion) Then
			Exit For
		End If	
	Next
	
	If IsBlank(Trim(sVersion)) Then
		sVersion = "0.0.0"
	End If
	
	GetAnyConnectVersion = sVersion
End Function

Function GetThisScriptsDirectory()	
	Dim objFSO : Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")	
	Dim objFile : Set objFile = objFSO.GetFile(WScript.ScriptFullName)
	Dim strFolder
	strFolder = objFSO.GetParentFolderName(objFile) 

	GetThisScriptsDirectory = RemoveTrailingPathDelimiter(strFolder)
End Function

Sub RunHiddenAndWait(sCmd)
	'WScript.Echo("running: " & sCmd)
	'LogToEventlog("Executing command: " & sCmd)
	Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")	
	objShell.Run sCmd,0,True	
End Sub

Function ConsoleExecAndCapture(sCmd)
	Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")
	Dim objExecObject : Set objExecObject = objShell.Exec(sCmd)
	Dim sCurLine
	Dim strText
	strText = ""

	Do While Not objExecObject.StdOut.AtEndOfStream
		sCurLine = objExecObject.StdOut.ReadLine()
		If Not IsBlank(sCurLine) Then
			strText = strText & vbcrlf & sCurLine
		End If
	Loop

	ConsoleExecAndCapture = strText
End Function

Function RemoveTrailingPathDelimiter(sPath)
	Dim sUpdatedPath
	sUpdatedPath = sPath

	If Right(sUpdatedPath,1) = "\" Then
		sUpdatedPath = Left(sUpdatedPath,Len(sUpdatedPath)-1)
	End If

	RemoveTrailingPathDelimiter = sUpdatedPath
End Function

Function MakeSureDirectoryTreeExists(dirName)
    Dim aFolders, newFolder, i
    Dim oFS
    
    Set oFS = CreateObject("Scripting.FileSystemObject")

   ' Check the folder's existence
   If Not oFS.FolderExists(dirName) Then
      ' Split the various components of the folder's name
      aFolders = split(dirName, "\")

      ' Get the root of the drive
      newFolder = oFS.BuildPath(aFolders(0), "\")

      ' Scan the various folder and create them
      For i = 1 To UBound(aFolders)
         newFolder = oFS.BuildPath(newFolder, aFolders(i))

         If Not oFS.FolderExists(newFolder) Then
            oFS.CreateFolder newFolder
         End If
      Next
   End If
	
	Set oFS = Nothing
End Function

Function IsBlank(Value)	
	'returns True if Empty or NULL or Zero
	If IsEmpty(Value) or IsNull(Value) Then
		IsBlank = True
	ElseIf VarType(Value) = vbString Then
		If Value = "" Then
			IsBlank = True
		End If
	ElseIf IsObject(Value) Then
		If Value Is Nothing Then
			IsBlank = True
		End If
	ElseIf IsNumeric(Value) Then
		If Value = 0 Then
			'wscript.echo " Zero value found"
			IsBlank = True
		End If
	Else
		IsBlank = False
	End If
End Function