'undo Skype's optimization
'ver 1.0
'By doodk
'2017-3-26

'check if skype is running
dim loopEnd
do
	strComputer = "."
	set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'Skype.exe'")
	loopEnd = colProcessList.Count
	if colProcessList.Count > 0 then
		dim rtn
		rtn = msgbox("Skype is running... Please select options blow." & vbCrlf & vbCrlf & "Abort = Abort Skype's process and continue on this script." & vbCrlf & "Retry = Recheck whether Skype is running." & vbCrlf & "Ignore = Exit this script.", 2+48+256, "Skype is Running...")
		if rtn = 3 then	'abort
			for each uu in colProcessList
				uu.terminate
				exit do
			next
		elseif rtn = 5 then	'ignore
			WScript.Quit
		end if
	end if
loop Until loopEnd=0


'check if script is elevated. (call sub, sub is at the end of script)
Call UACself


'hide AD bar in Skype (change in [username]\config.xml)
set objShell = WScript.CreateObject("WScript.Shell")
set fso = CreateObject("Scripting.FileSystemObject")
dim AppData
AppData = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Skype"

if fso.FolderExists(AppData) = True then
	set f = fso.GetFolder(AppData)
	set subf = f.SubFolders
	for each subFolder in subf
		dim configPath
		configPath = subFolder.Path & "\config.xml"
		if fso.FileExists(configPath) then
			set origFile = fso.OpenTextFile(configPath)
			dim tmpFile
			tmpFile = objShell.ExpandEnvironmentStrings("%TEMP%") & "\SkypeConfig.tmp"
			set newFile = fso.CreateTextFile(tmpFile, true)
			dim modifiedXML
			modifiedXML = 0
			do Until origFile.AtendofStream
				strLine = origFile.ReadLine

				if InStr(strLine,"<AdvertPlaceholder>0</AdvertPlaceholder>") > 0  then
					newFile.WriteLine(Replace(strLine,"0","1"))
					modifiedXML = 1
				Else
					newFile.WriteLine(strLine)
				end if
			loop
			origFile.Close()
			newFile.Close()

			if modifiedXML=1 then
				fso.CopyFile tmpFile, configPath, True
			end if
		end if
	next
end if


'unblock xx sources (change HOSTS file)
set hostFile = fso.OpenTextFile(objShell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\drivers\etc\hosts", 1)
dim arr
arr = Array("0","0","0")
dim hostAllText
hostAllText = hostFile.ReadAll()

if InStr(hostAllText,"rad.msn.com") > 0  then
	arr(0) = 1
elseif InStr(hostAllText,"rad.live.com") > 0  then
	arr(1) = 1
elseif InStr(hostAllText,"apps.skype.com") > 0  then
	arr(2) = 1
end if
hostFile.Close()

if arr(0)+arr(1)+arr(2) > 0 then
	msgbox "The script will unblock Skype's optimization related address in HOSTS file." & vbCrlf & "Anti-virus Software may alert or block this operation." & vbCrlf & vbCrlf & "Please select [Allow] if necessary.", 48, "Notice!"
	
	hostAllText = Replace(hostAllText, "127.0.0.1 rad.msn.com", "")
	hostAllText = Replace(hostAllText, "127.0.0.1 rad.live.com", "")
	hostAllText = Replace(hostAllText, "127.0.0.1 apps.skype.com", "")

	Set hostFileWrite = fso.OpenTextFile(objShell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\drivers\etc\hosts", 2)
	hostFileWrite.WriteLine(hostAllText)
	hostFileWrite.Close()
end if

msgbox "Successfully UNDO Skype's optimization!", 64, "Successful!"


' check and elevate script itself
Sub UACself
	if WScript.Arguments.length =0 then
		set objShell = CreateObject("Shell.Application")

		objShell.ShellExecute "wscript.exe", """" & _
		WScript.ScriptFullName & """ AutoElevate", "", "runas", 1
		WScript.Quit
	Else
		if WScript.Arguments(0) <> "AutoElevate" then
			msgbox "No arguments allowed.", 16,"Error"
			WScript.Quit
		end if
	end if
end Sub

