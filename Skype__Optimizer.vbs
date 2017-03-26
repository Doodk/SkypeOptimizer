'Skype's optimization
'ver 1.02
'By doodk
'2016-2-03

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

if fso.FolderExists(AppData) = False then
	msgbox "Cannot find Skype profiles. (installed?)" & vbCrlf & "Please login Skpye account at least once.", 48, "Error"
	WScript.Quit
end if
set f = fso.GetFolder(AppData)
set subf = f.SubFolders
dim foundxml
for each subFolder in subf
	dim configPath
	configPath = subFolder.Path & "\config.xml"
	if fso.FileExists(configPath) then
		foundxml=true
		set origFile = fso.OpenTextFile(configPath)
		dim tmpFile
		tmpFile = objShell.ExpandEnvironmentStrings("%TEMP%") & "\SkypeConfig.tmp"
		set newFile = fso.CreateTextFile(tmpFile, true)
		dim modifiedXML
		modifiedXML = 0
		do Until origFile.AtendofStream
			strLine = origFile.ReadLine

			if InStr(strLine,"<AdvertPlaceholder>1</AdvertPlaceholder>") > 0  then
				newFile.WriteLine(Replace(strLine,"1","0"))
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

if foundxml = False then
	msgbox "Cannot find Skype profiles." & vbCrlf & "Please login Skpye account at least once.", 48, "Error"
	WScript.Quit
end if



'block AD sources (change HOSTS file)
set hostFile = fso.OpenTextFile(objShell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\drivers\etc\hosts")
dim arr
arr = Array("0","0","0")

do Until hostFile.AtendofStream
	strLine = hostFile.ReadLine
	if InStr(strLine,"rad.msn.com") > 0  then
		arr(0) = 1
	elseif InStr(strLine,"rad.live.com") > 0  then
		arr(1) = 1
	elseif InStr(strLine,"apps.skype.com") > 0  then
		arr(2) = 1
	end if
loop
hostFile.Close()

dim modifiedHost
modifiedHost=0
if arr(0)+arr(1)+arr(2) < 3 then
	msgbox "This script will block Skype's AD in HOSTS file." & vbCrlf & "Anti-virus Software may alert or block this operation." & vbCrlf & vbCrlf & "Please select [Allow] if necessary.", 48, "Notice!"
	set hostFile = fso.OpenTextFile(objShell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\drivers\etc\hosts", 8, false)
	hostFile.WriteLine()
	if arr(0) = 0 then
		hostFile.WriteLine("127.0.0.1 rad.msn.com")
	end if
	if arr(1) = 0 then
		hostFile.WriteLine("127.0.0.1 rad.live.com")
	end if
	if arr(2) = 0 then
		hostFile.WriteLine("127.0.0.1 apps.skype.com")
	end if
	hostFile.Close()
	modifiedHost = 1
end if

dim runRunMsg
runRunMsg = vbCrlf & vbCrlf & "[If login onto any new Skype accounts on this computer, re-run this script.]" & vbCrlf & vbCrlf & "This script is for educational purpose ONLY, please run <Skype__Optimizer_UNDO.vbs> to restore any changes."

if modifiedXML=0 and modifiedHost=0 then
	msgbox "Skype's optimization was already made! Scrpit doesn't make any changes." & runRunMsg, 64, "Remain optimized!"
elseif modifiedXML=1 and modifiedHost=0 then
	msgbox "Successfully optimizing Skype's (new account), Enjoy!" & runRunMsg, 64, "Successful!"
elseif modifiedXML=0 and modifiedHost=1 then
	msgbox "Successfully optimizing Skype's (HOST only), Enjoy!" & runRunMsg, 64, "Successful!"
else
	msgbox "Successfully optimizing Skype, Enjoy!" & runRunMsg, 64, "Successful!"
end if




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


