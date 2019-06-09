' BKUSER script v1.0 - June 9th, 2019 - Danilo Chiavari (www.danilochiavari.com)
' ------------------------------------------------------------------------------
'
'This script was born to facilitate backup policy creations for workstations in Veeam Backup & Replication / Veeam Agent for Windows.
'
'The script does the following:
'
'  -  Parses a list of computers (hostnames or IPs) from a CSV file ("list.csv"), then for each of them:
'  	  -  Retrieves the last logged on user name from the Windows registry
'	  -  Sets an environmental variable ("bkuser") in the system context, with the last logged on user name as value
'
'After doing this, it is easy to leverage the "bkuser" environmental variable when creating a backup policy from within Veeam Backup & Replication (or standalone Veeam Agent as well).
'
'For example: in order to back up only the 'Desktop' folders, "C:\Users\\%bkuser%\Desktop" can be used.

'USAGE: open an elevated Command Prompt (CMD.EXE / Run as Administrator) and run the script directly.
'Running the script with 'cscript' is recommended, as all "logging" (wscript.echo) output would be written to console (command prompt window) rather than pop-up boxes.
'
'EXAMPLE: cscript C:\scripts\bkuser\bkuser.vbs 

Const HKEY_LOCAL_MACHINE = &H80000002

Set objFSO = CreateObject("Scripting.FileSystemObject")	'create object to read input CSV (text) file
Set inputFile = objFSO.OpenTextFile("list.csv", 1)		'open CSV file to parse computer list
ComputerList = Split(inputFile.ReadAll, vbCrLf)			'populate ComputerList array with entries from CSV file, using newlines (vbCrLf) as delimiters

For Each Computer in ComputerList
	
	wscript.echo "Processing computer: " & Computer
	
	'Set variables needed to read the value from the registry
	Set objRegistry = GetObject("winmgmts:\\" & Computer & "\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"
	strValueName = "LastLoggedOnUser"

	objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue							'Read the value from the registry
	wscript.echo "Retrieved LastLoggedOnUser key from registry: " & strValue & " for computer: " & Computer
	
	SlashPosition = InStr(strValue,"\")						'look for first occurrence of slash (\) inside string. Store position in variable
	AmpersandPosition = InStr(strValue,"@")					'look for first occurrence of ampersand (@) inside string. Store position in variable

	If SlashPosition > 0 then 								'if username string contains a slash...
		LastUsername = Mid(strValue,SlashPosition+1)		'... cut out the domain name that precedes the actual username
	ElseIf AmpersandPosition > 0 then 						'if username contains an ampersand...
		LastUsername = Left(strValue,AmpersandPosition-1)	'... cut out the domain name that follows the actual username
	Else 
		LastUsername = "ERROR - Something's fishy here... LastLoggedOnUser appears to be '"&strValue&"' ?? Exiting..."	'if no slash or ampersand is found, something's not quite right
		wscript.echo LastUsername
		wscript.quit
	End If
	
	Set objShell = CreateObject("WScript.Shell")										'create object to execute a command
	cmdStringSetVar = "SETX " & "bkuser " & LastUsername & " /M"						'define the command: 'SETX bkuser <username> /M'
	wscript.echo "Sending command: '" & cmdStringSetVar & "' to computer: " & Computer	'log the command to console before executing it
	
	Set sobjWMIService = GetObject("winmgmts:\\" & Computer & "\root\cimv2:Win32_Process")	'define WMI object needed to run command remotely
    sintReturn = sobjWMIService.Create(cmdStringSetVar, null, null, sintProcessID)			'run the command and store return code in variable 'sintReturn'
    Select Case sintReturn
        Case 0 'Successful Completion
            wscript.Echo("Successful Completion")
        Case 2 'Access Denied
			wscript.Echo("Access Denied")
        Case 3 'Insufficient Privilege
            wscript.Echo("Insufficient Privilege")
        Case 8 'Unknown Failure
            wscript.Echo("Unknown Failure")
        Case 9 'Path not found
            wscript.Echo("Path not found")
        Case 21 'Invalid Parameter
            wscript.Echo("Invalid Parameter")
        Case Else 
            wscript.Echo("Error code " & sintReturn & " not found")
    End Select 
	
	wscript.echo vbNewLine
	
Next