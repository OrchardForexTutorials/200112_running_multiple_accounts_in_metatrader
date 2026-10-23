'
'	Script to create Metatrader shortcuts
'	Copyright 2020 Novateq Pty Ltd, OrchardForex and FXCHartWorks
'
'	Creates shortcuts with names matching the containing folder in
'		The folder (only if not under program files or program files (x86)
'		The desktop if requested
'		The start menu (only if running with administrator rights
'

' Create a scripting shell object, needed for the rest of the script
Set ws = WScript.CreateObject("WScript.Shell")

' Create a filesystemobject, we need this to locate files
Set fso = CreateObject("Scripting.FileSystemObject")

' Are we running as administrator, if not should we change
adminPrivilege	=	false
If WScript.Arguments.Named.Exists("elevated") = False Then

	if Ask("The script is not running as administrator. You will not be able to create a start menu shortcut. Do you want to restart with administrator privileges? (You must be a member of the administrators group)", "Run as Administrator?") Then
		'Launch the script again as administrator
		CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
		WScript.Quit
	End If
Else
	'Change the working directory from the system32 folder back to the script's folder.
	ws.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
	adminPrivilege	=	true
End If

' Get the current folder
set fld = fso.GetFile(WScript.ScriptFullName).ParentFolder
fldPath = fld.Path & "\"

' This line gets the path to the executing vbscript. The script needs to be in the same folder
' where the terminal.exe and terminal.ico are located and where the shortcut
' will be created

'	Find the exe file. By default this will be terminal.exe for MT4 and terminal64 for MT5
'
'	First look for a default terminal.exe file
terminalExe = ""
If fso.FileExists(fldPath & "Terminal.exe") Then
	terminalExe = "Terminal.exe"
ElseIf fso.FileExists(fldPath & "Terminal64.exe") Then
	terminalExe = "Terminal64.exe"
End If
terminalExe =fldPath & terminalExe

'	make sure the file exists
if Not fso.FileExists(terminalExe) Then
	msgbox "Could not create a shortcut to " & terminalExe & ". File does not exist"
	WScript.Quit
End If


'	Find the ico file. By default this will be terminal.ico
'
'	First look for a default terminal.ico file
terminalIco = ""
If fso.FileExists(fldPath & "Terminal.ico") Then
	terminalIco = "Terminal.ico"
End If
terminalIco = fldPath & terminalIco

If terminalIco <> "" Then

	If LCase(Right(terminalIco,4)) <> ".ico" Then
		terminalIco = terminalIco & ".ico"
	End If

'	make sure the file exists
	if Not fso.FileExists(terminalIco) Then
		msgbox "Icon file " & terminalIco & " does not exist. Creating shortcut with the default icon"
		terminalIco = ""
	End If

End If

' Ask about some options
createLocal = not IsFolderProtected(fldPath)
createPortable = Ask("Create this as a portable link?", "Create Portable Link")
createDesktop = Ask("Do you want to create a desktop icon?", "Create Desktop Icon")
If adminPrivilege Then
	createStartMenu = Ask("Do you want to create a start menu icon?", "Create Start Menu Icon")
Else
	createStartMenu = false
End If

If createPortable Then
	linkArguments = "/portable"
Else
	linkArguments = ""
End If

' Get the name of the current folder, this will also be the shortcut name
' If you don't want to use the folder name then change this line
linkName = fld.Name & ".lnk"

if createLocal then
	CreateShortcut fldPath & linkName, terminalExe, terminalIco, linkArguments
end if

' Should we create a desktop icon
If createDesktop then
	CreateShortcut ws.SpecialFolders("Desktop") & "\" & linkName, terminalExe, terminalIco, linkArguments
End If

' Should we create a start menu icon
If createStartMenu Then
	CreateShortcut ws.SpecialFolders("AllUsersStartMenu") & "\Programs\MT\" & linkName, terminalExe, terminalIco, linkArguments
End If

Function Ask(msg, title)

	Ask = (MsgBox(msg, vbYesNo, title) = vbYes)

End Function

function IsFolderProtected(path)

	IsFolderProtected =  IsChildFolderOf(path, ws.ExpandEnvironmentStrings("%PROGRAMFILES%")) or IsChildFolderOf(path, ws.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%"))
	
end function

function IsChildFolderOf(path, parentpath)

	IsChildFolderOf = (lcase(left(path, len(parentpath))) = lcase(parentpath))
	
end function

Function CreateShortcut(name, exePath, iconPath, arguments)

'	Make sure the parent folder exists
	CreateParentFolder fso.GetParentFolderName(name)
	
'	Create the link
	Set link = ws.CreateShortcut(name)

' 	Set the path to the exe file and the icon file.
	link.TargetPath = exePath
	link.IconLocation = iconPath
	link.Arguments = arguments

' 	Finally save the link
	link.Save

End Function

Function CreateParentFolder(path)

	If fso.FolderExists(path) Then
		Exit Function
	Else
		CreateParentFolder(fso.GetParentFolderName(path))
		fso.CreateFolder(path)
	End If

End Function
