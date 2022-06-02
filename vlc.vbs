Set objShell = CreateObject("Wscript.Shell")
Set objArguments = WScript.Arguments
Set objFSO = CreateObject("Scripting.FileSystemObject")

If (objArguments.Count = 1) Then
	If (objFSO.FolderExists(objArguments(0))) Then
		For Each f In objFSO.GetFolder(objArguments(0)).Files
			objShell.Run "vlc --one-instance --playlist-enqueue " + Chr(34) + objArguments(0) + "\" + f.Name + Chr(34)
		Next
	Else
		MsgBox "Could not find folder " + Chr(34) + objArguments(0) + Chr(34), VBOKOnly, "Folder not found"
	End If
Else
	MsgBox "Please start the script with the folder to use as the first parameter.", VBOKOnly, "Missing Parameter"
End If
