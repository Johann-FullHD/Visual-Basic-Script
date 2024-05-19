' Create a new instance of the WScript.Shell object
Set objShell = WScript.CreateObject("WScript.Shell")

' Create a new instance of the InputBox function
strFolderName = objShell.Popup("Please enter the name of the new folder:", 0, "New Folder", 0x40000 + 0x1000000, -1, -1)

' Check if the user clicked "Cancel"
If strFolderName = "" Then
  WScript.Quit
End If

' Create a new folder
objShell.Run "cmd /c mkdir ""C:\temp\" & strFolderName & """", 0, True

' Create a new instance of the InputBox function
strFileName = objShell.Popup("Please enter the name of the file to copy:", 0, "Copy File", 0x40000 + 0x1000000, -1, -1)

' Check if the user clicked "Cancel"
If strFileName = "" Then
  WScript.Quit
End If

' Copy the file
objShell.Run "cmd /c copy ""C:\temp\example.txt"" ""C:\temp\" & strFileName & """", 0, True

' Create a new instance of the InputBox function
strFilePath = objShell.Popup("Please enter the path of the file to delete:", 0, "Delete File", 0x40000 + 0x1000000, -1, -1)

' Check if the user clicked "Cancel"
If strFilePath = "" Then
  WScript.Quit
End If

' Delete the file
objShell.Run "cmd /c del """ & strFilePath & """", 0, True

' Display a message box to confirm the tasks were completed
objShell.Popup "The tasks were completed successfully!", 0, "Task Completed", 0x40000 + 0x1000000, -1, -1

' Quit the script
WScript.Quit