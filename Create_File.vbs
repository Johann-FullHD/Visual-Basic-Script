Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("C:\temp\myfile.txt", True)

objFile.WriteLine("Hello, World!")
objFile.WriteLine("This is a test file.")

objFile.Close()

MsgBox "The file was created successfully!"