' Create a new instance of the WScript.Shell object
Set objShell = WScript.CreateObject("WScript.Shell")
strName = objShell.Popup("Please enter your name:", 0, "Name Input", 0x40000 + 0x1000000, -1, -1)

If strName = "" Then
  WScript.Quit
End If


strAge = objShell.Popup("Please enter your age:", 0, "Age Input", 0x40000 + 0x1000000, -1, -1)

If strAge = "" Then
  WScript.Quit
End If

intAge = CInt(strAge)
strMessage = "Hello, " & strName & "! You are " & intAge & " years old."
objShell.Popup strMessage, 0, "Personalized Message", 0x40000 + 0x1000000, -1, -1
WScript.Quit