Set objShell = WScript.CreateObject("WScript.Shell")

' Prompt the user for the desired shortcut name
shortcutName = InputBox("Enter a name for the shortcut:", "Create Chrome Shortcut")

' Create the shortcut on the desktop
Set objShortcut = objShell.CreateShortcut(objShell.SpecialFolders("Desktop") & "\" & shortcutName & ".lnk")

' Set the target path and command line arguments
objShortcut.TargetPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
objShortcut.Arguments = "--disable-extensions"

' Save the shortcut
objShortcut.Save

' Close all open Chrome windows and tabs
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name LIKE 'chrome.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next

' Wait for Chrome to close
WScript.Sleep 1000

' Open the new shortcut
objShell.Run """" & objShortcut.FullName & """"

WScript.Sleep 500

' Display a message box with instructions for the user
MsgBox "Shortcut created on desktop. All Chrome windows and tabs have been closed. Launch Chrome using this shortcut to disable all extensions. Remember to close all instances of 'blocked/normal' chrome to use next time. To avoid suspicion, it is encouraged that you delete this .VBS program after finish running it."