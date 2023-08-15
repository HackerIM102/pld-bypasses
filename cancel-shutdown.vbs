Set objShell = CreateObject("WScript.Shell")

Sub RunCommand()
    ' Run the "shutdown /a" command to abort shutdown
    objShell.Run "shutdown /a", 0, False
End Sub

Sub RepeatCommand()
    ' Run the command initially
    RunCommand
    
    ' Repeat the command every 5 minutes
    Do
        ' Sleep for 5 minutes (300,000 milliseconds)
        WScript.Sleep 300000
        
        ' Run the command again
        RunCommand
    Loop
End Sub

' Start repeating the command
RepeatCommand