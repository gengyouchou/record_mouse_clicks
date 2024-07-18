Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check if mouse_clicks.txt exists
If Not objFSO.FileExists("mouse_clicks.txt") Then
    MsgBox "mouse_clicks.txt not found!", vbCritical
    WScript.Quit
End If

' Prompt user to start playback
response = MsgBox("Click OK to start playback. Click Cancel to quit.", vbOKCancel)
If response = vbCancel Then
    WScript.Quit
End If

' Open mouse_clicks.txt for reading
Set objFile = objFSO.OpenTextFile("mouse_clicks.txt", 1)

' Loop through each line in the file
Do Until objFile.AtEndOfStream
    line = objFile.ReadLine
    coords = Split(line, ",")
    x = CInt(coords(0))
    y = CInt(coords(1))
    
    ' Set cursor position
    Call SetCursorPos(x, y)
    WScript.Sleep 1000
    
    ' Perform mouse click
    Call MouseClick
    WScript.Sleep 1000
Loop

objFile.Close

' Prompt user to restart playback
response = MsgBox("Playback finished. Click OK to restart playback, or Cancel to quit.", vbOKCancel)
If response = vbOK Then
    ' Restart the script (you might need to implement the logic to restart the script here)
    ' Example: WScript.Run "cscript.exe YourScriptName.vbs"
Else
    WScript.Quit
End If

' Function to set cursor position
Sub SetCursorPos(x, y)
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "powershell -ExecutionPolicy Bypass -File .\SetMousePosition.ps1 -x " & x & " -y " & y, 0, True
End Sub

' Function to perform a mouse click
Sub MouseClick()
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "powershell -ExecutionPolicy Bypass -File .\MouseClick.ps1", 0, True
End Sub
