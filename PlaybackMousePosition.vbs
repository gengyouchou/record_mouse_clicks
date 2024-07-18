Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

If Not objFSO.FileExists("mouse_clicks.txt") Then
    MsgBox "mouse_clicks.txt not found!", vbCritical
    WScript.Quit
End If

Set objFile = objFSO.OpenTextFile("mouse_clicks.txt", 1)
Do Until objFile.AtEndOfStream
    line = objFile.ReadLine
    coords = Split(line, ",")
    x = CInt(coords(0))
    y = CInt(coords(1))
    Call SetCursorPos(x, y)
    WScript.Sleep 1000
    Call MouseClick
    WScript.Sleep 1000
Loop
objFile.Close
MsgBox "Playback finished."
WScript.Quit

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
