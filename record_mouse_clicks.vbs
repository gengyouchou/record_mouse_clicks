Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Ask user if they want to record or play back mouse clicks
msg = "Click OK to start recording mouse clicks. Click Cancel to start playback."
response = MsgBox(msg, vbOKCancel)
If response = vbCancel Then
    ' Playback mode
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
Else
    ' Recording mode
    Set objFile = objFSO.CreateTextFile("mouse_clicks.txt", True)
    Do
        If MsgBox("Click OK to record a mouse position. Click Cancel to stop.", vbOKCancel, "Mouse Click Recorder") = vbCancel Then
            Exit Do
        End If
		
		 ' Wait for 5 seconds
        WScript.Sleep 5000

        ' Get cursor position
        GetCursorPos x, y

        ' Write the coordinates to the output file
        objFile.WriteLine x & "," & y

        ' Prompt user that recording is done for this click
        MsgBox "Mouse click recorded at (" & x & ", " & y & "). Click OK to record another click, or Cancel to finish.", vbOKCancel, "Mouse Click Recorder"
    Loop While MsgBox("Click OK to record another mouse position. Click Cancel to finish recording.", vbOKCancel) = vbOK

    objFile.Close
    MsgBox "Mouse click positions saved to mouse_clicks.txt."
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

' Function to get cursor position using PowerShell script
Sub GetCursorPos(ByRef x, ByRef y)
    ' Execute PowerShell script to get mouse position
    Dim cmd, result
    cmd = "powershell -ExecutionPolicy Bypass -File .\GetMousePosition.ps1"
    
    ' Execute PowerShell command and capture output
    Set objExec = objShell.Exec(cmd)
    Do While Not objExec.StdOut.AtEndOfStream
        result = objExec.StdOut.ReadLine()
    Loop
    
    ' Parse coordinates from PowerShell output
    coords = Split(result, ",")
    x = CInt(coords(0))
    y = CInt(coords(1))
End Sub
