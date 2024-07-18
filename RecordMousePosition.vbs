Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Ask user to start recording mouse clicks
msg = "Press 's' to record the current mouse position. Press 'Cancel' to stop."
response = MsgBox(msg, vbOKCancel, "Mouse Click Recorder")

If response = vbCancel Then
    WScript.Quit
End If

' Recording mode
Set objFile = objFSO.CreateTextFile("mouse_clicks.txt", True)
Dim count
count = 0
Do
    ' Prompt user to press 's' key to record mouse position
    response = MsgBox("Press 's' to record the current mouse position. Press 'Cancel' to stop.", vbOKCancel, "Mouse Click Recorder")
    If response = vbCancel Then
        Exit Do
    End If

    ' Wait for 's' key press
    Do
        WScript.Sleep 100
        If GetAsyncKeyState(&H53) Then ' &H53 is the virtual-key code for 's'
            Exit Do
        End If
    Loop

    ' Get cursor position
    GetCursorPos x, y

    ' Write the coordinates to the output file
    objFile.WriteLine x & "," & y
    count = count + 1

    ' Display message that recording is done for this click
    response = MsgBox("Mouse click recorded at (" & x & ", " & y & "). Press 's' to record another click, or 'Cancel' to finish.", vbOKCancel, "Mouse Click Recorder")
    If response = vbCancel Then
        Exit Do
    End If
Loop

objFile.Close
MsgBox count & " mouse click positions saved to mouse_clicks.txt.", vbInformation, "Mouse Click Recorder"

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

' Function to check if a specific key is pressed using GetAsyncKeyState API
Function GetAsyncKeyState(vKey)
    Dim objExec, result
    Set objExec = objShell.Exec("powershell -ExecutionPolicy Bypass -File .\GetKeyState.ps1 -vKey " & vKey)
    result = objExec.StdOut.ReadLine()
    If CInt(result) <> 0 Then
        GetAsyncKeyState = True
    Else
        GetAsyncKeyState = False
    End If
End Function
