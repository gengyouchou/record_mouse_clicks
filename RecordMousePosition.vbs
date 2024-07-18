Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Ask user if they want to start recording mouse clicks
msg = "Click OK to start recording mouse clicks."
response = MsgBox(msg, vbOKOnly)
If response = vbOK Then
    ' Recording mode
    Set objFile = objFSO.CreateTextFile("mouse_clicks.txt", True)
    Dim count
    count = 0
    Do
        ' Wait for the user to press 's' key to record mouse position
        Do
            If objShell.AppActivate("Record Mouse Position") Then
                If objShell.SendKeys("{s}") Then
                    Exit Do
                End If
            End If
            WScript.Sleep 100
        Loop

        ' Get cursor position
        GetCursorPos x, y

        ' Write the coordinates to the output file
        objFile.WriteLine x & "," & y
        count = count + 1

        ' Display message that recording is done for this click
        MsgBox "Mouse click recorded at (" & x & ", " & y & "). Press 's' to record another click, or Cancel to finish.", vbOKCancel, "Mouse Click Recorder"
    Loop While MsgBox("Press 's' to record another mouse position. Click Cancel to finish recording.", vbOKCancel) = vbOK

    objFile.Close
    MsgBox count & " mouse click positions saved to mouse_clicks.txt."
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
