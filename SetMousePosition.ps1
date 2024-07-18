param (
    [int]$x,
    [int]$y
)

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class User32 {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);
}
"@

[User32]::SetCursorPos($x, $y)
