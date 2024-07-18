Add-Type @"
using System;
using System.Runtime.InteropServices;

public class User32 {
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT {
        public int X;
        public int Y;
    }

    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT pt);
}
"@

$point = New-Object User32+POINT
[User32]::GetCursorPos([ref]$point) | Out-Null
"$($point.X),$($point.Y)"
