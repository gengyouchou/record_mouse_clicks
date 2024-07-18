Add-Type @"
using System;
using System.Runtime.InteropServices;

public class User32 {
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

    public const uint MOUSEEVENTF_LEFTDOWN = 0x02;
    public const uint MOUSEEVENTF_LEFTUP = 0x04;
}
"@

[User32]::mouse_event([User32]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
Start-Sleep -Milliseconds 100
[User32]::mouse_event([User32]::MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
