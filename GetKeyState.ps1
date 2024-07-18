param (
    [int]$vKey
)

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class User32 {
    [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
    public static extern short GetAsyncKeyState(int vkey);
}
"@

$state = [User32]::GetAsyncKeyState($vKey)
$state
