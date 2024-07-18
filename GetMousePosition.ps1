Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class MousePosition {
        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(ref POINT lpPoint);
        public struct POINT {
            public int X;
            public int Y;
        }
        public static POINT GetCursorPosition() {
            POINT point = new POINT();
            GetCursorPos(ref point);
            return point;
        }
    }
"@

$pos = [MousePosition]::GetCursorPosition()
"$($pos.X),$($pos.Y)"
