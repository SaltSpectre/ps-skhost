// This code defines a Mouse class that can be used to simulate mouse movements in a Windows environment.
// Based on https://github.com/ryanries/ImAlive

using System;
using System.Runtime.InteropServices;
public static class Mouse {
    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT {
    public int type;
    public MOUSEINPUT mi;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MOUSEINPUT {
    public int dx;
    public int dy;
    public uint mouseData;
    public uint dwFlags;
    public uint time;
    public IntPtr dwExtraInfo;
    }

    [DllImport("user32.dll")]
    public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    public static extern int GetSystemMetrics(int nIndex);

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT {
    public int x;
    public int y;
    }

    const int INPUT_MOUSE = 0;
    const int MOUSEEVENTF_MOVE = 0x0001;
    const int MOUSEEVENTF_ABSOLUTE = 0x8000;
    const int SM_CXSCREEN = 0;
    const int SM_CYSCREEN = 1;

    public static void MoveMouse() {
    INPUT Input = new INPUT();
    POINT CurrentPosition;

    GetCursorPos(out CurrentPosition);

    Input.type = INPUT_MOUSE;
    Input.mi.dwFlags = MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE;

    Input.mi.dx = (int)(CurrentPosition.x * (65536.0f / GetSystemMetrics(SM_CXSCREEN)));
    Input.mi.dy = (int)(CurrentPosition.y * (65536.0f / GetSystemMetrics(SM_CYSCREEN)));

    SendInput(1, new INPUT[] { Input }, Marshal.SizeOf(typeof(INPUT)));
    }
}
