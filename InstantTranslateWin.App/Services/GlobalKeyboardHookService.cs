using System.Runtime.InteropServices;
using System.Text;

namespace InstantTranslateWin.App.Services;

public sealed class GlobalKeyboardHookService : IDisposable
{
    private const int WhKeyboardLl = 13;
    private const int WmKeyDown = 0x0100;
    private const int WmSysKeyDown = 0x0104;
    private const int VirtualKeyLeftControl = 0xA2;
    private const int VirtualKeyRightControl = 0xA3;
    private const int VirtualKeyLeftShift = 0xA0;
    private const int VirtualKeyRightShift = 0xA1;
    private const int VirtualKeyLeftMenu = 0xA4;
    private const int VirtualKeyRightMenu = 0xA5;
    private const int VirtualKeyLeftWindows = 0x5B;
    private const int VirtualKeyRightWindows = 0x5C;

    private readonly LowLevelKeyboardProc _hookCallback;
    private IntPtr _hookHandle;

    public event EventHandler<GlobalKeyPressedEventArgs>? KeyPressed;

    [StructLayout(LayoutKind.Sequential)]
    private struct KbdLlHookStruct
    {
        public uint VirtualKeyCode;
        public uint ScanCode;
        public uint Flags;
        public uint Time;
        public nuint ExtraInfo;
    }

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int virtualKey);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetKeyboardState(byte[] lpKeyState);

    [DllImport("user32.dll")]
    private static extern IntPtr GetKeyboardLayout(uint idThread);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int ToUnicodeEx(
        uint wVirtKey,
        uint wScanCode,
        byte[] lpKeyState,
        [Out] StringBuilder pwszBuff,
        int cchBuff,
        uint wFlags,
        IntPtr dwhkl
    );

    public GlobalKeyboardHookService()
    {
        _hookCallback = HookProc;
    }

    public bool Start()
    {
        if (_hookHandle != IntPtr.Zero)
        {
            return true;
        }

        _hookHandle = SetWindowsHookEx(WhKeyboardLl, _hookCallback, IntPtr.Zero, 0);
        return _hookHandle != IntPtr.Zero;
    }

    public void Dispose()
    {
        if (_hookHandle == IntPtr.Zero)
        {
            return;
        }

        UnhookWindowsHookEx(_hookHandle);
        _hookHandle = IntPtr.Zero;
    }

    private IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            var message = unchecked((int)(long)wParam);
            if (message is WmKeyDown or WmSysKeyDown)
            {
                var hookData = Marshal.PtrToStructure<KbdLlHookStruct>(lParam);
                var eventArgs = BuildEventArgs(hookData);
                KeyPressed?.Invoke(this, eventArgs);
            }
        }

        return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
    }

    private static GlobalKeyPressedEventArgs BuildEventArgs(KbdLlHookStruct hookData)
    {
        var text = ResolveTypedText(hookData.VirtualKeyCode, hookData.ScanCode);

        return new GlobalKeyPressedEventArgs(
            (int)hookData.VirtualKeyCode,
            IsKeyDown(VirtualKeyLeftControl) || IsKeyDown(VirtualKeyRightControl),
            IsKeyDown(VirtualKeyLeftShift) || IsKeyDown(VirtualKeyRightShift),
            IsKeyDown(VirtualKeyLeftMenu) || IsKeyDown(VirtualKeyRightMenu),
            IsKeyDown(VirtualKeyLeftWindows) || IsKeyDown(VirtualKeyRightWindows),
            text
        );
    }

    private static string? ResolveTypedText(uint virtualKeyCode, uint scanCode)
    {
        var keyboardState = new byte[256];
        if (!GetKeyboardState(keyboardState))
        {
            return null;
        }

        var buffer = new StringBuilder(8);
        var layout = GetKeyboardLayout(0);
        var charCount = ToUnicodeEx(virtualKeyCode, scanCode, keyboardState, buffer, buffer.Capacity, 0, layout);
        if (charCount <= 0)
        {
            return null;
        }

        var text = buffer.ToString(0, charCount);
        return string.IsNullOrEmpty(text) ? null : text;
    }

    private static bool IsKeyDown(int virtualKey)
    {
        return (GetAsyncKeyState(virtualKey) & 0x8000) != 0;
    }
}

public sealed class GlobalKeyPressedEventArgs : EventArgs
{
    public int VirtualKeyCode { get; }
    public bool IsControlDown { get; }
    public bool IsShiftDown { get; }
    public bool IsAltDown { get; }
    public bool IsWindowsDown { get; }
    public string? TypedText { get; }

    public GlobalKeyPressedEventArgs(
        int virtualKeyCode,
        bool isControlDown,
        bool isShiftDown,
        bool isAltDown,
        bool isWindowsDown,
        string? typedText
    )
    {
        VirtualKeyCode = virtualKeyCode;
        IsControlDown = isControlDown;
        IsShiftDown = isShiftDown;
        IsAltDown = isAltDown;
        IsWindowsDown = isWindowsDown;
        TypedText = typedText;
    }
}
