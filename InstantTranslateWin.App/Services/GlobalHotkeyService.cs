using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace InstantTranslateWin.App.Services;

public sealed class GlobalHotkeyService : IDisposable
{
    private const int WmHotkey = 0x0312;

    private readonly Window _window;
    private readonly int _hotkeyId;
    private readonly uint _modifiers;
    private readonly uint _virtualKeyCode;

    private HwndSource? _hwndSource;
    private bool _registered;

    public event EventHandler? HotkeyPressed;

    public GlobalHotkeyService(Window window, int hotkeyId, ModifierKeys modifiers, Key key)
    {
        _window = window;
        _hotkeyId = hotkeyId;
        _modifiers = ConvertModifierKeys(modifiers);
        _virtualKeyCode = (uint)KeyInterop.VirtualKeyFromKey(key);
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    public bool Register()
    {
        var handle = new WindowInteropHelper(_window).Handle;
        if (handle == IntPtr.Zero)
        {
            return false;
        }

        _hwndSource = HwndSource.FromHwnd(handle);
        _hwndSource?.AddHook(HwndHook);

        _registered = RegisterHotKey(handle, _hotkeyId, _modifiers, _virtualKeyCode);
        return _registered;
    }

    public void Dispose()
    {
        var handle = new WindowInteropHelper(_window).Handle;

        if (_registered && handle != IntPtr.Zero)
        {
            UnregisterHotKey(handle, _hotkeyId);
            _registered = false;
        }

        if (_hwndSource is not null)
        {
            _hwndSource.RemoveHook(HwndHook);
            _hwndSource = null;
        }
    }

    private IntPtr HwndHook(IntPtr hwnd, int message, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (message == WmHotkey && wParam.ToInt32() == _hotkeyId)
        {
            handled = true;
            HotkeyPressed?.Invoke(this, EventArgs.Empty);
        }

        return IntPtr.Zero;
    }

    private static uint ConvertModifierKeys(ModifierKeys modifiers)
    {
        var result = 0u;

        if (modifiers.HasFlag(ModifierKeys.Alt))
        {
            result |= 0x0001;
        }

        if (modifiers.HasFlag(ModifierKeys.Control))
        {
            result |= 0x0002;
        }

        if (modifiers.HasFlag(ModifierKeys.Shift))
        {
            result |= 0x0004;
        }

        if (modifiers.HasFlag(ModifierKeys.Windows))
        {
            result |= 0x0008;
        }

        return result;
    }
}
