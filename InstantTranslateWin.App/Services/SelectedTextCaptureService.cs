using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;

namespace InstantTranslateWin.App.Services;

public sealed class SelectedTextCaptureService
{
    private const int VirtualKeyShift = 0x10;
    private const int VirtualKeyControlState = 0x11;
    private const int VirtualKeyAlt = 0x12;
    private const int VirtualKeyLeftWindows = 0x5B;
    private const int VirtualKeyRightWindows = 0x5C;

    private const ushort VirtualKeyControl = 0x11;
    private const ushort VirtualKeyC = 0x43;
    private const ushort VirtualKeyInsert = 0x2D;

    private const uint InputKeyboard = 1;
    private const uint KeyEventKeyUp = 0x0002;
    private const int ModifierReleasePollAttempts = 12;
    private const int ModifierReleasePollDelayMs = 10;
    private const int PostHotkeySettlingDelayMs = 12;
    private const int CopyPollAttempts = 12;
    private const int CopyPollDelayMs = 18;
    private const int CopyFallbackDelayMs = 24;
    private const int ClipboardWriteRetryAttempts = 4;
    private const int ClipboardWriteRetryDelayMs = 12;

    [DllImport("user32.dll")]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);

    [DllImport("user32.dll")]
    private static extern uint GetClipboardSequenceNumber();

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint Type;
        public InputUnion Union;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public KEYBDINPUT KeyboardInput;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public ushort VirtualKey;
        public ushort ScanCode;
        public uint Flags;
        public uint Time;
        public nuint ExtraInfo;
    }

    private sealed record CaptureResult(string? Text, uint ClipboardSequence);

    public async Task<string?> CaptureSelectedTextAsync()
    {
        // Bước nhanh: thử UI Automation trước để tránh đụng clipboard khi có thể.
        await WaitForModifierKeysReleaseAsync();
        await Task.Delay(PostHotkeySettlingDelayMs);
        var automationCapturedText = TryCaptureFromUiAutomation();
        if (!string.IsNullOrWhiteSpace(automationCapturedText))
        {
            return automationCapturedText;
        }

        IDataObject? backupClipboard = null;
        // Sentinel giúp phân biệt "clipboard cũ" với dữ liệu vừa copy từ app đang focus.
        var sentinel = $"__instant_translate_clipboard_sentinel_{Guid.NewGuid():N}__";
        var sentinelApplied = false;
        var baselineSequence = 0u;
        CaptureResult result;

        try
        {
            backupClipboard = Clipboard.GetDataObject();
        }
        catch
        {
            // Ignore clipboard read exceptions.
        }

        baselineSequence = GetClipboardSequenceNumber();

        sentinelApplied = await TrySetClipboardTextWithRetryAsync(sentinel);

        result = await TryCaptureByCopyAsync(sentinel, sentinelApplied, baselineSequence, useInsertFallback: false);
        if (string.IsNullOrWhiteSpace(result.Text))
        {
            // Một số app không phản hồi tốt với Ctrl+C khi focus thay đổi nhanh,
            // nên thử lại bằng Ctrl+Insert.
            await Task.Delay(CopyFallbackDelayMs);
            result = await TryCaptureByCopyAsync(sentinel, sentinelApplied, baselineSequence, useInsertFallback: true);
        }

        automationCapturedText = string.IsNullOrWhiteSpace(result.Text)
            ? TryCaptureFromUiAutomation()
            : automationCapturedText;

        RestoreClipboardIfUnchanged(backupClipboard, sentinel, sentinelApplied, result.ClipboardSequence);
        return !string.IsNullOrWhiteSpace(result.Text) ? result.Text : automationCapturedText;
    }

    private static async Task<CaptureResult> TryCaptureByCopyAsync(
        string sentinel,
        bool sentinelApplied,
        uint baselineSequence,
        bool useInsertFallback
    )
    {
        // Mỗi lần thử sẽ gửi tổ hợp copy riêng và chờ clipboard đổi sequence.
        SimulateCopyShortcut(useInsertFallback);

        var sequenceAtCapture = 0u;
        for (var i = 0; i < CopyPollAttempts; i++)
        {
            if (i > 0)
            {
                await Task.Delay(CopyPollDelayMs);
            }

            try
            {
                var currentSequence = GetClipboardSequenceNumber();
                if (!sentinelApplied && currentSequence == baselineSequence)
                {
                    // Clipboard chưa đổi kể từ trước khi gửi copy.
                    continue;
                }

                if (!Clipboard.ContainsText())
                {
                    continue;
                }

                var text = TryReadClipboardText();
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                if (sentinelApplied && string.Equals(text, sentinel, StringComparison.Ordinal))
                {
                    // Vẫn còn sentinel -> app đích chưa copy dữ liệu mới.
                    continue;
                }

                sequenceAtCapture = currentSequence;
                return new CaptureResult(text.Trim(), sequenceAtCapture);
            }
            catch
            {
                // Keep polling until timeout.
            }
        }

        return new CaptureResult(null, sequenceAtCapture);
    }

    private static void SimulateCopyShortcut(bool useInsertFallback)
    {
        var copyKey = useInsertFallback ? VirtualKeyInsert : VirtualKeyC;
        var inputs = new[]
        {
            // Ép nhả modifier dễ gây nhiễu trước khi gửi Ctrl+<CopyKey>.
            CreateKeyInput(VirtualKeyShift, keyUp: true),
            CreateKeyInput(VirtualKeyAlt, keyUp: true),
            CreateKeyInput(VirtualKeyLeftWindows, keyUp: true),
            CreateKeyInput(VirtualKeyRightWindows, keyUp: true),
            CreateKeyInput(VirtualKeyControl, keyUp: false),
            CreateKeyInput(copyKey, keyUp: false),
            CreateKeyInput(copyKey, keyUp: true),
            CreateKeyInput(VirtualKeyControl, keyUp: true)
        };

        _ = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
    }

    private static INPUT CreateKeyInput(int virtualKey, bool keyUp)
    {
        return new INPUT
        {
            Type = InputKeyboard,
            Union = new InputUnion
            {
                KeyboardInput = new KEYBDINPUT
                {
                    VirtualKey = (ushort)virtualKey,
                    ScanCode = 0,
                    Flags = keyUp ? KeyEventKeyUp : 0,
                    Time = 0,
                    ExtraInfo = 0
                }
            }
        };
    }

    private static async Task WaitForModifierKeysReleaseAsync()
    {
        // Wait briefly so global hotkey keys are released before simulating Ctrl+C.
        for (var i = 0; i < ModifierReleasePollAttempts; i++)
        {
            if (!IsAnyModifierPressed())
            {
                return;
            }

            await Task.Delay(ModifierReleasePollDelayMs);
        }
    }

    private static async Task<bool> TrySetClipboardTextWithRetryAsync(string text)
    {
        for (var i = 0; i < ClipboardWriteRetryAttempts; i++)
        {
            try
            {
                Clipboard.SetText(text);
                return true;
            }
            catch
            {
                if (i < ClipboardWriteRetryAttempts - 1)
                {
                    await Task.Delay(ClipboardWriteRetryDelayMs);
                }
            }
        }

        return false;
    }

    private static string? TryReadClipboardText()
    {
        try
        {
            if (!Clipboard.ContainsText())
            {
                return null;
            }

            return Clipboard.GetText();
        }
        catch
        {
            return null;
        }
    }

    private static string? TryCaptureFromUiAutomation()
    {
        try
        {
            var focusedElement = AutomationElement.FocusedElement;
            if (focusedElement is null)
            {
                return null;
            }

            if (focusedElement.TryGetCurrentPattern(TextPattern.Pattern, out var textPatternObject) &&
                textPatternObject is TextPattern textPattern)
            {
                var selections = textPattern.GetSelection();
                if (selections.Length > 0)
                {
                    var selectedText = selections[0].GetText(-1)?.Trim();
                    if (!string.IsNullOrWhiteSpace(selectedText))
                    {
                        return selectedText;
                    }
                }
            }
        }
        catch
        {
            // Ignore UI automation fallback failures.
        }

        return null;
    }

    private static bool IsAnyModifierPressed()
    {
        return IsKeyDown(VirtualKeyShift) ||
               IsKeyDown(VirtualKeyControlState) ||
               IsKeyDown(VirtualKeyAlt) ||
               IsKeyDown(VirtualKeyLeftWindows) ||
               IsKeyDown(VirtualKeyRightWindows);
    }

    private static bool IsKeyDown(int virtualKey)
    {
        return (GetAsyncKeyState(virtualKey) & 0x8000) != 0;
    }

    private static void RestoreClipboardIfUnchanged(
        IDataObject? backupClipboard,
        string sentinel,
        bool sentinelApplied,
        uint capturedSequence
    )
    {
        if (backupClipboard is null || !sentinelApplied)
        {
            return;
        }

        try
        {
            // Avoid overriding user's newer clipboard content.
            var currentSequence = GetClipboardSequenceNumber();
            if (capturedSequence != 0 && currentSequence == capturedSequence)
            {
                // Chỉ restore khi clipboard vẫn là kết quả từ lần capture này.
                Clipboard.SetDataObject(backupClipboard, true);
                return;
            }

            if (Clipboard.ContainsText() &&
                string.Equals(Clipboard.GetText(), sentinel, StringComparison.Ordinal))
            {
                Clipboard.SetDataObject(backupClipboard, true);
            }
        }
        catch
        {
            // Ignore clipboard restore exceptions.
        }
    }
}
