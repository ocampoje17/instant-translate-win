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
    private const uint KeyEventScanCode = 0x0008;
    private const uint MapVirtualKeyToScanCode = 0;
    private const uint WindowMessageCopy = 0x0301;
    private const int ClipboardBadDataHResult = unchecked((int)0x800401D3);
    private const int ModifierReleasePollAttempts = 12;
    private const int ModifierReleasePollDelayMs = 10;
    private const int PostHotkeySettlingDelayMs = 12;
    private const int CopyPollAttempts = 12;
    private const int CopyPollDelayMs = 18;
    private const int CopyFallbackDelayMs = 24;
    // WM_COPY phải fail-fast để không kéo dài cảm giác đơ popup/app khi app đích phản hồi chậm.
    private const int WindowMessageCopyPollAttempts = 6;
    private const int WindowMessageCopyPollDelayMs = 14;
    private const int WindowMessageCopyDelayMs = 20;

    [DllImport("user32.dll")]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);

    [DllImport("user32.dll")]
    private static extern uint GetClipboardSequenceNumber();

    [DllImport("user32.dll")]
    private static extern uint MapVirtualKey(uint uCode, uint uMapType);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct GUITHREADINFO
    {
        public int cbSize;
        public int flags;
        public IntPtr hwndActive;
        public IntPtr hwndFocus;
        public IntPtr hwndCapture;
        public IntPtr hwndMenuOwner;
        public IntPtr hwndMoveSize;
        public IntPtr hwndCaret;
        public RECT rcCaret;
    }

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

    public async Task<string?> CaptureSelectedTextAsync(Action<string>? progressCallback = null)
    {
        // Capture fallback chain (in order):
        // 1) UI Automation direct selection read
        // 2) Clipboard copy via Ctrl+C
        // 3) Clipboard copy via Ctrl+Insert
        // 4) Clipboard copy via scan-code based Ctrl+C (Electron compatibility)
        // 5) WM_COPY to foreground window
        // 6) UI Automation second pass
        // Bước nhanh: thử UI Automation trước để tránh đụng clipboard khi có thể.
        ReportProgress(progressCallback, "Đang thử đọc text bằng UI Automation...");
        await WaitForModifierKeysReleaseAsync();
        await Task.Delay(PostHotkeySettlingDelayMs);
        var automationCapturedText = TryCaptureFromUiAutomation();
        if (!string.IsNullOrWhiteSpace(automationCapturedText))
        {
            ReportProgress(progressCallback, "Đã đọc được text bằng UI Automation.");
            return automationCapturedText;
        }

        IDataObject? backupClipboard = null;
        // Không dùng sentinel text để tránh rác trong Clipboard History.
        // Mọi fallback copy sẽ dựa vào clipboard sequence number thay đổi sau từng attempt.
        var baselineSequence = 0u;
        CaptureResult result;

        try
        {
            backupClipboard = Clipboard.GetDataObject();
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("SelectedTextCaptureService.CaptureSelectedTextAsync.GetClipboardBackup", ex);
        }

        // Sequence gốc trước attempt đầu tiên (Ctrl+C).
        baselineSequence = GetClipboardSequenceNumber();

        ReportProgress(progressCallback, "UI Automation chưa có kết quả. Đang thử Ctrl+C...");
        result = await TryCaptureByCopyAsync(
            useInsertFallback: false,
            sequenceBeforeAction: baselineSequence
        );
        if (string.IsNullOrWhiteSpace(result.Text))
        {
            ErrorFileLogger.LogMessage(
                "SelectedTextCaptureService.CaptureSelectedTextAsync",
                "Ctrl+C fallback did not capture any text."
            );

            // Một số app không phản hồi tốt với Ctrl+C khi focus thay đổi nhanh,
            // nên thử lại bằng Ctrl+Insert.
            ReportProgress(progressCallback, "Ctrl+C chưa có kết quả. Đang thử Ctrl+Insert...");
            await Task.Delay(CopyFallbackDelayMs);
            // Mỗi fallback lấy lại sequence mới nhất để so sánh đúng theo từng bước.
            result = await TryCaptureByCopyAsync(
                useInsertFallback: true,
                sequenceBeforeAction: GetClipboardSequenceNumber()
            );
        }

        if (string.IsNullOrWhiteSpace(result.Text))
        {
            ErrorFileLogger.LogMessage(
                "SelectedTextCaptureService.CaptureSelectedTextAsync",
                "Ctrl+Insert fallback did not capture any text."
            );

            // Electron-based apps may ignore VK-based synthesized shortcuts.
            // Try Ctrl+C again but with scan-code based input.
            ReportProgress(progressCallback, "Ctrl+Insert chưa có kết quả. Đang thử ScanCode Ctrl+C...");
            result = await TryCaptureByCopyAsync(
                useInsertFallback: false,
                useScanCode: true,
                sequenceBeforeAction: GetClipboardSequenceNumber()
            );
        }

        // Chụp sequence trước WM_COPY để nhận biết dữ liệu mới do WM_COPY tạo ra.
        var sequenceBeforeWmCopy = GetClipboardSequenceNumber();
        if (string.IsNullOrWhiteSpace(result.Text) && TryRequestCopyFromForegroundWindow())
        {
            ErrorFileLogger.LogMessage(
                "SelectedTextCaptureService.CaptureSelectedTextAsync",
                "Scan-code fallback did not capture text. Triggering WM_COPY on foreground window."
            );

            // Final fallback: yêu cầu app foreground copy bằng WM_COPY rồi poll clipboard ngắn.
            // WM_COPY ở đây dùng PostMessage (non-blocking), không dùng SendMessageTimeout để tránh treo luồng UI.
            ReportProgress(progressCallback, "ScanCode Ctrl+C chưa có kết quả. Đang thử WM_COPY...");
            await Task.Delay(WindowMessageCopyDelayMs);
            result = await TryCaptureByCopyAsync(
                useInsertFallback: false,
                triggerCopyShortcut: false,
                pollAttempts: WindowMessageCopyPollAttempts,
                pollDelayMs: WindowMessageCopyPollDelayMs,
                sequenceBeforeAction: sequenceBeforeWmCopy
            );
        }
        else if (string.IsNullOrWhiteSpace(result.Text))
        {
            ErrorFileLogger.LogMessage(
                "SelectedTextCaptureService.CaptureSelectedTextAsync",
                "Could not trigger WM_COPY on foreground window."
            );
        }

        if (string.IsNullOrWhiteSpace(result.Text))
        {
            ReportProgress(progressCallback, "Đang thử lại UI Automation lần cuối...");
        }

        automationCapturedText = string.IsNullOrWhiteSpace(result.Text)
            ? TryCaptureFromUiAutomation()
            : automationCapturedText;

        if (string.IsNullOrWhiteSpace(result.Text) && string.IsNullOrWhiteSpace(automationCapturedText))
        {
            ReportProgress(progressCallback, "Không đọc được text đang chọn.");
            ErrorFileLogger.LogMessage(
                "SelectedTextCaptureService.CaptureSelectedTextAsync",
                "All capture fallbacks failed: Ctrl+C, Ctrl+Insert, scan-code Ctrl+C, WM_COPY, UI Automation."
            );
        }
        else
        {
            ReportProgress(progressCallback, "Đã đọc được text đang chọn.");
        }

        // Chỉ restore clipboard khi chắc chắn clipboard hiện tại vẫn là dữ liệu do lần capture này tạo ra.
        RestoreClipboardIfUnchanged(backupClipboard, result.ClipboardSequence);
        return !string.IsNullOrWhiteSpace(result.Text) ? result.Text : automationCapturedText;
    }

    private static async Task<CaptureResult> TryCaptureByCopyAsync(
        bool useInsertFallback,
        bool useScanCode = false,
        bool triggerCopyShortcut = true,
        int pollAttempts = CopyPollAttempts,
        int pollDelayMs = CopyPollDelayMs,
        uint? sequenceBeforeAction = null
    )
    {
        // Nếu caller chưa truyền sequence mốc, dùng sequence hiện tại ngay trước action copy.
        var baselineSequence = sequenceBeforeAction ?? GetClipboardSequenceNumber();

        // Mỗi lần thử sẽ gửi tổ hợp copy riêng (nếu được yêu cầu) và chờ clipboard đổi sequence.
        if (triggerCopyShortcut)
        {
            SimulateCopyShortcut(useInsertFallback, useScanCode);
        }

        var sequenceAtCapture = 0u;
        for (var i = 0; i < pollAttempts; i++)
        {
            if (i > 0)
            {
                await Task.Delay(pollDelayMs);
            }

            try
            {
                var currentSequence = GetClipboardSequenceNumber();
                if (currentSequence == baselineSequence)
                {
                    // Clipboard chưa đổi kể từ trước action => chưa có dữ liệu mới để đọc.
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

    private static void SimulateCopyShortcut(bool useInsertFallback, bool useScanCode = false)
    {
        var copyKey = useInsertFallback ? VirtualKeyInsert : VirtualKeyC;
        var inputs = new List<INPUT>(8)
        {
            // Ép nhả modifier dễ gây nhiễu trước khi gửi Ctrl+<CopyKey>.
            CreateKeyInput(VirtualKeyShift, keyUp: true),
            CreateKeyInput(VirtualKeyAlt, keyUp: true),
            CreateKeyInput(VirtualKeyLeftWindows, keyUp: true),
            CreateKeyInput(VirtualKeyRightWindows, keyUp: true)
        };

        // Prefer scan-code for Ctrl and copy key as fallback for Electron-like input handling.
        inputs.Add(CreateKeyInput(VirtualKeyControl, keyUp: false, useScanCode));
        inputs.Add(CreateKeyInput(copyKey, keyUp: false, useScanCode));
        inputs.Add(CreateKeyInput(copyKey, keyUp: true, useScanCode));
        inputs.Add(CreateKeyInput(VirtualKeyControl, keyUp: true, useScanCode));

        var inputArray = inputs.ToArray();
        _ = SendInput((uint)inputArray.Length, inputArray, Marshal.SizeOf<INPUT>());
    }

    private static INPUT CreateKeyInput(int virtualKey, bool keyUp, bool useScanCode = false)
    {
        var flags = keyUp ? KeyEventKeyUp : 0;
        var scanCode = (ushort)0;
        var virtualKeyCode = (ushort)virtualKey;

        if (useScanCode)
        {
            flags |= KeyEventScanCode;
            scanCode = (ushort)MapVirtualKey((uint)virtualKeyCode, MapVirtualKeyToScanCode);
            virtualKeyCode = 0;
        }

        return new INPUT
        {
            Type = InputKeyboard,
            Union = new InputUnion
            {
                KeyboardInput = new KEYBDINPUT
                {
                    VirtualKey = virtualKeyCode,
                    ScanCode = scanCode,
                    Flags = flags,
                    Time = 0,
                    ExtraInfo = 0
                }
            }
        };
    }

    private static bool TryRequestCopyFromForegroundWindow()
    {
        try
        {
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero)
            {
                ErrorFileLogger.LogMessage(
                    "SelectedTextCaptureService.TryRequestCopyFromForegroundWindow",
                    "Foreground window handle is null."
                );
                return false;
            }

            var targetWindow = ResolveFocusedWindowInForegroundThread(foregroundWindow);
            // Gửi vào cả hwnd focus và foreground window để tăng khả năng trúng control đang chọn text.
            var targetPosted = PostCopyMessage(targetWindow);
            var foregroundPosted = targetWindow == foregroundWindow || PostCopyMessage(foregroundWindow);
            var sent = targetPosted || foregroundPosted;

            if (!sent)
            {
                ErrorFileLogger.LogMessage(
                    "SelectedTextCaptureService.TryRequestCopyFromForegroundWindow",
                    "WM_COPY post message did not succeed for focused/foreground window."
                );
            }

            return sent;
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("SelectedTextCaptureService.TryRequestCopyFromForegroundWindow", ex);
            return false;
        }
    }

    private static IntPtr ResolveFocusedWindowInForegroundThread(IntPtr foregroundWindow)
    {
        if (foregroundWindow == IntPtr.Zero)
        {
            return IntPtr.Zero;
        }

        var foregroundThreadId = GetWindowThreadProcessId(foregroundWindow, out _);
        if (foregroundThreadId == 0)
        {
            return foregroundWindow;
        }

        try
        {
            var guiThreadInfo = new GUITHREADINFO
            {
                cbSize = Marshal.SizeOf<GUITHREADINFO>()
            };

            if (!GetGUIThreadInfo(foregroundThreadId, ref guiThreadInfo))
            {
                return foregroundWindow;
            }

            var focusedWindow = guiThreadInfo.hwndFocus;
            return focusedWindow != IntPtr.Zero ? focusedWindow : foregroundWindow;
        }
        catch
        {
            return foregroundWindow;
        }
    }

    private static bool PostCopyMessage(IntPtr windowHandle)
    {
        if (windowHandle == IntPtr.Zero)
        {
            return false;
        }

        // Non-blocking: chỉ enqueue WM_COPY vào message queue của cửa sổ đích.
        return PostMessage(windowHandle, WindowMessageCopy, IntPtr.Zero, IntPtr.Zero);
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

    private static void ReportProgress(Action<string>? progressCallback, string message)
    {
        try
        {
            progressCallback?.Invoke(message);
        }
        catch
        {
            // Progress callback is best-effort only.
        }
    }

    private static void RestoreClipboardIfUnchanged(
        IDataObject? backupClipboard,
        uint capturedSequence
    )
    {
        if (backupClipboard is null)
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
            }
            // Nếu sequence đã đổi, coi như user/app khác đã ghi clipboard mới -> không restore.
        }
        catch (COMException ex) when (ex.HResult == ClipboardBadDataHResult)
        {
            // Some apps (e.g., Electron) can temporarily expose invalid clipboard data.
            // Skip restore in this case because this path is best-effort only.
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("SelectedTextCaptureService.RestoreClipboardIfUnchanged", ex);
        }
    }
}
