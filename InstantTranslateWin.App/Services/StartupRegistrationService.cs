using System.Diagnostics;
using Microsoft.Win32;

namespace InstantTranslateWin.App.Services;

public sealed class StartupRegistrationService
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string RunValueName = "InstantTranslateWin";

    public bool IsEnabled()
    {
        using var runKey = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
        var value = runKey?.GetValue(RunValueName) as string;
        return !string.IsNullOrWhiteSpace(value);
    }

    public void SetEnabled(bool enabled)
    {
        using var runKey = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true)
                           ?? throw new InvalidOperationException("Không mở được registry Run key.");

        if (!enabled)
        {
            runKey.DeleteValue(RunValueName, throwOnMissingValue: false);
            return;
        }

        var executablePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
        if (string.IsNullOrWhiteSpace(executablePath))
        {
            throw new InvalidOperationException("Không xác định được đường dẫn file thực thi.");
        }

        runKey.SetValue(RunValueName, $"\"{executablePath}\"");
    }
}
