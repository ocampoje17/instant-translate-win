using System.IO;
using System.Text;

namespace InstantTranslateWin.App.Services;

public static class ErrorFileLogger
{
    private const string AppFolderName = "InstantTranslateWin";
    private const string TextLogFileName = "runtime-errors.txt";
    private const string LegacyLogFileName = "runtime-errors.log";
    private const long MaxLogBytes = 1024 * 1024; // 1 MB
    private static readonly Encoding LogEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
    private static readonly object WriteLock = new();

    public static void LogException(string source, Exception exception)
    {
        if (exception is null)
        {
            LogMessage(source, "Exception is null.");
            return;
        }

        var now = DateTimeOffset.Now;
        var stackTrace = string.IsNullOrWhiteSpace(exception.StackTrace)
            ? "(no stack trace)"
            : exception.StackTrace;

        var builder = new StringBuilder(768);
        builder.AppendLine($"Date: {now:yyyy-MM-dd}");
        builder.AppendLine($"Time: {now:HH:mm:ss.fff zzz}");
        builder.AppendLine($"Source: {source}");
        builder.AppendLine($"ErrorType: {exception.GetType().FullName}");
        builder.AppendLine($"ErrorMessage: {exception.Message}");
        builder.AppendLine("StackTrace:");
        builder.AppendLine(stackTrace);
        builder.AppendLine("ExceptionDetails:");
        builder.AppendLine(exception.ToString());
        builder.AppendLine(new string('-', 80));
        builder.AppendLine();

        AppendEntry(builder.ToString());
    }

    public static void LogMessage(string source, string message)
    {
        var now = DateTimeOffset.Now;
        var safeMessage = string.IsNullOrWhiteSpace(message) ? "(empty)" : message;

        var builder = new StringBuilder(384);
        builder.AppendLine($"Date: {now:yyyy-MM-dd}");
        builder.AppendLine($"Time: {now:HH:mm:ss.fff zzz}");
        builder.AppendLine($"Source: {source}");
        builder.AppendLine("Message:");
        builder.AppendLine(safeMessage);
        builder.AppendLine(new string('-', 80));
        builder.AppendLine();

        AppendEntry(builder.ToString());
    }

    private static void AppendEntry(string entry)
    {
        try
        {
            var appDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                AppFolderName
            );
            Directory.CreateDirectory(appDir);

            var textLogPath = Path.Combine(appDir, TextLogFileName);
            var legacyLogPath = Path.Combine(appDir, LegacyLogFileName);

            lock (WriteLock)
            {
                CleanupLegacyLogFile(legacyLogPath);

                if (ShouldResetLogBeforeAppend(textLogPath, entry))
                {
                    File.WriteAllText(textLogPath, string.Empty, LogEncoding);
                }

                File.AppendAllText(textLogPath, entry, LogEncoding);
            }
        }
        catch
        {
            // Never throw from logging path.
        }
    }

    private static bool ShouldResetLogBeforeAppend(string logPath, string entry)
    {
        try
        {
            if (!File.Exists(logPath))
            {
                return false;
            }

            var existingBytes = new FileInfo(logPath).Length;
            var incomingBytes = LogEncoding.GetByteCount(entry);
            return existingBytes + incomingBytes > MaxLogBytes;
        }
        catch
        {
            return false;
        }
    }

    private static void CleanupLegacyLogFile(string legacyLogPath)
    {
        try
        {
            if (File.Exists(legacyLogPath))
            {
                File.Delete(legacyLogPath);
            }
        }
        catch
        {
            // Best-effort cleanup only.
        }
    }

    public static string GetTextLogPath()
    {
        var appDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            AppFolderName
        );
        return Path.Combine(appDir, TextLogFileName);
    }
}
