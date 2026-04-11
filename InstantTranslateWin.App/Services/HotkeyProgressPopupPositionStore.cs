using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows;

namespace InstantTranslateWin.App.Services;

public sealed class HotkeyProgressPopupPositionStore
{
    // Tách riêng file vị trí popup để việc lưu state UI nhỏ không ảnh hưởng settings/history chính.
    private const string PositionFileName = "hotkey-progress-popup-position.json";
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    private readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly string _appDir;
    private readonly string _positionFilePath;

    public HotkeyProgressPopupPositionStore()
    {
        _appDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "InstantTranslateWin"
        );
        Directory.CreateDirectory(_appDir);
        _positionFilePath = Path.Combine(_appDir, PositionFileName);
    }

    public bool TryLoad(out Point savedPosition)
    {
        savedPosition = default;

        if (!File.Exists(_positionFilePath))
        {
            return false;
        }

        try
        {
            var json = File.ReadAllText(_positionFilePath, Utf8NoBom);
            if (string.IsNullOrWhiteSpace(json))
            {
                return false;
            }

            var payload = JsonSerializer.Deserialize<PopupPositionPayload>(json, _jsonOptions);
            if (payload is null || double.IsNaN(payload.Left) || double.IsNaN(payload.Top))
            {
                return false;
            }

            savedPosition = new Point(payload.Left, payload.Top);
            return true;
        }
        catch (JsonException ex)
        {
            ErrorFileLogger.LogException("HotkeyProgressPopupPositionStore.TryLoad", ex);
            BackupCorruptFile();
            return false;
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("HotkeyProgressPopupPositionStore.TryLoad", ex);
            return false;
        }
    }

    public void Save(Point position)
    {
        try
        {
            var payload = new PopupPositionPayload
            {
                Left = position.X,
                Top = position.Y
            };
            WriteJsonAtomically(payload);
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("HotkeyProgressPopupPositionStore.Save", ex);
        }
    }

    public void Clear()
    {
        try
        {
            if (File.Exists(_positionFilePath))
            {
                File.Delete(_positionFilePath);
            }
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("HotkeyProgressPopupPositionStore.Clear", ex);
        }
    }

    private void WriteJsonAtomically(PopupPositionPayload payload)
    {
        var tempPath = Path.Combine(
            _appDir,
            $"{Path.GetFileName(_positionFilePath)}.{Guid.NewGuid():N}.tmp"
        );

        try
        {
            var json = JsonSerializer.Serialize(payload, _jsonOptions);
            File.WriteAllText(tempPath, json, Utf8NoBom);

            // Ghi qua file tạm rồi replace/move để tránh làm hỏng file nếu app bị tắt giữa lúc save.
            if (File.Exists(_positionFilePath))
            {
                File.Replace(tempPath, _positionFilePath, destinationBackupFileName: null, ignoreMetadataErrors: true);
            }
            else
            {
                File.Move(tempPath, _positionFilePath);
            }
        }
        finally
        {
            try
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
            catch
            {
                // Best-effort temp cleanup only.
            }
        }
    }

    private void BackupCorruptFile()
    {
        try
        {
            if (!File.Exists(_positionFilePath))
            {
                return;
            }

            var backupPath = Path.Combine(
                _appDir,
                $"hotkey-progress-popup-position.corrupt-{DateTime.Now:yyyyMMdd-HHmmssfff}.json"
            );
            // Giữ lại file lỗi để còn tra nguyên nhân, đồng thời cho app fallback về vị trí mặc định.
            File.Move(_positionFilePath, backupPath);
            ErrorFileLogger.LogMessage(
                "HotkeyProgressPopupPositionStore.BackupCorruptFile",
                $"Corrupt popup position file was backed up to: {backupPath}"
            );
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("HotkeyProgressPopupPositionStore.BackupCorruptFile", ex);
        }
    }

    private sealed class PopupPositionPayload
    {
        public double Left { get; set; }

        public double Top { get; set; }
    }
}
