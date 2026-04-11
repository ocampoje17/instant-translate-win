using System.IO;
using System.Text;
using System.Text.Json;
using InstantTranslateWin.App.Models;

namespace InstantTranslateWin.App.Services;

public sealed class AppStateStore
{
    private const string LegacyStateFileName = "app-state.json";
    private const string SettingsFileName = "app-settings.json";
    private const string HistoryFileName = "translation-history.json";

    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    private readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly string _appDir;
    private readonly string _legacyStateFilePath;
    private readonly string _settingsFilePath;
    private readonly string _historyFilePath;

    public AppStateStore()
    {
        _appDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "InstantTranslateWin"
        );

        Directory.CreateDirectory(_appDir);
        _legacyStateFilePath = Path.Combine(_appDir, LegacyStateFileName);
        _settingsFilePath = Path.Combine(_appDir, SettingsFileName);
        _historyFilePath = Path.Combine(_appDir, HistoryFileName);
    }

    public async Task<AppState> LoadAsync()
    {
        try
        {
            if (File.Exists(_settingsFilePath) || File.Exists(_historyFilePath))
            {
                return await LoadSplitStateAsync();
            }

            if (File.Exists(_legacyStateFilePath))
            {
                return await LoadLegacyStateAsync();
            }
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("AppStateStore.LoadAsync", ex);
        }

        return CreateDefaultState();
    }

    public async Task SaveAsync(AppState state)
    {
        var normalizedState = NormalizeState(state);
        await WriteJsonAtomicallyAsync(_settingsFilePath, normalizedState.Settings);
        await WriteJsonAtomicallyAsync(_historyFilePath, normalizedState.History);
    }

    private async Task<AppState> LoadSplitStateAsync()
    {
        var settings = await LoadJsonFileAsync<AppSettings>(
            _settingsFilePath,
            static () => new AppSettings(),
            "AppStateStore.LoadSplitStateAsync.Settings"
        );
        var history = await LoadJsonFileAsync<List<TranslationRecord>>(
            _historyFilePath,
            static () => [],
            "AppStateStore.LoadSplitStateAsync.History"
        );

        return new AppState
        {
            Settings = settings,
            History = history
        };
    }

    private async Task<AppState> LoadLegacyStateAsync()
    {
        string legacyJson;
        try
        {
            legacyJson = await File.ReadAllTextAsync(_legacyStateFilePath, Utf8NoBom);
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("AppStateStore.LoadLegacyStateAsync.Read", ex);
            return CreateDefaultState();
        }

        if (string.IsNullOrWhiteSpace(legacyJson))
        {
            return CreateDefaultState();
        }

        try
        {
            var state = JsonSerializer.Deserialize<AppState>(legacyJson, _jsonOptions);
            var normalizedState = NormalizeState(state);
            await TryPersistMigratedLegacyStateAsync(normalizedState, "Loaded legacy combined state file.");
            return normalizedState;
        }
        catch (JsonException ex)
        {
            ErrorFileLogger.LogException("AppStateStore.LoadLegacyStateAsync.Deserialize", ex);

            var salvagedState = SalvageLegacyState(legacyJson);
            var backupPath = BackupCorruptFile(_legacyStateFilePath, "legacy-state");
            if (!string.IsNullOrEmpty(backupPath))
            {
                ErrorFileLogger.LogMessage(
                    "AppStateStore.LoadLegacyStateAsync.Backup",
                    $"Legacy state file was corrupt and has been backed up to: {backupPath}"
                );
            }

            await TryPersistMigratedLegacyStateAsync(
                salvagedState,
                "Recovered settings/history from a corrupt legacy combined state file."
            );
            return salvagedState;
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("AppStateStore.LoadLegacyStateAsync", ex);
            return CreateDefaultState();
        }
    }

    private async Task<T> LoadJsonFileAsync<T>(string path, Func<T> defaultFactory, string logSource)
    {
        if (!File.Exists(path))
        {
            return defaultFactory();
        }

        try
        {
            var json = await File.ReadAllTextAsync(path, Utf8NoBom);
            if (string.IsNullOrWhiteSpace(json))
            {
                return defaultFactory();
            }

            var value = JsonSerializer.Deserialize<T>(json, _jsonOptions);
            return value ?? defaultFactory();
        }
        catch (JsonException ex)
        {
            ErrorFileLogger.LogException(logSource, ex);
            var backupPath = BackupCorruptFile(path, Path.GetFileNameWithoutExtension(path));
            if (!string.IsNullOrEmpty(backupPath))
            {
                ErrorFileLogger.LogMessage(
                    $"{logSource}.Backup",
                    $"Corrupt JSON file was backed up to: {backupPath}"
                );
            }

            return defaultFactory();
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException(logSource, ex);
            return defaultFactory();
        }
    }

    private async Task WriteJsonAtomicallyAsync<T>(string targetPath, T value)
    {
        Directory.CreateDirectory(_appDir);

        var tempPath = Path.Combine(
            _appDir,
            $"{Path.GetFileName(targetPath)}.{Guid.NewGuid():N}.tmp"
        );

        try
        {
            var json = JsonSerializer.Serialize(value, _jsonOptions);
            await File.WriteAllTextAsync(tempPath, json, Utf8NoBom);

            if (File.Exists(targetPath))
            {
                File.Replace(tempPath, targetPath, destinationBackupFileName: null, ignoreMetadataErrors: true);
            }
            else
            {
                File.Move(tempPath, targetPath);
            }
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("AppStateStore.WriteJsonAtomicallyAsync", ex);
            throw;
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
                // Temp cleanup is best-effort only.
            }
        }
    }

    private AppState SalvageLegacyState(string legacyJson)
    {
        var state = CreateDefaultState();

        if (TryDeserializeExtractedProperty(legacyJson, "settings", out AppSettings? settings) && settings is not null)
        {
            state.Settings = settings;
            ErrorFileLogger.LogMessage(
                "AppStateStore.SalvageLegacyState.Settings",
                "Recovered settings from corrupt legacy combined state file."
            );
        }

        if (TryDeserializeExtractedProperty(legacyJson, "history", out List<TranslationRecord>? history) && history is not null)
        {
            state.History = history;
            ErrorFileLogger.LogMessage(
                "AppStateStore.SalvageLegacyState.History",
                "Recovered history from corrupt legacy combined state file."
            );
        }

        return NormalizeState(state);
    }

    private bool TryDeserializeExtractedProperty<T>(string json, string propertyName, out T? value)
    {
        value = default;

        try
        {
            var bytes = Utf8NoBom.GetBytes(json);
            var reader = new Utf8JsonReader(bytes, new JsonReaderOptions
            {
                AllowTrailingCommas = true,
                CommentHandling = JsonCommentHandling.Skip
            });

            while (reader.Read())
            {
                if (reader.TokenType != JsonTokenType.PropertyName || !reader.ValueTextEquals(propertyName))
                {
                    continue;
                }

                if (!reader.Read())
                {
                    return false;
                }

                using var valueDocument = JsonDocument.ParseValue(ref reader);
                value = valueDocument.Deserialize<T>(_jsonOptions);
                return value is not null;
            }
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException($"AppStateStore.TryDeserializeExtractedProperty.{propertyName}", ex);
        }

        return false;
    }

    private async Task TryPersistMigratedLegacyStateAsync(AppState state, string reason)
    {
        try
        {
            await SaveAsync(state);
            ErrorFileLogger.LogMessage("AppStateStore.TryPersistMigratedLegacyStateAsync", reason);
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("AppStateStore.TryPersistMigratedLegacyStateAsync", ex);
        }
    }

    private string? BackupCorruptFile(string path, string backupLabel)
    {
        try
        {
            if (!File.Exists(path))
            {
                return null;
            }

            var extension = Path.GetExtension(path);
            var fileName = $"{backupLabel}.corrupt-{DateTime.Now:yyyyMMdd-HHmmssfff}{extension}";
            var backupPath = Path.Combine(_appDir, fileName);
            var suffix = 1;

            while (File.Exists(backupPath))
            {
                fileName = $"{backupLabel}.corrupt-{DateTime.Now:yyyyMMdd-HHmmssfff}-{suffix}{extension}";
                backupPath = Path.Combine(_appDir, fileName);
                suffix++;
            }

            File.Move(path, backupPath);
            return backupPath;
        }
        catch (Exception ex)
        {
            ErrorFileLogger.LogException("AppStateStore.BackupCorruptFile", ex);
            return null;
        }
    }

    private static AppState NormalizeState(AppState? state)
    {
        return new AppState
        {
            Settings = state?.Settings ?? new AppSettings(),
            History = state?.History ?? []
        };
    }

    private static AppState CreateDefaultState()
    {
        return new AppState
        {
            Settings = new AppSettings(),
            History = []
        };
    }
}
