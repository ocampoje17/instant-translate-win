using System.IO;
using System.Text.Json;
using InstantTranslateWin.App.Models;

namespace InstantTranslateWin.App.Services;

public sealed class AppStateStore
{
    private readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly string _stateFilePath;

    public AppStateStore()
    {
        var appDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "InstantTranslateWin"
        );

        Directory.CreateDirectory(appDir);
        _stateFilePath = Path.Combine(appDir, "app-state.json");
    }

    public async Task<AppState> LoadAsync()
    {
        if (!File.Exists(_stateFilePath))
        {
            return new AppState();
        }

        await using var stream = File.OpenRead(_stateFilePath);
        var state = await JsonSerializer.DeserializeAsync<AppState>(stream, _jsonOptions);
        return state ?? new AppState();
    }

    public async Task SaveAsync(AppState state)
    {
        await using var stream = File.Create(_stateFilePath);
        await JsonSerializer.SerializeAsync(stream, state, _jsonOptions);
    }
}
