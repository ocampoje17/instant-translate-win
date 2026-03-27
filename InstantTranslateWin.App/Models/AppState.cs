namespace InstantTranslateWin.App.Models;

public sealed class AppState
{
    public AppSettings Settings { get; set; } = new();

    public List<TranslationRecord> History { get; set; } = [];
}
